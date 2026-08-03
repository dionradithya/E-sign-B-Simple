import { SPFI, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { ISignaturePlaceholder } from "../models/IEsignState";
import { LIST_TASKS, LIST_PROCESS } from "../constants";
import { getSP } from "../pnpjsConfig";
import { generatePDFHash } from "../utils/HashHelper";
import { uniqBy } from "@microsoft/sp-lodash-subset";
import { getSiteUrlFromPath } from "../utils";

/**
 * Result interface for file and metadata fetching
 */
export interface IEsignDataResult {
    fileUrl: string;
    placeholders: ISignaturePlaceholder[];
    serverRelativeUrl: string;
    processId: number;
    processTitle: string;
    hashHex?: string;
    firstAssignedUserId?: number;
    signerName?: string;
    isDownloadable?: boolean;
    isSecretaryOnly?: boolean;
}

/**
 * Access validation result
 */
export interface IAccessValidationResult {
    isValid: boolean;
    error?: {
        title: string;
        message: string;
    };
}

/**
 * Service for handling eSign document data operations
 */
export class EsignDataService {

    /**
     * Main method to fetch file and metadata for eSign process
     * 
     * @param context - Form Customizer context
     * @returns Promise resolving to file URL, placeholders, and metadata
     * @throws Error if item ID is missing, file not found, or access denied
     */
    public static async getFileAndMetadata(
        context: FormCustomizerContext
    ): Promise<IEsignDataResult> {
        const sp: SPFI = getSP(context);
        const listId = context.list.guid.toString();
        const itemId = context.itemId;

        if (!itemId) {
            throw new Error("Item ID is missing.");
        }

        // 1. Get Field Definition to know Target List ID for ProcessID lookup
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const fieldDef: any = await sp.web.lists
            .getById(listId)
            .fields.getByInternalNameOrTitle("ProcessID")
            .select("LookupList")();

        const targetListId = fieldDef.LookupList;

        // 2. Get Task Item -> Expand ProcessID (ID and Title) and AssignedTo
        const item = await sp.web.lists
            .getById(listId)
            .items.getById(itemId)
            .select(
                "ProcessID/Id",
                "ProcessID/Title",
                "Status",
                "AssignedTo/Id",
                "AssignedTo/Title",
                "AssignedTo/EMail",
                "Secretary/Id",    
                "Secretary/Title",
                "Secretary/EMail"
            )
            .expand("ProcessID", "AssignedTo", "Secretary")();

        // 3. Validate Access
        const currentUser = await sp.web.currentUser();
        const validationResult = this.validateTaskAccess(item, currentUser);

        
        let isAssigned = false;
        let isSecretary = false;

        // Check if current user is in AssignedTo
        if (Array.isArray(item.AssignedTo)) {
            isAssigned = item.AssignedTo.some(
                (u: any) => u.Id === currentUser.Id
            );
        } else if (item.AssignedTo) {
            isAssigned = item.AssignedTo.Id === currentUser.Id;
        }

        // Check if current user is in Secretary
        if (Array.isArray(item.Secretary)) {
            isSecretary = item.Secretary.some(
                (u: any) => u.Id === currentUser.Id
            );
        } else if (item.Secretary) {
            isSecretary = item.Secretary.Id === currentUser.Id;
        }

        const isSecretaryOnly = isSecretary && !isAssigned;

        if (!validationResult.isValid && validationResult.error) {
            const error = new Error(validationResult.error.message) as Error & {
                title?: string
            };
            error.title = validationResult.error.title;
            throw error;
        }

        // 4. Check ProcessID exists
        if (!item.ProcessID || !item.ProcessID.Id) {
            throw new Error("Linked Document (ProcessID) not found.");
        }

        // 5. Fetch Source Item from Target List to get ProcessData and FileRef0
        const sourceItem = await sp.web.lists
            .getById(targetListId)
            .items.getById(item.ProcessID.Id)
            .select("ProcessData", "FileRef0", "HashHex", "isDownloadable")();

        // 6. Parse Process Data & Placeholders
        let assigneeName = "";
        let firstAssignedUserId: number | undefined = undefined;

        if (Array.isArray(item.AssignedTo) && item.AssignedTo.length > 0) {
            assigneeName = item.AssignedTo[0].Title;
            firstAssignedUserId = item.AssignedTo[0].Id;
        } else if (item.AssignedTo) {
            assigneeName = item.AssignedTo.Title;
            firstAssignedUserId = item.AssignedTo.Id;
        }

        const extractedPlaceholders = this.parseProcessData(
            sourceItem.ProcessData,
            assigneeName
        );

        // 7. Validate File Reference
        const fileRef = sourceItem.FileRef0;

        if (!fileRef.toLowerCase().endsWith(".pdf")) {
            throw new Error("The linked item is not a PDF file.");
        }

        // 8. Fetch File Blob (Cross-Site Compatible)
        const blob = await this.fetchFileBlob(fileRef, context);
        const url = URL.createObjectURL(blob);

        // Parse isDownloadable: show download if not explicitly 'false'
        const isDownloadable = sourceItem.isDownloadable !== 'false';

        return {
            fileUrl: url,
            placeholders: extractedPlaceholders,
            serverRelativeUrl: fileRef,
            processId: item.ProcessID.Id,
            processTitle: item.ProcessID.Title,
            hashHex: sourceItem.HashHex,
            firstAssignedUserId: firstAssignedUserId,
            signerName: assigneeName,
            isDownloadable: isDownloadable,
            isSecretaryOnly: isSecretaryOnly,
        };
    }

    /**
     * Validates task access based on status and assignment
     * 
     * @param item - Task item from SharePoint
     * @param currentUser - Current user object
     * @returns Validation result with error details if invalid
     */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public static validateTaskAccess(item: any, currentUser: any): IAccessValidationResult {
        // Check 1: Assignment validation
        let isAssigned = false;
        let isSecretary = false;

        // Check if current user is in AssignedTo   
        if (Array.isArray(item.AssignedTo)) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            isAssigned = item.AssignedTo.some((u: any) => u.Id === currentUser.Id);
        } else if (item.AssignedTo) {
            isAssigned = item.AssignedTo.Id === currentUser.Id;
        }

        // Check if current user is in Secretary
        if (Array.isArray(item.Secretary)) {
            isSecretary = item.Secretary.some((u: any) => u.Id === currentUser.Id);
        } else if (item.Secretary) {
            isSecretary = item.Secretary.Id === currentUser.Id;
        }



        if (!isAssigned && !isSecretary) {
            return {
                isValid: false,
                error: {
                    title: "Access Denied",
                    message: "You are not authorized to review this document."
                }
            };
        }

        // Check 2: Status validation
        if (item.Status !== 'Pending') {
            let title = "Access Denied";
            let message = "This document cannot be accessed.";

            if (item.Status === 'Approved' || item.Status === 'Rejected') {
                title = "Action Already Taken";
                message = `This document has already been ${item.Status.toLowerCase()}.`;
            } else if (item.Status === 'Canceled') {
                title = "Cancelled";
                message = "This request has been cancelled by the requestor/system.";
            } else if (item.Status === 'Expired') {
                title = "Expired";
                message = "This request has expired.";
            }

            return {
                isValid: false,
                error: { title, message }
            };
        }

        return { isValid: true };
    }

    /**
     * Parses ProcessData JSON and extracts signature placeholders
     * 
     * @param processData - JSON string containing process signature data
     * @param assigneeName - Name of the assigned user to match
     * @returns Array of signature placeholders
     */
    public static parseProcessData(
        processData: string,
        assigneeName: string
    ): ISignaturePlaceholder[] {
        if (!processData) {
            return [];
        }

        try {
            const parsed = JSON.parse(processData);

            if (parsed.ProcessSignature && Array.isArray(parsed.ProcessSignature)) {
                // Relaxed matching: check if JSON name includes SP name or vice versa (case-insensitive)
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const matchingSignature = parsed.ProcessSignature.find((sig: any) => {
                    const jsonName = (sig.displayNameApprover || "").toLowerCase();
                    const spName = (assigneeName || "").toLowerCase();
                    return jsonName.includes(spName) || spName.includes(jsonName);
                });

                if (matchingSignature &&
                    matchingSignature.Placeholder &&
                    Array.isArray(matchingSignature.Placeholder)) {

                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    return matchingSignature.Placeholder.map((p: any) => ({
                        page: p.PageNo,
                        x: p.Left / 100, // Convert % to 0-1
                        y: p.Top / 100,
                        width: p.Width / 100,
                        height: p.Height / 100,
                        type: Number(p.PlaceType) === 2 ? 'initial' : 'signature',
                        approverName: p.displayName || matchingSignature.displayNameApprover,
                        includeName: String(p.CheckListName).toLowerCase() === 'true',
                        includeDate: String(p.ChecklistDate).toLowerCase() === 'true',
                        includeBadge: String(p.CheckListBadge).toLowerCase() === 'true',
                        badgeNumber: p.BadgeNumber || matchingSignature.badgeNumberApprover || ""
                    }));
                }
            }
        } catch (parseErr) {
            console.error("Error parsing ProcessData:", parseErr);
        }

        return [];
    }

    /**
     * Fetches file blob from SharePoint, handling cross-site scenarios
     * 
     * @param fileRef - Server relative URL of the file
     * @param context - Form Customizer context
     * @returns Promise resolving to file Blob
     */
    public static async fetchFileBlob(
        fileRef: string,
        context: FormCustomizerContext
    ): Promise<Blob> {
        // FileRef might be in another site (e.g., /sites/DMS/... or /gensys/SiteName/...)
        // Use helper function to determine correct site URL
        const siteUrl = getSiteUrlFromPath(fileRef);

        const targetWeb = Web(siteUrl).using(SPFx(context));
        const blob = await targetWeb.getFileByServerRelativePath(fileRef).getBlob();

        return blob;
    }

    /**
     * Fetches signature log data (Process info + History tasks)
     * 
     * @param fileRef - Server relative URL of the file
     * @param sp - Initialized SPFI object
     */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public static async getSignatureLogData(fileRef: string, sp: SPFI, context: any): Promise<any> {
        let currentHash = "";

        // 1. Generate hash of the current file content to find the matching process item
        try {
            // Handle cross-site file params using helper function
            const siteUrl = getSiteUrlFromPath(fileRef);
            const targetWeb = Web(siteUrl).using(SPFx(context));

            const blob = await targetWeb.getFileByServerRelativePath(fileRef).getBlob();
            currentHash = await generatePDFHash(blob);

        } catch (hashErr) {
            console.error("Error generating hash for log verification", hashErr);
            // If hash fails, we can only rely on FileRef, so we proceed but maybe mark hash match as uncertain?
            // For now, if hash generation fails, we might treat currentHash as empty string.
        }

        // 2. Get Process Item by FileRef
        const processItems = await sp.web.lists.getByTitle(LIST_PROCESS).items
            .select("Id", "Title", "Status", "Created", "Requestor/Title", "Requestor/EMail", "FileRef0", "HashHex")
            .expand("Requestor")
            .filter(`FileRef0 eq '${fileRef.replace(/'/g, "''")}'`)
            .orderBy("Created", false)(); // Fetch all matching revisions


        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        let processItem: any = null;
        let isHashMatch = true;
        let isFilenameChanged = false;

        // Logic Check
        if (processItems.length > 0) {
            // Case 1 & 2: Found by FileRef (Same FileRef)
            // processItems is already ordered by Created desc, so [0] is the latest
            processItem = processItems[0];

            if (currentHash && processItem.HashHex !== currentHash) {
                isHashMatch = false; // Case 2: Hash Mismatch -> Content Changed ("isi document berubah")
                console.warn("Hash Mismatch! The file content does not match the latest approval record.");
            } else {
                isHashMatch = true; // Case 1: Hash Match -> All good
            }
        } else if (currentHash) {
            // Case 3: Not found by FileRef, try finding by Hash (FileRef Changed -> "nama document berubah")
            const hashItems = await sp.web.lists.getByTitle(LIST_PROCESS).items
                .select("Id", "Title", "Status", "Created", "Requestor/Title", "Requestor/EMail", "FileRef0", "HashHex")
                .expand("Requestor")
                .filter(`HashHex eq '${currentHash}'`)
                .orderBy("Created", false)();

            if (hashItems.length > 0) {
                processItem = hashItems[0]; // Take the latest consistent with this content
                isFilenameChanged = true; // "nama document berubah"
                isHashMatch = true; // Content matches (since we searched by hash)
                console.warn("Filename Mismatch! The document content matches a record with a different name/path.");
            }
        }

        if (!processItem) {
            return null;
        }

        // 3. Get Related Tasks
        try {
            const tasks = await sp.web.lists.getByTitle(LIST_TASKS).items
                .select(
                    "Id",
                    "Title",
                    "Status",
                    "Created",
                    "Modified",
                    "AssignedTo/Title",
                    "Comments",
                    "ProcessID/Id",
                    "StepApprover",
                    "TotalApprover"
                )
                .expand("AssignedTo", "ProcessID")
                .filter(`ProcessID/Id eq ${processItem.Id}`)
                .orderBy("Created", true)();

            return {
                process: {
                    ...processItem,
                    OriginFilename: processItem.FileRef0 ? processItem.FileRef0.split('/').pop() : "Unknown.pdf"
                },
                isHashMatch: isHashMatch,
                isFilenameChanged: isFilenameChanged,
                tasks: tasks.map(t => {
                    let assignees = [];
                    if (Array.isArray(t.AssignedTo)) {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        assignees = t.AssignedTo.map((u: any) => ({ title: u.Title, id: u.Id }));
                    } else if (t.AssignedTo) {
                        assignees = [{ title: t.AssignedTo.Title, id: t.AssignedTo.Id }];
                    }

                    return {
                        assignedTo: assignees,
                        type: "Sign", // Default to Sign for now
                        step: t.StepApprover,
                        totalSteps: t.TotalApprover,

                        status: t.Status,
                        createdOn: t.Created,
                        completedOn: (t.Status === 'Approved' || t.Status === 'Rejected') ? t.Modified : null,
                        comments: t.Comments
                    };
                })
            };
        } catch (err) {
            console.error("Error fetching tasks for log:", err);
            // Return process info even if tasks fail
            return {
                process: {
                    ...processItem,
                    OriginFilename: processItem.FileRef0 ? processItem.FileRef0.split('/').pop() : "Unknown.pdf"
                },
                isHashMatch: isHashMatch,
                isFilenameChanged: isFilenameChanged,
                tasks: []
            };
        }
    }

    /**
     * Fetches a process by its unique Title (code)
     * @param uniqueCode - The unique code (Title) to search for
     * @param sp - Initialized SPFI object
     */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public static async getProcessByUniqueCode(uniqueCode: string, sp: SPFI): Promise<any> {
        const items = await sp.web.lists.getByTitle(LIST_PROCESS).items
            .select("Id", "Title", "FileRef0", "Status", "Created", "Requestor/Title")
            .expand("Requestor")
            .filter(`Title eq '${uniqueCode}'`)
            .top(1)();

        return items.length > 0 ? items[0] : null;
    }

    /**
     * Fetches only the HashHex for a given process code
     * @param uniqueCode - The unique code (Title) to search for
     * @param sp - Initialized SPFI object
     * @returns The HashHex string or null if not found
     */
    public static async getProcessHashByCode(uniqueCode: string, sp: SPFI): Promise<string | undefined> {
        const items = await sp.web.lists.getByTitle(LIST_PROCESS).items
            .select("HashHex")
            .filter(`Title eq '${uniqueCode}'`)
            .top(1)();

        return items.length > 0 ? items[0].HashHex : undefined;
    }

    /**
     * Fetches process items matching a given hash
     * @param hashHex - The hash to search for
     * @param sp - Initialized SPFI object
     * @returns The most recent process item matching the hash, or null
     */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public static async getProcessByHash(hashHex: string, sp: SPFI): Promise<any> {
        const items = await sp.web.lists.getByTitle(LIST_PROCESS).items
            .select("Id", "Title", "FileRef0", "Status", "Created", "Requestor/Title", "HashHex")
            .expand("Requestor")
            .filter(`HashHex eq '${hashHex}'`)
            .orderBy("Created", false)
            .top(1)();

        return items.length > 0 ? items[0] : null;
    }

    /**
     * Fetches all process items for a given file reference
     * @param fileRef - Server relative URL of the file
     * @param sp - Initialized SPFI object
     */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public static async getProcessesByFileRef(fileRef: string, sp: SPFI): Promise<any[]> {
        return sp.web.lists.getByTitle(LIST_PROCESS).items
            .select("Id", "Title", "Status", "Created", "Requestor/Title", "Requestor/EMail", "FileRef0", "HashHex", "ProcessData")
            .expand("Requestor")
            .filter(`FileRef0 eq '${fileRef.replace(/'/g, "''")}'`)
            .orderBy("Created", false)(); // Descending to see latest first
    }

    /**
     * Fetches all related process items matching either FileRef OR HashHex
     * @param fileRef - Server relative URL of the file
     * @param hashHex - Hash of the file content
     * @param sp - Initialized SPFI object
     */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public static async getRelatedProcesses(fileRef: string, hashHex: string, sp: SPFI): Promise<any[]> {
        const promises = [];

        // 1. Fetch by FileRef
        if (fileRef) {
            promises.push(
                sp.web.lists.getByTitle(LIST_PROCESS).items
                    .select("Id", "Title", "Status", "Created", "Requestor/Title", "Requestor/EMail", "FileRef0", "HashHex", "ProcessData")
                    .expand("Requestor")
                    .filter(`FileRef0 eq '${fileRef.replace(/'/g, "''")}'`)
                    .orderBy("Created", false)()
            );
        }

        // 2. Fetch by HashHex (if available)
        if (hashHex) {
            promises.push(
                sp.web.lists.getByTitle(LIST_PROCESS).items
                    .select("Id", "Title", "Status", "Created", "Requestor/Title", "Requestor/EMail", "FileRef0", "HashHex", "ProcessData")
                    .expand("Requestor")
                    .filter(`HashHex eq '${hashHex}'`)
                    .orderBy("Created", false)()
            );
        }

        const results = await Promise.all(promises);

        // Flatten and deduplicate
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const allItems = results.reduce((acc, curr) => acc.concat(curr), []) as any[];

        // Deduplicate by Id
        const uniqueItems = uniqBy(allItems, 'Id');

        // Sort by Created descending
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return uniqueItems.sort((a: any, b: any) => {
            return new Date(b.Created).getTime() - new Date(a.Created).getTime();
        });
    }

    /**
     * Fetches tasks for a specific process ID
     * @param processId - The ID of the process
     * @param sp - Initialized SPFI object
     */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public static async getTasksByProcessId(processId: number, sp: SPFI): Promise<any[]> {
        return sp.web.lists.getByTitle(LIST_TASKS).items
            .select(
                "Id",
                "Title",
                "Status",
                "Created",
                "Modified",
                "AssignedTo/Title",
                "Comments",
                "ProcessID/Id",
                "StepApprover",
                "TotalApprover"
            )
            .expand("AssignedTo", "ProcessID")
            .filter(`ProcessID/Id eq ${processId}`)
            .orderBy("Created", true)();
    }

    /**
     * Fetches pending tasks assigned to the current user
     * @param sp - Initialized SPFI object
     */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public static async getPendingTasksForCurrentUser(sp: SPFI): Promise<any[]> {
        const currentUser = await sp.web.currentUser();

        // 1. Fetch Tasks
        const tasks = await sp.web.lists.getByTitle(LIST_TASKS).items
            .select(
                "Id",
                "Title",
                "Status",
                "Created",
                "ProcessID/Id",
                "ProcessID/Title",
                "ProcessID/FileRef0",
                "ProcessID/isDownloadable",
                "ProcessID/GUID",
                "AssignedTo/Title",
                "AssignedTo/Id"

            )
            .expand("ProcessID", "AssignedTo")
            .filter(`Status eq 'Pending' and AssignedTo/Id eq ${currentUser.Id}`)
            .orderBy("Created", false)();

        if (tasks.length === 0) return [];

        // 2. Extract Unique Process IDs
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const processIds = [...new Set(tasks.map((t: any) => t.ProcessID ? t.ProcessID.Id : null).filter((id: any) => id !== null))];

        if (processIds.length === 0) return tasks;

        // 3. Fetch Process Details (Requestor)
        // Construct filter query: Id eq 1 or Id eq 2 ...
        // Batching might be better for many items, but for "My Pending Tasks" usually reasonable number.
        // We'll chunk if necessary or just fetch relevant ones. 
        // For simplicity/reliability, let's fetch by ID filter.

        const filterQuery = processIds.map(id => `Id eq ${id}`).join(' or ');

        const processes = await sp.web.lists.getByTitle(LIST_PROCESS).items
            .select("Id", "Requestor/Title", "FileRef0")
            .expand("Requestor")
            .filter(filterQuery)();

        // 4. Fetch actual File UniqueId for each process to support Embed.aspx
        const fileGuidsMap = new Map<number, string>();
        for (const p of processes) {
            if (p.FileRef0) {
                try {
                    // Fetch the file metadata using the file reference path
                    const siteUrl = getSiteUrlFromPath(p.FileRef0);
                    // Use PnPJS v3 isolated Web factory pattern pointing to the specific site
                    const targetWeb = Web([sp.web, siteUrl]);
                    const fileData = await targetWeb.getFileByServerRelativePath(p.FileRef0).select("UniqueId")();
                    fileGuidsMap.set(p.Id, fileData.UniqueId);
                } catch (err) {
                    console.warn(`Could not fetch UniqueId for file: ${p.FileRef0}`, err);
                }
            }
        }

        // 5. Map Process Info & File GUID back to Tasks
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const processMap = new Map(processes.map((p: any) => [p.Id, p]));

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return tasks.map((t: any) => {
            if (t.ProcessID && processMap.has(t.ProcessID.Id)) {
                return {
                    ...t,
                    ProcessID: {
                        ...t.ProcessID,
                        Requestor: processMap.get(t.ProcessID.Id).Requestor,
                        GUID: fileGuidsMap.get(t.ProcessID.Id) // Attach the real file GUID
                    }
                };
            }
            return t;
        });
    }

    /**
     * Fetches processing requests initiated by the current user
     * @param sp - Initialized SPFI object
     */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public static async getMyProcessingRequests(sp: SPFI): Promise<any[]> {
        const currentUser = await sp.web.currentUser();

        return sp.web.lists.getByTitle(LIST_PROCESS).items
            .select(
                "Id",
                "Title",
                "Status",
                "Created",
                "FileRef0",
                "Requestor/Title",
                "Requestor/Id"
            )
            .expand("Requestor")
            .filter(`Status eq 'Processing' and Requestor/Id eq ${currentUser.Id}`)
            .orderBy("Created", false)();
    }
}
