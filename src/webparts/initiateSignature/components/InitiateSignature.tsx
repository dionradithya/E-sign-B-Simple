import * as React from 'react';
import { useState, useEffect } from 'react';
import { IInitiateSignatureProps } from './IInitiateSignatureProps';
import { getSP } from '../../../common/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/files';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import PdfViewer, { ISignaturePlaceholder } from './PdfViewer';
import Sidebar, { IApprover } from './Sidebar';
import { TouchDragProvider } from './TouchDragContext';
import { ApprovalService, IApprovalMapItem } from '../../../common/services/ApprovalService';
import { generatePDFHash } from '../../../common/utils/HashHelper';
import { PDFDocument, rgb } from 'pdf-lib';
import { generateQrCodeImageBytes } from '../../../common/utils/QrCodeHelper';
import "bootstrap/dist/css/bootstrap.min.css";
import { sanitizeApproverName, getEncodedFolderUrl } from '../../../common/utils/helper';
import { Dialog, DialogType, DialogFooter, PrimaryButton } from '@fluentui/react';
import { LIST_PROCESS, LIST_ACTIVITY_LOG, TENANT_DOMAIN, SITES_ESIGN, SITE_REDIRECT } from '../../../common/constants';

/**
 * Parameters for createProcessForDocument — the reusable per-document creation routine shared by
 * the single-document save and the batch loop. When `batch` is omitted the created Approvals
 * Process item is identical to the original single-flow item (no BatchId/BatchMode). When `batch`
 * is supplied, every item in the batch shares one batchId and BatchMode='AllOrNothing'.
 */
interface ICreateProcessParams {
    fileRef: string;
    fileBlob: Blob;
    approvers: IApprover[];
    placeholders: ISignaturePlaceholder[];
    reviewerCanDownload: boolean;
    embedQrCode: boolean;
    currentUserId: number;
    spDoc: SPFI;
    spList: SPFI;
    batch?: { batchId: string; batchMode: string };
}

/**
 * A single document being configured in batch mode. Each descriptor keeps its own approvers and
 * placeholders so every document is configured independently before the all-or-nothing submit.
 */
interface IBatchDocument {
    fileRef: string;
    fileName: string;
    blob: Blob | undefined;
    fileUrl: string | undefined;
    approvers: IApprover[];
    placeholders: ISignaturePlaceholder[];
    reviewerCanDownload: boolean;
    embedQrCode: boolean;
    loaded: boolean;
}

/**
 * Generate a RFC-4122 v4 UUID. Uses crypto.randomUUID when available, otherwise a getRandomValues
 * fallback so the batch still gets a unique BatchId in older runtimes.
 */
const generateBatchId = (): string => {
    const c = (typeof crypto !== 'undefined') ? crypto : undefined;
    if (c && typeof c.randomUUID === 'function') {
        return c.randomUUID();
    }
    // Fallback UUID v4
    const bytes = new Uint8Array(16);
    if (c && typeof c.getRandomValues === 'function') {
        c.getRandomValues(bytes);
    } else {
        for (let i = 0; i < 16; i++) bytes[i] = Math.floor(Math.random() * 256);
    }
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    const hex = Array.from(bytes, b => b.toString(16).padStart(2, '0'));
    return `${hex[0]}${hex[1]}${hex[2]}${hex[3]}-${hex[4]}${hex[5]}-${hex[6]}${hex[7]}-${hex[8]}${hex[9]}-${hex[10]}${hex[11]}${hex[12]}${hex[13]}${hex[14]}${hex[15]}`;
};


const InitiateSignature: React.FC<IInitiateSignatureProps> = (props) => {
    const { context } = props;
    const [fileUrl, setFileUrl] = useState<string | null>(null);
    const [fileBlob, setFileBlob] = useState<Blob | null>(null);
    const [loading, setLoading] = useState<boolean>(false);
    const [approvers, setApprovers] = useState<IApprover[]>([]);
    const [approvalMap, setApprovalMap] = useState<IApprovalMapItem[]>([]);
    const [placeholders, setPlaceholders] = useState<ISignaturePlaceholder[]>([]);
    const [placementMode, setPlacementMode] = useState<'current' | 'all' | 'range'>('current');
    const [numPages, setNumPages] = useState<number>(0);
    const [visiblePage, setVisiblePage] = useState<number>(1); // Page the user is looking at
    const [isSaving, setIsSaving] = useState<boolean>(false);
    const [validationTriggered, setValidationTriggered] = useState<boolean>(false);
    const [hideDialog, setHideDialog] = useState<boolean>(true);
    const [hideErrorDialog, setHideErrorDialog] = useState<boolean>(true);
    const [isPendingDialogVisible, setIsPendingDialogVisible] = useState<boolean>(false);
    const [reviewerCanDownload, setReviewerCanDownload] = useState<boolean>(true);
    const [embedQrCode, setEmbedQrCode] = useState<boolean>(true);

    // Metadata for the current file
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const [fileMetadata, setFileMetadata] = useState<any>(null);

    const queryParams = new URLSearchParams(window.location.search);
    const fileRef = queryParams.get('fileRef');
    const sourcePage = queryParams.get('source');

    // Batch mode detection (FEAT-002): the batch command redirects with a `fileRefs` query param
    // holding a URI-encoded JSON array of server-relative URLs (see FEAT-001 findings). If present
    // and it parses to an array of >=1 strings we enter batch mode; otherwise the single-file path
    // below behaves exactly as before.
    const batchFileRefs = React.useMemo<string[] | null>(() => {
        const raw = queryParams.get('fileRefs');
        if (!raw) return null;
        try {
            const parsed = JSON.parse(raw);
            if (Array.isArray(parsed) && parsed.length >= 1 && parsed.every(r => typeof r === 'string')) {
                return parsed as string[];
            }
        } catch (e) {
            console.error('[InitiateSignature] Failed to parse fileRefs query param', e);
        }
        return null;
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);
    const isBatchMode = !!batchFileRefs;

    // Per-document batch state: each document keeps its own approvers/placeholders.
    const [batchDocs, setBatchDocs] = useState<IBatchDocument[]>([]);
    const [activeDocIndex, setActiveDocIndex] = useState<number>(0);
    const [batchValidationMessage, setBatchValidationMessage] = useState<string | null>(null);
    const [batchDialogVisible, setBatchDialogVisible] = useState<boolean>(false);

    // Single-document fetch effect. Skipped entirely in batch mode so there is zero behavior
    // change for the single flow (and so the no-fileRef redirect does not fire during a batch).
    useEffect(() => {
        if (isBatchMode) {
            return;
        }
        if (!fileRef) {
            window.location.href = `${TENANT_DOMAIN}/${SITE_REDIRECT}`;
            return;
        }

        const fetchFile = async (): Promise<void> => {
            setLoading(true);
            try {
                // Use the sourcePage (from query param) if available, otherwise default to current context
                const sp: SPFI = getSP(context, sourcePage || undefined);

                // Check for pending process
                const esignSiteUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}`;
                const spEsign: SPFI = getSP(context, esignSiteUrl);

                // Escape single quotes for OData filter
                const safeFileRef = (fileRef as string).replace(/'/g, "''");

                const pendingItems = await spEsign.web.lists.getByTitle(LIST_PROCESS).items.filter(`FileRef0 eq '${safeFileRef}' and Status eq 'Processing'`)();

                if (pendingItems.length > 0) {
                    setIsPendingDialogVisible(true);
                    setLoading(false);
                    return;
                }

                // 1. Get File Item to retrieve metadata (ID, ListId, etc)
                // We need to resolve the file item associated with the path
                const fileObj = sp.web.getFileByServerRelativePath(fileRef);
                const item = await fileObj.getItem();
                const itemFields = await item(); // get fields

                // Also get List Info
                // We can get list from the item parent context usually, or deduce it
                // A robust way works if we know the list name, but here we might just store what we have
                // For simplicity, we assume we can get basic info.

                // Let's get the Web URL and List Info more generically if possible
                // But for "Parameter" JSON, we need specific fields.
                const webUrl = context.pageContext.web.absoluteUrl;

                setFileMetadata({
                    baseUrl: context.pageContext.site.absoluteUrl, // or web url?
                    web: webUrl,
                    listId: "Unknown", // Would need extra call to parent list
                    folder: fileRef.substring(0, fileRef.lastIndexOf('/')),
                    listName: "Unknown",
                    fileName: fileRef.substring(fileRef.lastIndexOf('/') + 1),
                    itemId: itemFields.Id,
                    origin: window.location.origin,
                    fileRef: fileRef
                });

                // 2. Get Blob for Viewer
                const blob = await fileObj.getBlob();
                console.log("[InitiateSignature] PDF Blob Fetched Successfully:", blob.size, "bytes");
                setFileBlob(blob);
                const url = URL.createObjectURL(blob);
                setFileUrl(url);

            } catch (err) {
                console.error("Error fetching file:", err);
            } finally {
                setLoading(false);
            }
        };

        fetchFile().catch(console.error);

        return () => {
            if (fileUrl) URL.revokeObjectURL(fileUrl);
        };
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [fileRef, context, isBatchMode]);

    // Fetch Approval Map
    useEffect(() => {
        const fetchMap = async (): Promise<void> => {
            const esignSiteUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}`;
            const sp = getSP(context, esignSiteUrl);
            const service = new ApprovalService(sp);
            const mapItems = await service.getApprovalMap();
            setApprovalMap(mapItems);
        };

        fetchMap().catch(console.error);
    }, [context]);

    // Batch mode: load every selected PDF once and seed per-document state. Reuses the same
    // cross-site fetch logic as the single flow (getSP against the source site + getFileByServerRelativePath).
    useEffect(() => {
        if (!isBatchMode || !batchFileRefs) {
            return;
        }

        const loadBatch = async (): Promise<void> => {
            setLoading(true);
            try {
                const sp: SPFI = getSP(context, sourcePage || undefined);

                const docs: IBatchDocument[] = await Promise.all(batchFileRefs.map(async (ref): Promise<IBatchDocument> => {
                    const fileName = ref.substring(ref.lastIndexOf('/') + 1);
                    try {
                        const fileObj = sp.web.getFileByServerRelativePath(ref);
                        const blob = await fileObj.getBlob();
                        const url = URL.createObjectURL(blob);
                        return {
                            fileRef: ref,
                            fileName,
                            blob: blob,
                            fileUrl: url,
                            approvers: [],
                            placeholders: [],
                            reviewerCanDownload: true,
                            embedQrCode: true,
                            loaded: true
                        };
                    } catch (docErr) {
                        console.error(`[InitiateSignature] Failed to load batch document ${ref}`, docErr);
                        return {
                            fileRef: ref,
                            fileName,
                            blob: undefined,
                            fileUrl: undefined,
                            approvers: [],
                            placeholders: [],
                            reviewerCanDownload: true,
                            embedQrCode: true,
                            loaded: false
                        };
                    }
                }));

                setBatchDocs(docs);
                setActiveDocIndex(0);
            } catch (err) {
                console.error('[InitiateSignature] Error loading batch documents', err);
            } finally {
                setLoading(false);
            }
        };

        loadBatch().catch(console.error);

        return () => {
            // Revoke any object URLs created for the batch documents.
            setBatchDocs(prev => {
                prev.forEach(d => { if (d.fileUrl) URL.revokeObjectURL(d.fileUrl); });
                return prev;
            });
        };
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [isBatchMode, context]);

    // Batch mode: mirror the active document into the shared editing state (approvers,
    // placeholders, blob, metadata, toggles) so the existing Sidebar + PdfViewer render and
    // mutate the currently selected document without any prop changes.
    useEffect(() => {
        if (!isBatchMode) {
            return;
        }
        const doc = batchDocs[activeDocIndex];
        if (!doc) {
            return;
        }
        setApprovers(doc.approvers);
        setPlaceholders(doc.placeholders);
        setReviewerCanDownload(doc.reviewerCanDownload);
        setEmbedQrCode(doc.embedQrCode);
        setFileBlob(doc.blob || null);
        setFileUrl(doc.fileUrl || null);
        setValidationTriggered(false);
        setFileMetadata(doc.blob ? {
            fileName: doc.fileName,
            fileRef: doc.fileRef,
            folder: doc.fileRef.substring(0, doc.fileRef.lastIndexOf('/')),
            origin: window.location.origin
        } : null);
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [isBatchMode, activeDocIndex, batchDocs.length]);

    // Batch mode: write the shared editing state back into the active document descriptor so each
    // document retains its own approvers/placeholders/toggles when the user navigates away.
    useEffect(() => {
        if (!isBatchMode) {
            return;
        }
        setBatchDocs(prev => {
            const doc = prev[activeDocIndex];
            if (!doc) return prev;
            if (doc.approvers === approvers && doc.placeholders === placeholders &&
                doc.reviewerCanDownload === reviewerCanDownload && doc.embedQrCode === embedQrCode) {
                return prev;
            }
            const next = prev.slice();
            next[activeDocIndex] = { ...doc, approvers, placeholders, reviewerCanDownload, embedQrCode };
            return next;
        });
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [approvers, placeholders, reviewerCanDownload, embedQrCode]);

    // TEMPORARY
    useEffect(() => {
        console.log("TOTAL PDF PAGES:", numPages);
    }, [numPages]);

    useEffect(() => {
        console.log("NUMPAGES STATE:", numPages);
    }, [numPages]);

        const handleSignatureDrop = (
            page: number,
            x: number,
            y: number,
            approverId: number,
            type: 'initial' | 'signature',
            checklistName: boolean,
            checklistDate: boolean,
            checklistBadge: boolean,
            widthPercent?: number,
            heightPercent?: number
        ): void => {

            const approver = approvers.find(
                a => a.id === approverId
            );

            // ALL PAGES
            if (approver?.placementMode === 'all') {

            const groupId =
                `${Date.now()}-${approverId}-${type}`;

            const newPlaceholders: ISignaturePlaceholder[] = [];

            for (let p = 1; p <= numPages; p++) {

                newPlaceholders.push({
                    id: `${Date.now()}-${p}-${Math.random()}`,
                    groupId: groupId,
                    page: p,
                    x,
                    y,
                    approverId,
                    checklistName,
                    checklistDate,
                    checklistBadge,
                    type,
                    width: widthPercent || 20,
                    height: heightPercent || 15
                });
            }

            setPlaceholders(prev => [
                ...prev,
                ...newPlaceholders
            ]);

            return;
        }

            // PAGE RANGE
            if (approver?.placementMode === 'range') {

                const fromPage =
                    approver.rangeFromPage || 1;

                const toPage =
                    approver.rangeToPage || 1;

                if (fromPage > toPage) {

                    alert(
                        'From Page cannot be greater than To Page'
                    );

                    return;
                }

                if (toPage > numPages) {

                    alert(
                        `This PDF only has ${numPages} pages`
                    );

                    return;
                }

                const groupId =
                    `${Date.now()}-${approverId}-${type}`;

                const newPlaceholders: ISignaturePlaceholder[] = [];

                for (
                    let p = fromPage;
                    p <= toPage;
                    p++
                ) {

                    newPlaceholders.push({
                        id: `${Date.now()}-${p}-${Math.random()}`,
                        groupId: groupId,
                        page: p,
                        x,
                        y,
                        approverId,
                        checklistName,
                        checklistDate,
                        checklistBadge,
                        type,
                        width: widthPercent || 20,
                        height: heightPercent || 15
                    });
                }

                setPlaceholders(prev => [
                    ...prev,
                    ...newPlaceholders
                ]);

                return;
            }

            // SELECTED PAGES
            if (approver?.placementMode === 'selected') {

                const pages =
                    approver.selectedPages || [];

                if (pages.length === 0) {

                    alert(
                        'Please select at least one page'
                    );

                    return;
                }

                const groupId =
                    `${Date.now()}-${approverId}-${type}`;

                const newPlaceholders: ISignaturePlaceholder[] = [];

                pages.forEach(p => {

                    newPlaceholders.push({

                        id:
                            `${Date.now()}-${p}-${Math.random()}`,

                        groupId,

                        page: p,

                        x,
                        y,

                        approverId,

                        checklistName,
                        checklistDate,
                        checklistBadge,

                        type,

                        width:
                            widthPercent || 20,

                        height:
                            heightPercent || 15
                    });

                });

                setPlaceholders(prev => [
                    ...prev,
                    ...newPlaceholders
                ]);

                return;
            }

            // CURRENT PAGE
            const newPlaceholder: ISignaturePlaceholder = {
                id: Date.now().toString() + Math.random().toString(),
                page,
                x,
                y,
                approverId,
                checklistName,
                checklistDate,
                checklistBadge,
                type,
                width: widthPercent || 20,
                height: heightPercent || 15
            };

            setPlaceholders([
                ...placeholders,
                newPlaceholder
            ]);
        };

            const placeAtCenter = (
        approver: IApprover,
        type: 'initial' | 'signature',
    ): void => {
        // "Current Page" mode uses the page on screen; other modes ignore this value
        const targetPage = Math.min(Math.max(visiblePage, 1), numPages || 1);
        handleSignatureDrop(
            targetPage, 50, 50,
            approver.id,
            type,
            approver.includeName,
            approver.includeDate,
            approver.includeBadge,
            25,    // ← width (%): perlebar
            10,    // ← height (%): perpendek → jadi persegi panjang mendatar
        );
    };
        const updatePlaceholder = (
            id: string,
            updates: Partial<ISignaturePlaceholder>
        ): void => {

            setPlaceholders(prev => {

                const target = prev.find(p => p.id === id);

                if (!target) {
                    return prev;
                }

                if (target.groupId) {

                    return prev.map(p => {

                        if (p.groupId !== target.groupId) {
                            return p;
                        }

                        return {
                            ...p,
                            x: updates.x !== undefined
                                ? updates.x
                                : p.x,
                            y: updates.y !== undefined
                                ? updates.y
                                : p.y,
                            width: updates.width !== undefined
                                ? updates.width
                                : p.width,
                            height: updates.height !== undefined
                                ? updates.height
                                : p.height
                        };
                    });
                }

                return prev.map(p =>
                    p.id === id
                        ? {
                            ...p,
                            ...updates
                        }
                        : p
                );
            });
        };

    const removePlaceholder = (id: string): void => {
        setPlaceholders(placeholders.filter(p => p.id !== id));
    };

    const clearApproverPlaceholders = (
        approverId: number
    ): void => {

        setPlaceholders(
            placeholders.filter(
                p => p.approverId !== approverId
            )
        );
    };

    const handleRemoveApprover = (id: number): void => {
        setApprovers(approvers.filter(a => a.id !== id));
        setPlaceholders(placeholders.filter(p => p.approverId !== id));
    };

    // Reusable per-document creation routine shared by the single-flow save and the batch loop.
    // When `params.batch` is undefined the LIST_PROCESS payload is byte-for-byte identical to the
    // original single-document save; when provided it ALSO stamps BatchId/BatchMode on the item.
    const createProcessForDocument = async (params: ICreateProcessParams): Promise<void> => {
        const { fileRef: docFileRef, fileBlob: docBlob, approvers: docApprovers, placeholders: docPlaceholders, reviewerCanDownload: docReviewerCanDownload, embedQrCode: docEmbedQrCode, currentUserId, spDoc, spList, batch } = params;

        // Generate Title ESIN-ddmmyy
        const now = new Date();
        const day = String(now.getDate()).padStart(2, '0');
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const year = String(now.getFullYear()).slice(-2);
        const hours = String(now.getHours()).padStart(2, '0');
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const seconds = String(now.getSeconds()).padStart(2, '0');
        const randomSuffix = Math.random().toString(36).substring(2, 5).toUpperCase();
        const esinTitle = `REQ-ESIGN-${day}${month}${year}${hours}${minutes}${seconds}${randomSuffix}`;

        // Embed Barcode Logic Start 
        let finalBlobToHash = docBlob;
        if (docEmbedQrCode) {
            try {
                const qrContent = `${TENANT_DOMAIN}/${SITE_REDIRECT}/SitePages/Document-Tracker.aspx?code=${esinTitle}`;
                const qrImageBytes = await generateQrCodeImageBytes(qrContent);

                // 2. Manipulate PDF
                const arrayBuffer = await docBlob.arrayBuffer();
                const pdfDoc = await PDFDocument.load(arrayBuffer);
                const pngImage = await pdfDoc.embedPng(qrImageBytes);

                // 3. Draw on the last page bottom-right, sized as 12% of the smaller page dimension
                const pages = pdfDoc.getPages();

                if (pages.length > 0) {
                    const lastPage = pages[pages.length - 1];
                    const { width: pageWidth, height: pageHeight } = lastPage.getSize();

                    // Determine if page is portrait or landscape
                    const isPortrait = pageHeight > pageWidth;
                    // Size: 14% for portrait, 16% for landscape
                    const sizeFactor = isPortrait ? 0.12 : 0.16;

                    const qrSize = Math.min(pageWidth, pageHeight) * sizeFactor;
                    const padding = 2;

                    // Position bottom-right with a margin
                    const marginX = 15;
                    const marginY = 15;
                    const qrX = pageWidth - qrSize - marginX;
                    const qrY = marginY;

                    // Draw white background rectangle behind QR code
                    lastPage.drawRectangle({
                        x: qrX - padding,
                        y: qrY - padding,
                        width: qrSize + padding * 2,
                        height: qrSize + padding * 2,
                        color: rgb(1, 1, 1),
                        borderWidth: 0,
                    });
                    // Draw QR code on top
                    lastPage.drawImage(pngImage, {
                        x: qrX,
                        y: qrY,
                        width: qrSize,
                        height: qrSize,
                    });
                }

                // 4. Save and Overwrite SharePoint File
                const modifiedPdfBytes = await pdfDoc.save();
                finalBlobToHash = new Blob([modifiedPdfBytes as unknown as BlobPart], { type: 'application/pdf' });

                // In @pnp/sp, setContentChunked usually takes a Blob or Stream
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                await spDoc.web.getFileByServerRelativePath(docFileRef).setContentChunked(finalBlobToHash as any, {
                    progress: (data) => console.log("Upload progress:", data.stage, data.offset)
                });

            } catch (qrError) {
                console.error("Failed to embed QR code", qrError);
                throw qrError; // Rethrow to halt the actual processing and show normal error log
            }
        } // end if embedQrCode
        // Embed Barcode Logic End

        // Generate Hash from modified (or original) PDF
        const hashHex = await generatePDFHash(finalBlobToHash);

        // Construct ProcessSignature
        const processSignature = docApprovers.map(approver => {
            const appPlaceholders = docPlaceholders.filter(p => p.approverId === approver.id);

            console.log("APPROVER", approver);

            const placeholderPayload = appPlaceholders.map(p => ({
                PageNo: p.page,
                Left: p.x,
                Top: p.y,
                Width: p.width,
                Height: p.height,
                PlaceType: p.type === 'initial' ? 2 : 1,
                displayName: sanitizeApproverName(approver.displayName),
                BadgeNumber: approver.badgeNumber || "",
                ChecklistDate: p.checklistDate.toString(),
                CheckListName: p.checklistName.toString(),
                CheckListBadge: (p.checklistBadge || false).toString()
            }));

            const mapItem = approvalMap.find(m => m.Approver.Id === approver.id);
            const authorizedApprovers = mapItem ? mapItem.AuthorizedApprovers : "";
            const secretaries = mapItem ? mapItem.Secretaries : "";

            return {
                userId: approver.id,
                displayNameApprover: approver.displayName,
                emailApprover: approver.email,
                badgeNumberApprover: approver.badgeNumber || "",
                authorizedApprovers: authorizedApprovers,
                secretaries: secretaries,
                Placeholder: placeholderPayload
            };
        });

        // Title is already generated above
        // Construct Full JSON
        const payload = {
            title: esinTitle,
            requestor: {
                email: context.pageContext.user.email,
                displayName: context.pageContext.user.displayName,
            },
            hashHex: hashHex,
            fileRef: docFileRef,
            status: "Processing",
            ProcessSignature: processSignature
        };

        const jsonString = JSON.stringify(payload);
        const emailString = docApprovers.map(a => a.email).join(';');

        // Build the LIST_PROCESS add payload. Keep this byte-for-byte identical to the original
        // single-document save; only stamp BatchId/BatchMode when running as part of a batch.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const listProcessPayload: any = {
            Title: payload.title,
            HashHex: payload.hashHex,
            FileRef0: payload.fileRef,
            RequestorId: currentUserId,
            Status: payload.status,
            ProcessData: jsonString,
            UserNotif: emailString,
            isDownloadable: docReviewerCanDownload.toString()
        };
        if (batch) {
            listProcessPayload.BatchId = batch.batchId;
            listProcessPayload.BatchMode = batch.batchMode;
        }

        // Save to LIST_PROCESS list
        await spList.web.lists.getByTitle(LIST_PROCESS).items.add(listProcessPayload);

        // Save to LIST_ACTIVITY_LOG
        await spList.web.lists.getByTitle(LIST_ACTIVITY_LOG).items.add({
            Activity: "Initiated Signature Request",
            UserIdId: currentUserId,
            ProcessTitle: payload.title
        });
    };

    const handleSaveWorkflow = async (): Promise<void> => {
        if (!fileMetadata || !fileBlob) return;

        // Validation: Check if all approvers have at least one box
        let hasMissing = false;
        for (const approver of approvers) {
            const hasBox = placeholders.some(p => p.approverId === approver.id);
            if (!hasBox) {
                hasMissing = true;
            }
        }

        if (hasMissing) {
            setValidationTriggered(true);
            setHideErrorDialog(false);
            // alert("Box Specimen approver ada yang belum di tambahkan");
            return;
        }

        // Clear validation if passed
        setValidationTriggered(false);
        setIsSaving(true);

        try {
            const sourceUrl = new URLSearchParams(window.location.search).get('source') || context.pageContext.web.absoluteUrl;
            const esignSiteUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}`;

            // SP for document operations (source site where the file lives)
            const spDoc: SPFI = getSP(context, sourceUrl);
            // SP for list operations (eSign site where LIST_PROCESS and LIST_ACTIVITY_LOG live)
            const spList: SPFI = getSP(context, esignSiteUrl);

            // Fetch Current User to get ID (resolved once, passed into the helper)
            const currentUser = await spList.web.currentUser();

            // Single flow: call the shared helper with `batch` undefined so the LIST_PROCESS
            // payload stays byte-for-byte identical (no BatchId/BatchMode written).
            await createProcessForDocument({
                fileRef: fileMetadata.fileRef,
                fileBlob: fileBlob,
                approvers: approvers,
                placeholders: placeholders,
                reviewerCanDownload: reviewerCanDownload,
                embedQrCode: embedQrCode,
                currentUserId: currentUser.Id,
                spDoc: spDoc,
                spList: spList
            });

            // alert("Workflow processing initiated successfully!");
            setHideDialog(false);

            /* Redirect moved to Dialog Dismiss/Action
            if (fileMetadata && fileMetadata.folder) {
                window.location.href = window.location.origin + fileMetadata.folder;
            }
            */

        } catch (err) {
            console.error("Error processing workflow:", err);
            alert("Error processing workflow. Check console.");
        } finally {
            setIsSaving(false);
        }
    };

    // Batch submit (FEAT-002): validate every document, then create one Approvals Process item per
    // document, all sharing one BatchId with BatchMode='AllOrNothing'. The existing Power Automate
    // flow triggers on item-created in LIST_PROCESS, so each created item fires its own initiate
    // email automatically — no email code is added here.
    const handleBatchSubmit = async (): Promise<void> => {
        // Compose the authoritative document list by folding the live editing state into the active
        // document (the write-back effect is async and may not have flushed the latest keystroke).
        const docs = batchDocs.map((d, i) => i === activeDocIndex
            ? { ...d, approvers, placeholders, reviewerCanDownload, embedQrCode }
            : d);

        // Validation: every document must have loaded, have >=1 approver, and every approver in a
        // document must have >=1 placeholder box (same rule the single flow enforces).
        for (let i = 0; i < docs.length; i++) {
            const doc = docs[i];
            if (!doc.loaded || !doc.blob) {
                setBatchValidationMessage(`Document ${i + 1} ("${doc.fileName}") failed to load and cannot be submitted.`);
                setActiveDocIndex(i);
                return;
            }
            if (doc.approvers.length === 0) {
                setBatchValidationMessage(`Document ${i + 1} ("${doc.fileName}") has no approver. Add at least one approver.`);
                setActiveDocIndex(i);
                return;
            }
            const missingBox = doc.approvers.some(a => !doc.placeholders.some(p => p.approverId === a.id));
            if (missingBox) {
                setBatchValidationMessage(`Document ${i + 1} ("${doc.fileName}") has an approver without a signature box. Add a box for every approver.`);
                setActiveDocIndex(i);
                return;
            }
        }

        setBatchValidationMessage(null);
        setIsSaving(true);

        try {
            const sourceUrl = sourcePage || context.pageContext.web.absoluteUrl;
            const esignSiteUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}`;

            // SP for document operations (source site where the files live)
            const spDoc: SPFI = getSP(context, sourceUrl);
            // SP for list operations (eSign site where LIST_PROCESS and LIST_ACTIVITY_LOG live)
            const spList: SPFI = getSP(context, esignSiteUrl);

            // Resolve current user once and share across every document in the batch.
            const currentUser = await spList.web.currentUser();

            // One BatchId for the whole batch; AllOrNothing is the only supported mode today.
            const batchId = generateBatchId();
            const batchMode = "AllOrNothing";

            // Create one item per document. If a document's creation throws mid-loop we surface the
            // error and stop. NOTE: distributed rollback of already-created items is OUT OF SCOPE —
            // the all-or-nothing enforcement happens in the approval flow (a reject fails the batch),
            // not by deleting Process items here.
            for (let i = 0; i < docs.length; i++) {
                const doc = docs[i];
                try {
                    await createProcessForDocument({
                        fileRef: doc.fileRef,
                        fileBlob: doc.blob as Blob,
                        approvers: doc.approvers,
                        placeholders: doc.placeholders,
                        reviewerCanDownload: doc.reviewerCanDownload,
                        embedQrCode: doc.embedQrCode,
                        currentUserId: currentUser.Id,
                        spDoc: spDoc,
                        spList: spList,
                        batch: { batchId, batchMode }
                    });
                } catch (docErr) {
                    console.error(`[InitiateSignature] Batch creation failed on document ${i + 1} (${doc.fileRef})`, docErr);
                    setBatchValidationMessage(`Batch failed while creating document ${i + 1} ("${doc.fileName}"). ${i} document(s) were already created and are NOT rolled back. Check the console.`);
                    return;
                }
            }

            setHideDialog(true);
            setHideErrorDialog(true);
            setBatchDialogVisible(true);
        } catch (err) {
            console.error("Error processing batch workflow:", err);
            alert("Error processing batch workflow. Check console.");
        } finally {
            setIsSaving(false);
        }
    };

    const activeDoc = isBatchMode ? batchDocs[activeDocIndex] : undefined;

    return (
        <TouchDragProvider>
            <div className="d-flex flex-column-reverse flex-md-row" style={{ height: 'calc(100vh - 100px)', overflow: 'hidden' }}>
                <style>{`
            .sidebar-responsive {
                width: 100%;
                height: auto;
                max-height: 40vh;
                overflow-y: auto;
                flex-shrink: 0; /* Prevent sidebar from shrinking */
            }
            .pdf-responsive {
                height: auto;
                min-height: 0; /* Crucial for scrolling in flex column */
                flex-basis: 0; /* Start from 0 and grow */
            }
            @media (min-width: 768px) {
                .sidebar-responsive {
                    width: 350px !important;
                    min-width: 350px !important;
                    height: 100%;
                    max-height: 100%;
                    overflow-y: visible;
                }
                .pdf-responsive {
                    height: 100%;
                    flex-basis: auto;
                }
            }
            /* Visual feedback for touch drag mode */
            .touch-drop-active {
                outline: 3px dashed #0d6efd;
                outline-offset: -3px;
                animation: pulse-border 1s infinite;
            }
            @keyframes pulse-border {
                0%, 100% { outline-color: #0d6efd; }
                50% { outline-color: #6ea8fe; }
            }
        `}</style>

                {/* Main Content: PDF */}
                <div className="flex-grow-1 bg-light position-relative border rounded pdf-responsive d-flex flex-column overflow-hidden">
                    {loading && <div className="p-4">Loading...</div>}

                    {!loading && !fileUrl && <div className="p-4 alert alert-warning m-4">No file loaded. Check URL parameters.</div>}

                    {/* Batch navigation: lets the requestor switch between the N documents, each
                        configured independently. Rendered only in batch mode. */}
                    {isBatchMode && batchDocs.length > 0 && (
                        <div className="p-2 border-bottom bg-white sticky-top shadow-sm d-flex flex-wrap align-items-center" style={{ zIndex: 10, gap: 6 }}>
                            <span className="text-muted me-2" style={{ fontSize: 12 }}>
                                Batch ({batchDocs.length} documents) — configuring {activeDocIndex + 1}
                            </span>
                            {batchDocs.map((d, i) => {
                                const complete = d.loaded && d.approvers.length > 0 &&
                                    d.approvers.every(a => d.placeholders.some(p => p.approverId === a.id));
                                const isActive = i === activeDocIndex;
                                return (
                                    <button
                                        key={d.fileRef}
                                        type="button"
                                        onClick={() => setActiveDocIndex(i)}
                                        title={d.fileName}
                                        className={`btn btn-sm ${isActive ? 'btn-primary' : (complete ? 'btn-success' : 'btn-outline-secondary')}`}
                                        style={{ maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                                    >
                                        {i + 1}. {d.fileName}{complete ? ' ✓' : ''}
                                    </button>
                                );
                            })}
                        </div>
                    )}

                    {fileUrl && (
                        <>
                            <div className="p-3 border-bottom bg-white sticky-top shadow-sm" style={{ zIndex: 9 }}>
                                <h5 className="m-0 text-truncate" title={fileMetadata?.fileName}>{fileMetadata?.fileName}</h5>
                            </div>
                            <PdfViewer
                                fileUrl={fileUrl}
                                placeholders={placeholders}
                                approvers={approvers}
                                onDrop={handleSignatureDrop}
                                onUpdatePlaceholder={updatePlaceholder}
                                onRemovePlaceholder={removePlaceholder}
                                onDocumentLoaded={setNumPages}
                                onVisiblePageChange={setVisiblePage}
                            />
                        </>
                    )}
                </div>

                {/* Sidebar */}
                <div className="shadow-lg sidebar-responsive mb-4 mb-md-0 ms-md-3 border rounded" style={{ zIndex: 100 }}>
                    <Sidebar
                        context={context}
                        approvers={approvers}
                        approvalMap={approvalMap}
                        placeholders={placeholders}
                        onApproversChange={setApprovers}
                        onRemoveApprover={handleRemoveApprover}
                        onSave={isBatchMode ? handleBatchSubmit : handleSaveWorkflow}
                        isSaving={isSaving}
                        fileLoaded={!!fileUrl}
                        validationTriggered={validationTriggered}
                        reviewerCanDownload={reviewerCanDownload}
                        onReviewerCanDownloadChange={setReviewerCanDownload}
                        embedQrCode={embedQrCode}
                        onEmbedQrCodeChange={setEmbedQrCode}
                        placementMode={placementMode}
                        onPlacementModeChange={setPlacementMode}
                        onClearApproverPlaceholders={clearApproverPlaceholders}
                        numPages={numPages}
                        visiblePage={visiblePage}
                        onPlaceAtCenter={placeAtCenter}
                    />
                </div>

                {/* Success Dialog */}
                <Dialog
                    hidden={hideDialog}
                    onDismiss={() => {
                        // Redirect on dismiss as well
                        if (fileMetadata && fileMetadata.fileRef) {
                            const redirectUrl = getEncodedFolderUrl(fileMetadata.fileRef, window.location.origin);
                            if (redirectUrl) {
                                window.location.href = redirectUrl;
                            }
                        }
                    }}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Success',
                        subText: 'Workflow processing initiated successfully!'
                    }}
                    modalProps={{
                        isBlocking: true,
                        styles: { main: { maxWidth: 450 } }
                    }}
                >
                    <DialogFooter>
                        <PrimaryButton onClick={() => {
                            if (fileMetadata && fileMetadata.fileRef) {
                                const redirectUrl = getEncodedFolderUrl(fileMetadata.fileRef, window.location.origin);
                                if (redirectUrl) {
                                    window.location.href = redirectUrl;
                                }
                            }
                        }} text="OK" />
                    </DialogFooter>
                </Dialog>

                {/* Validation Error Dialog */}
                <Dialog
                    hidden={hideErrorDialog}
                    onDismiss={() => setHideErrorDialog(true)}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Validation Error',
                        subText: 'Box Specimen approver ada yang belum di tambahkan'
                    }}
                    modalProps={{
                        isBlocking: false,
                        styles: { main: { maxWidth: 450 } }
                    }}
                >
                    <DialogFooter>
                        <PrimaryButton onClick={() => setHideErrorDialog(true)} text="OK" />
                    </DialogFooter>
                </Dialog>

                {/* Pending Process Dialog */}
                <Dialog
                    hidden={!isPendingDialogVisible}
                    onDismiss={() => {
                        if (fileRef) {
                            const redirectUrl = getEncodedFolderUrl(fileRef, window.location.origin);
                            if (redirectUrl) {
                                window.location.href = redirectUrl;
                            }
                        }
                    }}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Document Processing',
                        subText: 'This document is currently being processed.'
                    }}
                    modalProps={{
                        isBlocking: true,
                        styles: { main: { maxWidth: 450 } }
                    }}
                >
                    <DialogFooter>
                        <PrimaryButton onClick={() => {
                            if (fileRef) {
                                const redirectUrl = getEncodedFolderUrl(fileRef, window.location.origin);
                                if (redirectUrl) {
                                    window.location.href = redirectUrl;
                                }
                            }
                        }} text="OK" />
                    </DialogFooter>
                </Dialog>

                {/* Batch Success Dialog */}
                <Dialog
                    hidden={!batchDialogVisible}
                    onDismiss={() => {
                        if (activeDoc && activeDoc.fileRef) {
                            const redirectUrl = getEncodedFolderUrl(activeDoc.fileRef, window.location.origin);
                            if (redirectUrl) {
                                window.location.href = redirectUrl;
                            }
                        }
                    }}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Success',
                        subText: `Batch of ${batchDocs.length} document(s) initiated successfully!`
                    }}
                    modalProps={{
                        isBlocking: true,
                        styles: { main: { maxWidth: 450 } }
                    }}
                >
                    <DialogFooter>
                        <PrimaryButton onClick={() => {
                            if (activeDoc && activeDoc.fileRef) {
                                const redirectUrl = getEncodedFolderUrl(activeDoc.fileRef, window.location.origin);
                                if (redirectUrl) {
                                    window.location.href = redirectUrl;
                                }
                            }
                        }} text="OK" />
                    </DialogFooter>
                </Dialog>

                {/* Batch Validation / Error Dialog */}
                <Dialog
                    hidden={!batchValidationMessage}
                    onDismiss={() => setBatchValidationMessage(null)}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Batch Incomplete',
                        subText: batchValidationMessage || ''
                    }}
                    modalProps={{
                        isBlocking: false,
                        styles: { main: { maxWidth: 450 } }
                    }}
                >
                    <DialogFooter>
                        <PrimaryButton onClick={() => setBatchValidationMessage(null)} text="OK" />
                    </DialogFooter>
                </Dialog>
            </div>
        </TouchDragProvider>
    );
};

export default InitiateSignature;
