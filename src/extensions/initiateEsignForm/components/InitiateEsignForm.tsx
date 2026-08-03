import * as React from "react";
import { Log } from "@microsoft/sp-core-library";
import { Stack, Spinner, SpinnerSize, MessageBar, MessageBarType, Dialog, DialogType, DialogFooter, TextField, PrimaryButton, DefaultButton, Pivot, PivotItem, Checkbox, Text } from "@fluentui/react";
import { Document, Page, pdfjs } from 'react-pdf';
// import { PDFDocument, rgb } from 'pdf-lib'; // Moved to PdfUtils
import 'react-pdf/dist/Page/AnnotationLayer.css';
import 'react-pdf/dist/Page/TextLayer.css';
import { getSP, getGraph } from "../../../common/pnpjsConfig";
import { SPFI, SPFx } from "@pnp/sp";
import { GraphFI } from "@pnp/graph";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import SignaturePad from "signature_pad";
import { TENANT_DOMAIN, LIST_ACTIVITY_LOG, LIST_PROCESS, SITE_REDIRECT, DATABASE_SPECIMEN } from "../../../common/constants";

// Centralized Models
import {
    IInitiateEsignFormProps,
    IInitiateEsignFormState,
    IDrawingHelper,
    ITextAnnotation
} from "../../../common/models/IEsignState";

// Services
import { UserService, EsignDataService } from "../../../common/services";

// Utilities
import { CanvasUtils, PdfUtils, getSiteUrlFromPath, formatBadgeNumber } from "../../../common/utils";

import styles from "./InitiateEsignForm.module.scss";

// Components
import { AnnotationToolbar, AnnotationMode } from "./AnnotationToolbar";
import SignatureLogDialog from "../../initiateEsign/components/SignatureLogDialog";
import { generatePDFHash } from "../../../common/utils/HashHelper";

// Worker Setup - Create Blob URL from bundled worker content
// eslint-disable-next-line @typescript-eslint/no-var-requires
const workerContent = require('../../../common/assets/pdf.worker.min.js');
const workerBlob = new Blob([typeof workerContent === 'string' ? workerContent : workerContent.default || ''], { type: 'application/javascript' });
pdfjs.GlobalWorkerOptions.workerSrc = URL.createObjectURL(workerBlob);

// Interfaces now imported from ../../../common/models/IEsignState

const LOG_SOURCE: string = "InitiateEsignForm";

export default class InitiateEsignForm extends React.Component<IInitiateEsignFormProps, IInitiateEsignFormState> {
    private _containerRef = React.createRef<HTMLDivElement>();
    private _drawingHelper: IDrawingHelper = { isDrawing: false, currentPath: [], x: 0, y: 0 };

    // Signature Pad Refs
    private _signaturePad: SignaturePad | undefined = undefined;
    // private _signaturePadCanvasRef = React.createRef<HTMLCanvasElement>(); // Removed in favor of callback
    // private _signaturePadHelper: IDrawingHelper = { isDrawing: false, currentPath: [], x: 0, y: 0 }; // Removed

    constructor(props: IInitiateEsignFormProps) {
        super(props);
        // Calculate initial width based on window size for mobile
        const isMobile = typeof window !== 'undefined' && window.innerWidth <= 768;
        const initialWidth = isMobile ? Math.min(window.innerWidth - 32, 600) : 600;
        
        this.state = {
            fileUrl: undefined,
            numPages: undefined,
            currentWidth: initialWidth,
            loading: true,
            error: undefined,
            placeholders: [],

            annotationMode: 'view',
            rotation: 0,
            paths: [],
            texts: [],
            penColor: '#ff0000',  // Default red
            penSize: 2,           // Default size
            textSize: 14,         // Default text size

            isSaving: false,
            editingTextId: undefined,
            draggingTextId: undefined,
            isRejectDialogOpen: false,
            rejectComment: '',
            rejectCommentError: '',
            isSuccessDialogOpen: false,

            isBlockingDialogOpen: false,
            blockingDialogTitle: '',
            blockingDialogMessage: '',

            isSignatureDialogOpen: false,
            activePlaceholderIndex: -1,
            userSignatureData: undefined,
            isSignatureLoading: false,
            signatureError: '',
            isValidationErrorOpen: false,
            successMessage: '',
            isApproveDialogOpen: false,
            approveComment: '',
            isLogDialogOpen: false,

            isFileNotFoundDialogOpen: false,
            fileNotFoundProcessId: undefined
        };
    }

    public async componentDidMount(): Promise<void> {
        Log.info(LOG_SOURCE, "React Element: InitiateEsignForm mounted");
        this._updateWidth();
        window.addEventListener("resize", this._updateWidth);

        try {
            const graph: GraphFI = getGraph(this.props.context);
            const employeeId = await UserService.getCurrentUserEmployeeId(graph);
            this.setState({ currentUserId: formatBadgeNumber(employeeId) });
        } catch (e) {
            console.error("Failed to fetch user profile", e);
        }

        await this._fetchFile();
    }

    public componentWillUnmount(): void {
        Log.info(LOG_SOURCE, "React Element: InitiateEsignForm unmounted");
        window.removeEventListener("resize", this._updateWidth);
        if (this.state.fileUrl) {
            URL.revokeObjectURL(this.state.fileUrl);
        }
    }

    public componentDidUpdate(prevProps: IInitiateEsignFormProps, prevState: IInitiateEsignFormState): void {
        if (prevState.paths !== this.state.paths || prevState.currentWidth !== this.state.currentWidth || prevState.rotation !== this.state.rotation) {
            // Re-draw all canvases
            if (this.state.numPages) {
                for (let i = 1; i <= this.state.numPages; i++) {
                    const canvas = this._containerRef.current?.querySelector(`canvas[data-page="${i}"]`) as HTMLCanvasElement;
                    if (canvas) {
                        CanvasUtils.drawPathsOnCanvas(canvas, this.state.paths, i);
                    }
                }
            }
        }

        // Auto-scroll logic removed as per user request (only scroll on validation error)
        // if (prevState.loading === true && this.state.loading === false && this.state.fileUrl) {
        //     setTimeout(() => this._scrollToFirstUnsignedPlaceholder(), 1500);
        // }
    }

    private _updateWidth = (): void => {
        if (typeof window === 'undefined') return;
        if (this._containerRef.current) {
            const containerWidth = this._containerRef.current.clientWidth;
            // Mobile: use most of the screen width, Desktop: use container width with padding
            const isMobile = window.innerWidth <= 768;
            const padding = isMobile ? 16 : 40;
            const maxWidth = isMobile ? window.innerWidth - 32 : containerWidth - padding;
            this.setState({ currentWidth: maxWidth > 0 ? maxWidth : 600 });
        }
    };

    private _fetchFile = async (): Promise<void> => {
        try {
            const result = await EsignDataService.getFileAndMetadata(this.props.context);
            
            this.setState({
                fileUrl: result.fileUrl,
                loading: false,
                placeholders: result.placeholders,
                serverRelativeUrl: result.serverRelativeUrl,
                processId: result.processId,
                processTitle: result.processTitle,
                signingUserId: result.firstAssignedUserId,
                signingUserDisplayName: result.signerName,
                isDownloadable: result.isDownloadable,
                isSecretaryOnly: result.isSecretaryOnly
            });

            // NEW: Hash Verification
            if (result.hashHex) {
                try {
                    // Generate hash from fetched blob (we need to fetch blob again or get it from service)
                    // We used URL.createObjectURL in service. We can fetch it back from Blob URL or modify service to return blob.
                    // Simpler: fetch from localhost/blob url
                    const blob = await fetch(result.fileUrl).then(r => r.blob());
                    const currentHash = await generatePDFHash(blob);

                    if (currentHash !== result.hashHex) {
                        console.warn("Hash Mismatch!", { expected: result.hashHex, actual: currentHash });
                        await this._cancelProcess(result.processId, "Document Hash Mismatch");
                        this.setState({
                            isBlockingDialogOpen: true,
                            blockingDialogTitle: "Document Integrity Error",
                            blockingDialogMessage: "The document source has been modified externally. The process has been automatically canceled."
                        });
                    }
                } catch (hashErr) {
                    console.error("Hash verification failed", hashErr);
                }
            }

        } catch (err) {
            console.error("Error fetching file:", err);

            // Check if error has title property (from validation)
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const error = err as any;
            if (error.title) {
                this.setState({
                    loading: false,
                    isBlockingDialogOpen: true,
                    blockingDialogTitle: error.title,
                    blockingDialogMessage: error.message
                });
            } else {
                let errorMessage = error.message || "Failed to load document.";
                
                // Check if file does not exist (SharePoint error usually contains "does not exist")
                if (errorMessage && errorMessage.indexOf("does not exist") !== -1) {
                    errorMessage = `Document path has changed or file not found in the expected location.`;

                    // Attempt to fetch ProcessID to allow cancellation
                    if (this.props.context.itemId) {
                        try {
                            const sp = getSP(this.props.context);
                            // eslint-disable-next-line @typescript-eslint/no-explicit-any
                            const item: any = await sp.web.lists.getById(this.props.context.list.guid.toString())
                                .items.getById(this.props.context.itemId)
                                .select("ProcessID/Id")
                                .expand("ProcessID")();
                            
                            if (item && item.ProcessID) {
                                 this.setState({ 
                                     isFileNotFoundDialogOpen: true,
                                     fileNotFoundProcessId: item.ProcessID.Id,
                                     loading: false, 
                                     error: errorMessage 
                                 });
                                 return; 
                            }
                        } catch (pidErr) {
                            console.warn("Failed to fetch process ID for cancellation", pidErr);
                        }
                    }
                }

                this.setState({ loading: false, error: errorMessage });
            }
        }
    };

    private _scrollToFirstUnsignedPlaceholder = (): void => {
        const { placeholders, loading, error } = this.state;
        if (loading || error || !placeholders || placeholders.length === 0) return;

        // Find first unsigned placeholder index
        const firstUnsignedIndex = placeholders.findIndex(p => !p.signatureData);

        if (firstUnsignedIndex !== -1) {
            // Find the specific placeholder element by ID
            const element = document.getElementById(`placeholder-${firstUnsignedIndex}`);
            if (element) {
                element.scrollIntoView({ behavior: 'smooth', block: 'center' });
                Log.info(LOG_SOURCE, `Auto-scrolled to unsigned placeholder #${firstUnsignedIndex}`);
            }
        }
    };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private _onDocumentLoadSuccess = async (pdf: any): Promise<void> => {
        const numPages = pdf.numPages;
        this.setState({ numPages });

        // Get rotation of the first page to set initial view correctly
        try {
            const page = await pdf.getPage(1);
            const rotation = page.rotate || 0;
            this.setState({ rotation: rotation });
        } catch (e) {
            console.warn("Failed to get page rotation", e);
        }
    };

    private _handleModeChange = (mode: AnnotationMode): void => {
        this.setState({ annotationMode: mode });
    };

    private _handleRotate = (): void => {
        this.setState((prev) => ({ rotation: (prev.rotation + 90) % 360 }));
    };



    private _handleClear = (): void => {
        if (confirm("Are you sure you want to clear all annotations?")) {
            this.setState({ paths: [], texts: [] });
        }
    };

    private _handleUndo = (): void => {
        this.setState(prev => {
            if (prev.paths.length > 0) {
                const newPaths = [...prev.paths];
                newPaths.pop(); // Remove last path
                return { paths: newPaths };
            }
            return null; // eslint-disable-line @rushstack/no-new-null
        });
    };

    private _handleReject = (): void => {
        this.setState({ isRejectDialogOpen: true });
    };

    private _handleDownloadPdf = async (): Promise<void> => {
        if (!this.state.fileUrl) return;
        try {
            const response = await fetch(this.state.fileUrl);
            const blob = await response.blob();
            const fileName = this.state.serverRelativeUrl
                ? this.state.serverRelativeUrl.split('/').pop() || 'document.pdf'
                : 'document.pdf';
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } catch (err) {
            console.error('Error downloading PDF:', err);
        }
    };

    private _closeRejectDialog = (): void => {
        this.setState({ isRejectDialogOpen: false, rejectComment: '', rejectCommentError: '' });
    };

    private _closeSuccessDialog = (): void => {
        this.setState({ isSuccessDialogOpen: false });
        this.props.onClose();
    };

    private _getClientInfo = async (): Promise<string> => {
        try {
            const response = await fetch('https://ipapi.co/json/');
            const data = await response.json();
            return JSON.stringify({
                ip: data.ip,
                country: data.country_name,
                agent: navigator.userAgent
            });
        } catch (e) {
            console.warn("Failed to fetch client info", e);
            return JSON.stringify({
                ip: "Unknown",
                country: "Unknown",
                agent: navigator.userAgent
            });
        }
    };

    private _closeSignatureDialog = (): void => {
        this.setState({
            isSignatureDialogOpen: false,
            activePlaceholderIndex: -1,
            userSignatureData: undefined,
            signatureError: '',
            // Reset Pad state
            signatureDialogTab: 'saved',
            newSignatureData: undefined,
            saveAsDefault: false,
            signAll: false
        });
    };

    private _handlePlaceholderClick = async (index: number): Promise<void> => {

        if (this.state.isSecretaryOnly) {
            return;
        }

        this.setState({
            isSignatureDialogOpen: true,
            activePlaceholderIndex: index,
            isSignatureLoading: true,
            signatureError: '',
            userSignatureData: undefined, // Reset previous data to ensure fresh fetch
            // Reset Pad state
            signatureDialogTab: 'saved', // Default to saved
            newSignatureData: undefined,
            saveAsDefault: false,
            signAll: false
        });

        try {
            const placeholder = this.state.placeholders[index];
            // We could use placeholder.type if we had distinguished it, but we have 'initial' | 'signature'.
            // Let's rely on that.
            const isInitial = placeholder.type === 'initial';

            await this._fetchUserSignature(isInitial);

        } catch (err) {
            console.error("Error handling placeholder click:", err);
            this.setState({ isSignatureLoading: false, signatureError: "Failed to load signature." });
        }
    };

    private _fetchUserSignature = async (isInitial: boolean): Promise<void> => {
        try {
            const sp: SPFI = getSP(this.props.context);

            const currentUser = !this.state.signingUserId ? await sp.web.currentUser() : undefined;
            const userId = this.state.signingUserId || currentUser?.Id;
            if (!userId) throw new Error("Could not determine user ID");

            let blob: Blob | undefined;

            const tryFetch = async (ext: string): Promise<Blob> => {
                const filename = `${userId}-${isInitial ? 'initial' : 'signature'}.${ext}`;
                const file = sp.web.getFolderByServerRelativePath(DATABASE_SPECIMEN).files.getByUrl(filename);
                return await file.getBlob();
            };

            try {
                blob = await tryFetch('png');
            } catch {
                try {
                    blob = await tryFetch('jpg');
                } catch {
                    blob = await tryFetch('jpeg');
                }
            }

            if (!blob) throw new Error("Signature file not found");

            // Convert to Data URL
            const reader = new FileReader();
            reader.readAsDataURL(blob);
            reader.onloadend = () => {
                const base64data = reader.result as string;
                this.setState({
                    userSignatureData: base64data,
                    isSignatureLoading: false
                });
            };

        } catch (err) {
            console.warn("Signature file not found:", err);

            let errorMessage = `Signature Specimen not found. Please ensure "${isInitial ? 'Initial' : 'Signature'}" is uploaded to Specimen library (PNG/JPG).`;


            const { signingUserId, signingUserDisplayName } = this.state;

            if (signingUserId) {
                errorMessage = `Signature Specimen for the assigned user (${signingUserDisplayName || signingUserId}) was not found. Please ask them to update their signature in the registered page.`;
            }

            this.setState({
                isSignatureLoading: false,
                signatureError: errorMessage
            });
        }
    };

    private _saveSignatureSpecimen = async (dataUrl: string, isInitial: boolean): Promise<void> => {
        try {
            const sp: SPFI = getSP(this.props.context);
            const currentUser = await sp.web.currentUser();
            const userId = currentUser.Id;
            const filename = `${userId}-${isInitial ? 'initial' : 'signature'}.png`;

            const res = await fetch(dataUrl);
            const blob = await res.blob();

            await sp.web.getFolderByServerRelativePath(DATABASE_SPECIMEN)
                .files.addUsingPath(filename, blob, { Overwrite: true });

        } catch (err) {
            console.error("Failed to save specimen:", err);
        }
    };




    private _confirmSignature = async (): Promise<void> => {
        let dataToApply = this.state.userSignatureData;

        if (this.state.signatureDialogTab === 'draw') {
            dataToApply = this.state.newSignatureData;
        }

        if (this.state.activePlaceholderIndex !== undefined && this.state.activePlaceholderIndex >= 0 && dataToApply) {
            // Check if we are drawing and want to save as default
            if (this.state.signatureDialogTab === 'draw' && this.state.saveAsDefault) {
                const p = this.state.placeholders[this.state.activePlaceholderIndex];
                const isInitial = p.type === 'initial';
                this.setState({ isSignatureLoading: true });
                // Save the raw drawing as specimen
                await this._saveSignatureSpecimen(dataToApply, isInitial);
                this.setState({ isSignatureLoading: false });
            }

            const activePlaceholder = this.state.placeholders[this.state.activePlaceholderIndex];
            const activeType = activePlaceholder.type;
            const activeApprover = activePlaceholder.approverName;
            
            // Create composite signature just once for the primary action (optimization: could do usually per placeholder if date changes, 
            // but here date is usually same day. However, placeholder dimensions might differ? 
            // CanvasUtils.createCompositeSignature takes placeholder as arg to fit dimensions. 
            // So we must call it for each placeholder if we want it to fit perfectly or if we rely on placeholder W/H.)
            
            // Actually, we should iterate.

            const newPlaceholders = [...this.state.placeholders];
            const placeholdersToUpdate: number[] = [this.state.activePlaceholderIndex];

            if (this.state.signAll) {
                // Find all OTHER placeholders that match criteria:
                // 1. Same Type (Signature vs Initial)
                // 2. Same Approver (User) - relying on approverName or we should rely on assignment logic?
                //    The logic is: "Sign all MY signature boxes". So generally checked by same user.
                //    Safest is to match 'approverName' or just rely on the fact the user can only click their own placeholders 
                //    (if we had that check). But here let's match Type and Name to be safe.
                //    Also ensure they are not already signed.
                
                newPlaceholders.forEach((p, idx) => {
                    if (idx !== this.state.activePlaceholderIndex &&
                        p.type === activeType && 
                        !p.signatureData && // Only unsigned
                        p.approverName === activeApprover // Match the "group"
                    ) {
                        placeholdersToUpdate.push(idx);
                    }
                });
            }

            this.setState({ isSignatureLoading: true });

            // Apply to all identified placeholders
            for (const idx of placeholdersToUpdate) {
                const targetPlaceholder = newPlaceholders[idx];
                // We generate composite for each because dimensions might vary slightly (unlikely in standard forms but possible)
                // and to be safe.
                const badgeNumberToUse = targetPlaceholder.badgeNumber || "";

                const compositeData = await CanvasUtils.createCompositeSignature(dataToApply, targetPlaceholder, badgeNumberToUse);
                newPlaceholders[idx].signatureData = compositeData;
            }

            this.setState({
                placeholders: newPlaceholders,
                isSignatureLoading: false
            });

            this._closeSignatureDialog();
        }
    };

    private _openApproveDialog = (): void => {
        // Check for missing signatures
        const hasUnsignedPlaceholders = this.state.placeholders.some(p => !p.signatureData);
        if (hasUnsignedPlaceholders) {
            this.setState({ isValidationErrorOpen: true });
            this._scrollToFirstUnsignedPlaceholder();
            return;
        }
        this.setState({ isApproveDialogOpen: true });
    };

    private _closeApproveDialog = (): void => {
        this.setState({ isApproveDialogOpen: false, approveComment: '' });
    };

    private _confirmApprove = async (): Promise<void> => {

        if (this.state.isSecretaryOnly) {
            return;
        }

        this.setState({ isSaving: true, error: undefined }); // Don't close dialog yet
        try {
            const sp: SPFI = getSP(this.props.context);
            const listId = this.props.context.list.guid.toString();
            const itemId = this.props.context.itemId;

            if (!itemId) throw new Error("Item ID invalid");

            // 1. Burn Annotations (including signatures and initials for approval)
            const hash = await this._burnAnnotations(true);

            // 2. Fetch current user
            const currentUser = await sp.web.currentUser();
            const clientInfo = await this._getClientInfo();

            // 3. Update Item (Current Task)
            await sp.web.lists.getById(listId).items.getById(itemId).update({
                Status: "Approved",
                ResponseData: clientInfo,
                Comments: this.state.approveComment,
                executedByIdId: currentUser.Id
            });

            // NEW: Update Hash in Process List
            if (this.state.processId && hash) {
                try {
                    await sp.web.lists.getByTitle(LIST_PROCESS).items.getById(this.state.processId).update({
                        HashHex: hash
                    });
                } catch (hashErr) {
                    console.warn("Failed to update HashHex in Process List", hashErr);
                }
            }

            // 4. Log Activity
            try {
                const fileName = this.state.serverRelativeUrl ? this.state.serverRelativeUrl.split('/').pop() : 'Document';
                const activityText = `Approved ${fileName}`;

                await sp.web.lists.getByTitle(LIST_ACTIVITY_LOG).items.add({
                    Activity: activityText,
                    UserIdId: currentUser.Id,
                    ProcessTitle: this.state.processTitle || "Unknown Process"
                });
            } catch (logErr) {
                console.warn("Failed to log activity:", logErr);
            }

            // Close dialog and show success
            this.setState({ isSaving: false, isApproveDialogOpen: false, isSuccessDialogOpen: true, successMessage: "Document has been approved successfully." });

        } catch (err) {
            console.error(err);
            this.setState({ isSaving: false, error: "Failed to approve document." });
        }
    };

    private _submitReject = async (): Promise<void> => {

        if (this.state.isSecretaryOnly) {
            return;
        }
        // Validation: Comment is mandatory
        if (!this.state.rejectComment || this.state.rejectComment.trim().length === 0) {
            this.setState({ rejectCommentError: "Reason for rejection is required." });
            return;
        }

        this.setState({ isSaving: true, rejectCommentError: '' }); // Don't close dialog yet
        try {
            const sp: SPFI = getSP(this.props.context);
            const listId = this.props.context.list.guid.toString();
            const itemId = this.props.context.itemId;

            if (!itemId) throw new Error("Item ID invalid");

            // CHECK IF ANNOTATIONS EXIST AND BURN THEM (but NOT signatures/initials for rejection)
            const hash = (this.state.paths.length > 0 || this.state.texts.length > 0) ? await this._burnAnnotations(false) : undefined;

            // Fetch current user details to get the ID for the Person field
            const currentUser = await sp.web.currentUser();
            const clientInfo = await this._getClientInfo();

            // Note: For Person columns, use the field name with 'Id' suffix and pass the generic integer ID
            await sp.web.lists.getById(listId).items.getById(itemId).update({
                Status: "Rejected",
                Comments: this.state.rejectComment,
                ResponseData: clientInfo,
                executedByIdId: currentUser.Id
            });

            // NEW: Update Hash in Process List if annotations were burned
            if (this.state.processId && hash) {
                try {
                    await sp.web.lists.getByTitle(LIST_PROCESS).items.getById(this.state.processId).update({
                        HashHex: hash
                    });
                } catch (hashErr) {
                    console.warn("Failed to update HashHex in Process List", hashErr);
                }
            }

            // LOG ACTIVITY
            try {
                const fileName = this.state.serverRelativeUrl ? this.state.serverRelativeUrl.split('/').pop() : 'Document';
                const activityText = `Rejected ${fileName}`;

                await sp.web.lists.getByTitle(LIST_ACTIVITY_LOG).items.add({
                    Activity: activityText,
                    UserIdId: currentUser.Id,
                    ProcessTitle: this.state.processTitle || "Unknown Process"
                });
            } catch (logErr) {
                console.warn("Failed to log activity:", logErr);
            }

            this.setState({ isSaving: false, isRejectDialogOpen: false, isSuccessDialogOpen: true, successMessage: "Document has been rejected successfully." });

        } catch (err) {
            console.error(err);
            this.setState({ isSaving: false, error: "Failed to reject document." });
        }
    };

    private _cancelProcess = async (processId: number, details: string = "Document Hash Mismatch"): Promise<void> => {
        try {
            const sp: SPFI = getSP(this.props.context);
            const currentUser = await sp.web.currentUser();

            // Update Status in Process List
            const item = sp.web.lists.getByTitle(LIST_PROCESS).items.getById(processId);
            const result = await item.select("FileRef", "FileLeafRef", "FileRef0")<{ FileRef: string, FileLeafRef: string, FileRef0: string }>();

            await item.update({
                Status: "Canceled",
                Notes: `System Canceled: ${details}`
            });

            // Update Status in Current List (Task List)
            const listId = this.props.context.list.guid.toString();
            const itemId = this.props.context.itemId;
            if (itemId) {
                await sp.web.lists.getById(listId).items.getById(itemId).update({
                    Status: "Canceled",
                    Comments: `System Canceled: ${details}`
                });
            }

            // Log Activity
            const fileName = result.FileRef ? result.FileRef.split('/').pop() : (this.state.serverRelativeUrl ? this.state.serverRelativeUrl.split('/').pop() : 'Document');
            await sp.web.lists.getByTitle(LIST_ACTIVITY_LOG).items.add({
                Activity: `System Canceled ${details} : ${fileName}`,
                UserIdId: currentUser.Id,
                ProcessTitle: this.state.processTitle || "Unknown Process"
            });

        } catch (err) {
            console.error("Failed to cancel process:", err);
        }
    };





    private _handleFileNotFoundCancel = async (): Promise<void> => {
        if (this.state.fileNotFoundProcessId) {
             await this._cancelProcess(this.state.fileNotFoundProcessId!, "Document Not Found (Moved/Deleted/Renamed)");
        }
        window.location.href = `${TENANT_DOMAIN}/${SITE_REDIRECT}`;
    };

    // --- Drawing Handlers ---




    private _burnAnnotations = async (burnSignatures: boolean = true): Promise<string | undefined> => {
        try {
            if (!this.state.fileUrl) return;

            // 1. Burn Annotations using PdfUtils (Robust Rotation Support)
            const newBlob = await PdfUtils.burnAnnotations(
                this.state.fileUrl,
                this.state.paths,
                this.state.texts,
                this.state.placeholders,
                burnSignatures
            );

            // 2. Overwrite File in SharePoint (Cross-Site Compatible)
            if (this.state.serverRelativeUrl) {
                const fileRef = this.state.serverRelativeUrl;
                // Use helper function to determine correct site URL
                const siteUrl = getSiteUrlFromPath(fileRef);

                const targetWeb = Web(siteUrl).using(SPFx(this.props.context));

                // Update Content
                await targetWeb.getFileByServerRelativePath(fileRef).setContent(newBlob);
            }

            // 3. Generate and return Hash
            return await generatePDFHash(newBlob);
        } catch (err) {
            console.error("Error burning annotations:", err);
            throw err;
        }
    };

    // eslint-disable-next-line @rushstack/no-new-null
    private _onSignatureCanvasRef = (canvas: HTMLCanvasElement | null): void => {
        if (!canvas) return;

        if (this._signaturePad) {
            this._signaturePad.off();
        }

        // FIX: Resize canvas to match display size for correct coordinate mapping
        // SignaturePad usage requirement for responsiveness
        const ratio = Math.max(window.devicePixelRatio || 1, 1);
        canvas.width = canvas.offsetWidth * ratio;
        canvas.height = canvas.offsetHeight * ratio;
        canvas.getContext("2d")?.scale(ratio, ratio);

        this._signaturePad = new SignaturePad(canvas, {
            penColor: "blue",
            minWidth: 1,
            maxWidth: 2.5
        });

        this._signaturePad.addEventListener("endStroke", () => {
            if (this._signaturePad && !this._signaturePad.isEmpty()) {
                // We need to trim or handle ratio here if saving? 
                // toDataURL will return the high-res image.
                const data = this._signaturePad.toDataURL('image/png');
                this.setState({ newSignatureData: data });
            } else {
                this.setState({ newSignatureData: undefined });
            }
        });
    };

    // Original annotation drawing handlers (keep these for PDF annotations)
    private _handleDrawStart = (pageNum: number, e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>): void => {
        if (this.state.annotationMode !== 'pen') return;
        
        // Prevent scrolling on touch devices during drawing
        if ('touches' in e) {
            e.preventDefault();
        }

        const canvas = e.currentTarget;
        const { x, y } = CanvasUtils.getRelativeCoords(e, canvas);

        this._drawingHelper = {
            isDrawing: true,
            currentPath: [{ x, y }],
            x,
            y
        };

        // Start visual feedback immediately on this canvas with selected color/size
        const ctx = canvas.getContext('2d');
        if (ctx) {
            ctx.beginPath();
            ctx.moveTo(x, y);
            ctx.strokeStyle = this.state.penColor;
            ctx.lineWidth = this.state.penSize;
            ctx.lineCap = 'round';
            ctx.lineJoin = 'round';
        }
    };

    private _handleDrawMove = (pageNum: number, e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>): void => {
        if (!this._drawingHelper.isDrawing || this.state.annotationMode !== 'pen') return;
        
        // Prevent scrolling on touch devices during drawing
        if ('touches' in e) {
            e.preventDefault();
        }

        const canvas = e.currentTarget;
        const { x, y } = CanvasUtils.getRelativeCoords(e, canvas);

        this._drawingHelper.currentPath.push({ x, y });

        const ctx = canvas.getContext('2d');
        if (ctx) {
            ctx.lineTo(x, y);
            ctx.stroke();
        }
    };

    private _handleDrawEnd = (pageNum: number): void => {
        if (!this._drawingHelper.isDrawing) return;

        this._drawingHelper.isDrawing = false;

        // Normalize paths to percentage of CURRENT width/height to be responsive
        const canvas = this._containerRef.current?.querySelector(`canvas[data-page="${pageNum}"]`) as HTMLCanvasElement;
        if (canvas) {
            const width = canvas.width;
            const height = canvas.height;

            const normalizedPoints = this._drawingHelper.currentPath.map(p => ({
                x: p.x / width,
                y: p.y / height
            }));

            // Use current pen color and size from state
            this.setState(prev => ({
                paths: [...prev.paths, {
                    page: pageNum,
                    points: normalizedPoints,
                    color: prev.penColor,
                    thickness: prev.penSize
                }]
            }));
        }
    };

    // --- Signature Pad Handlers ---

    // Manual signature handlers removed in favor of SignaturePad library

    private _clearSignatureCanvas = (): void => {
        if (this._signaturePad) {
            this._signaturePad.clear();
            this.setState({ newSignatureData: undefined });
        }
    };





    // eslint-disable-next-line @rushstack/no-new-null
    private _onCanvasRef = (elem: HTMLCanvasElement | null, pageNum: number): void => {
        if (elem) {
            const rect = elem.getBoundingClientRect();
            if (elem.width !== rect.width || elem.height !== rect.height) {
                elem.width = rect.width;
                elem.height = rect.height;
                CanvasUtils.drawPathsOnCanvas(elem, this.state.paths, pageNum);
            }
        }
    };



    private _handleCanvasClick = (pageNum: number, e: React.MouseEvent<HTMLCanvasElement>): void => {
        if (this.state.annotationMode === 'text') {
            const canvas = e.currentTarget;
            const { x, y } = CanvasUtils.getRelativeCoords(e, canvas);

            // Normalize
            const normalizedX = x / canvas.width;
            const normalizedY = y / canvas.height;

            const newId = Date.now().toString();

            this.setState(prev => ({
                texts: [...prev.texts, {
                    id: newId,
                    page: pageNum,
                    x: normalizedX,
                    y: normalizedY,
                    text: "", // Start empty, user will type
                    color: 'black',
                    size: prev.textSize * (1 / canvas.width) * 100 // Use selected text size, normalized
                }],
                editingTextId: newId
            }));
        }
    };




    // Revised Delete using object removal
    private _deleteText = (t: ITextAnnotation): void => {
        this.setState(prev => ({ texts: prev.texts.filter(item => item !== t) }));
    };






    public render(): React.ReactElement<IInitiateEsignFormProps> {
        const { fileUrl, numPages, loading, error, currentWidth, placeholders, annotationMode, rotation } = this.state;

        // Note: We no longer block scroll at container level for pen mode
        // Scroll prevention is handled at canvas level during active drawing only
        
        // Safe check for mobile - compute once at render
        const isMobile = typeof window !== 'undefined' && window.innerWidth <= 768;
        const containerPadding = isMobile ? 8 : 20;
        
        return (
            <div 
                className={styles.initiateEsignForm} 
                style={{ 
                    padding: containerPadding, 
                    backgroundColor: "#f3f2f1", 
                    minHeight: '100vh'
                    // Removed: overflow, height, touchAction restrictions
                    // Scrolling is now allowed in pen mode; only blocked during active drawing
                }}
            >
                {/* Style Overrides for Form Customizer */}
                <style>{`
            #spLeftNav { display: none !important; }
            div[class^="canvasContent"] { max-width: 100% !important; width: 100% !important; }
            div[class^="mainContent"] { margin-left: 0 !important; }
            .react-pdf__Page__canvas { 
                margin: 0 auto; 
                max-width: 100% !important;
                height: auto !important;
            }
            /* Mobile-specific responsive styles */
            @media (max-width: 768px) {
                .react-pdf__Page { 
                    max-width: 100% !important; 
                    overflow: hidden;
                }
                .react-pdf__Page__canvas {
                    width: 100% !important;
                    height: auto !important;
                }
            }
        `}</style>

                <div 
                    ref={this._containerRef} 
                    style={{ 
                        maxWidth: 1000, 
                        margin: '0 auto', 
                        backgroundColor: 'white', 
                        borderRadius: 4, 
                        boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
                        // Use flexbox for fixed header + scrollable body layout
                        display: 'flex',
                        flexDirection: 'column',
                        height: 'calc(100vh - 40px)', // Full height minus outer padding
                        overflow: 'hidden'
                    }}
                >
                    {/* Fixed Header Section */}
                    <div style={{ 
                        flexShrink: 0,
                        backgroundColor: 'white', 
                        padding: containerPadding,
                        paddingBottom: 10,
                        borderBottom: '1px solid #e1dfdd'
                    }}>
                        <Stack horizontalAlign="center" verticalAlign="center" className="mb-4">
                            <h2 style={{ margin: 0, color: '#0078d4' }}>Review Document</h2>
                            {this.state.serverRelativeUrl && (
                                <Text variant="medium" styles={{ root: { color: '#605e5c', marginTop: 4, fontWeight: 600 } }}>
                                    {this.state.serverRelativeUrl.split('/').pop()}
                                </Text>
                            )}
                        </Stack>

                        <AnnotationToolbar
                            currentMode={annotationMode}
                            onModeChange={this._handleModeChange}
                            onRotate={this._handleRotate}
                            onUndo={this._handleUndo}
                            canUndo={this.state.paths.length > 0}
                            textCount={this.state.texts.length}
                            signedCount={this.state.placeholders.filter(p => p.signatureData).length}
                            totalPlaceholders={this.state.placeholders.length}
                            onClear={this._handleClear}
                            onReject={this._handleReject}
                            onApprove={this._openApproveDialog}
                            onViewLog={this._openLogDialog}
                            disabled={this.state.isSaving || this.state.loading || !!this.state.error}
                            penColor={this.state.penColor}
                            penSize={this.state.penSize}
                            onPenColorChange={(color) => this.setState({ penColor: color })}
                            onPenSizeChange={(size) => this.setState({ penSize: size })}
                            textSize={this.state.textSize}
                            onTextSizeChange={(size) => this.setState({ textSize: size })}
                            isDownloadable={this.state.isDownloadable}
                            onDownload={this._handleDownloadPdf}
                            isSecretaryOnly={this.state.isSecretaryOnly}
                        />
                    </div>

                    {loading && <div style={{ padding: 20 }}><Spinner size={SpinnerSize.large} label="Loading Document..." /></div>}

                    {error && (
                        <div style={{ padding: containerPadding }}>
                            <MessageBar messageBarType={MessageBarType.error}>
                                {error}
                            </MessageBar>
                        </div>
                    )}

                    {/* Scrollable Document Area */}
                    {fileUrl && !loading && (
                        <div 
                            style={{ 
                                flex: 1,
                                overflowY: 'auto',
                                overflowX: 'hidden',
                                padding: containerPadding,
                                backgroundColor: '#f3f2f1',
                                WebkitOverflowScrolling: 'touch'
                            }}
                        >
                            <Document
                                file={fileUrl}
                                onLoadSuccess={this._onDocumentLoadSuccess}
                                loading={<Spinner label="Rendering PDF..." />}
                                error={<MessageBar messageBarType={MessageBarType.error}>Failed to render PDF.</MessageBar>}
                            >
                                {numPages && Array.from(new Array(numPages), (el, index) => {
                                    const pageNum = index + 1;
                                    const pagePlaceholders = placeholders.filter(p => p.page === pageNum);

                                    return (
                                        <div
                                            key={`page_${pageNum}`}
                                            data-page={pageNum}
                                            className="mb-4"
                                            style={{ 
                                                position: 'relative', 
                                                overflow: 'hidden',
                                                // Visual separation between pages
                                                border: '1px solid #e1dfdd',
                                                borderRadius: 4,
                                                boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
                                                marginBottom: 24,
                                                backgroundColor: '#fff'
                                            }}
                                            onMouseMove={(e) => {
                                                // Global Drag Handler for this page (Mouse)
                                                if (this.state.draggingTextId) {
                                                    const container = e.currentTarget;
                                                    const { x, y } = CanvasUtils.getRelativeCoords(e, container);
                                                    const normalizedX = x / container.clientWidth;
                                                    const normalizedY = y / container.clientHeight;

                                                    this.setState(prev => ({
                                                        texts: prev.texts.map(t =>
                                                            t.id === this.state.draggingTextId
                                                                ? { ...t, x: normalizedX, y: normalizedY }
                                                                : t
                                                        )
                                                    }));
                                                }
                                            }}
                                            onTouchMove={(e) => {
                                                // Global Drag Handler for this page (Touch)
                                                if (this.state.draggingTextId) {
                                                    const container = e.currentTarget;
                                                    const { x, y } = CanvasUtils.getRelativeCoords(e, container);
                                                    const normalizedX = x / container.clientWidth;
                                                    const normalizedY = y / container.clientHeight;

                                                    this.setState(prev => ({
                                                        texts: prev.texts.map(t =>
                                                            t.id === this.state.draggingTextId
                                                                ? { ...t, x: normalizedX, y: normalizedY }
                                                                : t
                                                        )
                                                    }));
                                                } else if (this.state.annotationMode === 'pen' && this._drawingHelper.isDrawing) {
                                                    // Drawing handler is on canvas
                                                }
                                            }}
                                            onMouseUp={() => {
                                                if (this.state.draggingTextId) {
                                                    this.setState({ draggingTextId: undefined });
                                                }
                                            }}
                                            onTouchEnd={() => {
                                                if (this.state.draggingTextId) {
                                                    this.setState({ draggingTextId: undefined });
                                                }
                                            }}
                                            onMouseLeave={() => {
                                                if (this.state.draggingTextId) {
                                                    this.setState({ draggingTextId: undefined });
                                                }
                                            }}
                                        >
                                            {/* Page number indicator */}
                                            <div style={{
                                                position: 'absolute',
                                                top: 8,
                                                right: 8,
                                                backgroundColor: 'rgba(0,0,0,0.6)',
                                                color: 'white',
                                                padding: '2px 8px',
                                                borderRadius: 12,
                                                fontSize: 11,
                                                fontWeight: 600,
                                                zIndex: 50,
                                                pointerEvents: 'none'
                                            }}>
                                                {pageNum} / {numPages}
                                            </div>
                                            <Page
                                                pageNumber={pageNum}
                                                width={currentWidth}
                                                rotate={rotation}
                                                renderTextLayer={false}
                                                renderAnnotationLayer={false}
                                                // Use higher devicePixelRatio for mobile to ensure text is sharp when zoomed
                                                // Mobile gets 3x, Desktop gets native DPR capped at 2
                                                devicePixelRatio={typeof window !== 'undefined' 
                                                    ? (window.innerWidth <= 768 ? 4 : Math.min(window.devicePixelRatio || 1, 2)) 
                                                    : 1}
                                                onLoadSuccess={(page) => {
                                                    // When page loads, we can access its dimensions if needed
                                                    // But we rely on overlay canvas.
                                                }}
                                            />
                                            {/* CANVAS OVERLAY FOR DRAWING */}
                                            <canvas
                                                data-page={pageNum}
                                                width={currentWidth}
                                                height={currentWidth * 1.414} // Aspect ratio assumption for A4
                                                style={{
                                                    position: 'absolute',
                                                    top: 0,
                                                    left: 0,
                                                    width: '100%',
                                                    height: '100%',
                                                    zIndex: 20,
                                                    cursor: annotationMode === 'pen' ? 'crosshair' : (annotationMode === 'text' ? 'text' : 'default'),
                                                    pointerEvents: (annotationMode === 'pen' || annotationMode === 'text') ? 'auto' : 'none',
                                                    // Touch action: none on canvas only when pen mode to allow drawing without scroll interference
                                                    // The e.preventDefault() in touch handlers also helps but this is more reliable
                                                    touchAction: annotationMode === 'pen' ? 'none' : 'auto'
                                                }}
                                                onMouseDown={(e) => this._handleDrawStart(pageNum, e)}
                                                onMouseMove={(e) => this._handleDrawMove(pageNum, e)}
                                                onMouseUp={() => this._handleDrawEnd(pageNum)}
                                                onMouseLeave={() => this._handleDrawEnd(pageNum)}
                                                onTouchStart={(e) => this._handleDrawStart(pageNum, e)}
                                                onTouchMove={(e) => this._handleDrawMove(pageNum, e)}
                                                onTouchEnd={() => this._handleDrawEnd(pageNum)}
                                                onClick={(e) => this._handleCanvasClick(pageNum, e)}
                                                ref={(el) => this._onCanvasRef(el, pageNum)}
                                            />

                                            {/* RENDER TEXT ANNOTATIONS */}
                                            {this.state.texts.filter(t => t.page === pageNum).map((t, i) => {
                                                const isEditing = this.state.editingTextId === t.id;
                                                const calculatedFontSize = t.size ? (t.size / 100) * currentWidth : 16; /* Base 16px fallback */

                                                return (
                                                    <div key={t.id} style={{
                                                        position: 'absolute',
                                                        left: `${t.x * 100}%`,
                                                        top: `${t.y * 100}%`,
                                                        zIndex: 30,
                                                        pointerEvents: 'auto',
                                                        cursor: isEditing ? 'text' : 'move'
                                                    }}
                                                        onMouseDown={(e) => {
                                                            if (this.state.annotationMode !== 'view' && !isEditing) {
                                                                e.stopPropagation();
                                                                if (e.target !== e.currentTarget && (e.target as HTMLElement).tagName === 'SPAN') return;
                                                                this.setState({ draggingTextId: t.id });
                                                            }
                                                        }}
                                                        onTouchStart={(e) => {
                                                            if (this.state.annotationMode !== 'view' && !isEditing) {
                                                                // e.stopPropagation(); // Usually good to bubble touch unless handling fully
                                                                this.setState({ draggingTextId: t.id });
                                                            }
                                                        }}
                                                        onClick={(e) => {
                                                            e.stopPropagation();
                                                            if (!isEditing && !this.state.draggingTextId) {
                                                                // Entering edit mode on click? Or double click?
                                                                // User said "click drag position ... then fill"
                                                                // Let's allow click to edit if not dragging
                                                                this.setState({ editingTextId: t.id });
                                                            }
                                                        }}
                                                    >
                                                        {isEditing ? (
                                                            <input
                                                                autoFocus
                                                                value={t.text}
                                                                style={{
                                                                    fontSize: `${calculatedFontSize}px`,
                                                                    border: '1px dashed blue',
                                                                    background: 'rgba(255,255,255,0.8)',
                                                                    outline: 'none',
                                                                    minWidth: 50,
                                                                    color: 'black'
                                                                }}
                                                                onChange={(e) => {
                                                                    const val = e.target.value;
                                                                    this.setState(prev => ({
                                                                        texts: prev.texts.map(item => item.id === t.id ? { ...item, text: val } : item)
                                                                    }));
                                                                }}
                                                                onBlur={() => {
                                                                    this.setState({ editingTextId: undefined });
                                                                    if (!t.text.trim()) {
                                                                        // Remove empty text on blur
                                                                        this._deleteText(t);
                                                                    }
                                                                }}
                                                                onKeyDown={(e) => {
                                                                    if (e.key === 'Enter') {
                                                                        this.setState({ editingTextId: undefined });
                                                                    }
                                                                }}
                                                            />
                                                        ) : (
                                                            <div style={{ userSelect: 'none', color: t.color, fontSize: `${calculatedFontSize}px`, whiteSpace: 'nowrap', display: 'flex', alignItems: 'center' }}>
                                                                {t.text || "Type here..."}
                                                                <div
                                                                    className="ms-1 text-danger cursor-pointer"
                                                                    style={{
                                                                        width: 16,
                                                                        height: 16,
                                                                        display: 'flex',
                                                                        justifyContent: 'center',
                                                                        alignItems: 'center',
                                                                        fontSize: 12,
                                                                        lineHeight: 1,
                                                                        cursor: 'pointer',
                                                                        background: 'white',
                                                                        border: '1px solid red',
                                                                        borderRadius: '50%',
                                                                        marginLeft: 4
                                                                    }}
                                                                    onClick={(e) => {
                                                                        e.stopPropagation();
                                                                        this._deleteText(t);
                                                                    }}
                                                                >x</div>
                                                            </div>
                                                        )}
                                                    </div>
                                                );
                                            })}




                                            {/* RENDER PLACEHOLDERS OVERLAY */}
                                            {pagePlaceholders.map((p, idx) => {
                                                const isInitial = p.type === 'initial';
                                                const borderColor = isInitial ? 'rgba(45, 137, 239, 0.8)' : 'rgba(0, 163, 0, 0.8)';
                                                const bgColor = isInitial ? 'rgba(45, 137, 239, 0.1)' : 'rgba(0, 163, 0, 0.1)';
                                                const labelBg = isInitial ? 'rgba(45, 137, 239, 0.8)' : 'rgba(0, 163, 0, 0.8)';

                                                const isLandscape = p.width > p.height;
                                                return (
                                                    <div
                                                        key={`ph_${idx}`}
                                                        onClick={(e) => {
                                                            e.stopPropagation();
                                                            // Find the real index in the global array
                                                            // We can do this by reference or ID if possible, but matching logic works too
                                                            // Since pagePlaceholders is just a filter, we can find the index in state.placeholders
                                                            const globalIndex = this.state.placeholders.indexOf(p);
                                                            this._handlePlaceholderClick(globalIndex).catch(console.error);
                                                        }}
                                                        id={`placeholder-${this.state.placeholders.indexOf(p)}`}
                                                        style={{
                                                            position: 'absolute',
                                                            left: `${p.x * 100}%`,
                                                            top: `${p.y * 100}%`,
                                                            width: `${p.width * 100}%`,
                                                            height: `${p.height * 100}%`,
                                                            border: p.signatureData ? 'none' : `2px dashed ${borderColor}`,
                                                            backgroundColor: p.signatureData ? 'transparent' : bgColor,
                                                            zIndex: 10,
                                                            cursor: 'pointer'
                                                        }}
                                                    >
                                                        {!p.signatureData && (
                                                            <>
                                                                <div
                                                                    style={{
                                                                        position: 'absolute',
                                                                        top: -20,
                                                                        left: -2,
                                                                        backgroundColor: labelBg,
                                                                        color: 'white',
                                                                        fontSize: '10px',
                                                                        padding: '1px 4px',
                                                                        borderRadius: '3px',
                                                                        whiteSpace: 'nowrap'
                                                                    }}
                                                                >
                                                                    {p.approverName}
                                                                </div>
                                                                <div style={{
                                                                    display: 'flex',
                                                                    justifyContent: 'center',
                                                                    alignItems: 'center',
                                                                    width: '100%',
                                                                    height: '100%',
                                                                    color: borderColor,
                                                                    fontWeight: 600,
                                                                    // Responsive font based on container width approx
                                                                    // Since we don't have container queries easily without css, let's use a conservative calc or smaller default
                                                                    fontSize: 'container',
                                                                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                                                    containerType: 'size' as any
                                                                }}
                                                                >
                                                                    <div style={{ fontSize: '12cqw', textAlign: 'center' }}>
                                                                        {isInitial ? "Initial Here" : "Sign Here"}
                                                                    </div>
                                                                </div>
                                                            </>
                                                        )}
                                                        {p.signatureData && (
                                                            <img
                                                                src={p.signatureData}
                                                                alt="Signed"
                                                                style={{
                                                                    width: '100%',
                                                                    height: '100%',
                                                                    objectFit: 'contain',
                                                                    ...(isLandscape ? { position: 'absolute', top: 0, left: 0 } : {})
                                                                }}
                                                            />
                                                        )}
                                                    </div>
                                                );
                                            })}
                                        </div>
                                    )
                                })}
                            </Document>
                        </div>
                    )}

                    {/* Specimen Signature Dialog */}
                    {(() => {
                        const idx = this.state.activePlaceholderIndex;
                        const activePlaceholder = (idx !== undefined && idx >= 0) ? this.state.placeholders[idx] : null;
                        const isInitial = activePlaceholder && activePlaceholder.type === 'initial';

                        // Logic to determine if assigned user matches current user due to proxy signing rules
                        // We use legacyPageContext to get the integer User ID of the current user
                        const legacyContext = this.props.context.pageContext.legacyPageContext;
                        const currentSpUserId = legacyContext ? parseInt(legacyContext.userId) : 0;

                        const signingUserId = this.state.signingUserId;

                        // Allow draw if:
                        // 1. No specific signingUserId is set
                        // 2. OR the signingUserId matches the current logged in user's ID
                        const canDraw = !signingUserId || (signingUserId === currentSpUserId);

                        return (
                            <Dialog
                                hidden={!this.state.isSignatureDialogOpen}
                                onDismiss={this._closeSignatureDialog}
                                dialogContentProps={{
                                    type: DialogType.normal,
                                    title: isInitial ? 'Adopt Initial' : 'Adopt Signature'
                                }}
                                modalProps={{
                                    isBlocking: true,
                                    styles: { main: { maxWidth: '900px' } }
                                }}
                            >
                                <Pivot
                                    aria-label="Signature Options"
                                    selectedKey={this.state.signatureDialogTab || 'saved'}
                                    onLinkClick={(item?: PivotItem) => this.setState({ signatureDialogTab: item?.props.itemKey })}
                                >
                                    <PivotItem headerText={isInitial ? "Saved Initial" : "Saved Signature"} itemKey="saved">
                                        <div style={{ textAlign: 'center', padding: 20 }}>
                                            {this.state.isSignatureLoading ? (
                                                <Spinner size={SpinnerSize.large} label="Fetching Specimen Signature..." />
                                            ) : this.state.signatureError ? (
                                                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                                                    <MessageBar messageBarType={MessageBarType.error}>{this.state.signatureError}</MessageBar>
                                                    {/* Only show switch button if allowed */}
                                                    {canDraw && (
                                                        <DefaultButton
                                                            text="Switch to Draw"
                                                            iconProps={{ iconName: 'EditB' }}
                                                            styles={{ root: { marginTop: 10 } }}
                                                            onClick={() => this.setState({ signatureDialogTab: 'draw' })}
                                                        />
                                                    )}
                                                </div>
                                            ) : (
                                                <div style={{ border: '1px solid #ddd', padding: 10, backgroundColor: '#f9f9f9', minHeight: 100, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                                    {this.state.userSignatureData ? (
                                                        <img
                                                            src={this.state.userSignatureData}
                                                            alt="Signature Specimen"
                                                            style={{ maxWidth: '100%', maxHeight: 200, objectFit: 'contain' }}
                                                        />
                                                    ) : (
                                                        <span>No signature data loaded.</span>
                                                    )}
                                                </div>
                                            )}

                                            {!this.state.isSignatureLoading && !this.state.signatureError && (
                                                <div style={{ marginTop: 10, fontSize: 12, color: '#666' }}>
                                                    This signature will be apply to the document.
                                                </div>
                                            )}

                                            {!this.state.isSignatureLoading && !this.state.signatureError && this.state.userSignatureData && (
                                                <div style={{ marginTop: 15, display: 'flex', justifyContent: 'center' }}>
                                                    <Checkbox
                                                        label={`Sign all my ${isInitial ? 'Initials' : 'Signatures'}`}
                                                        checked={this.state.signAll}
                                                        onChange={(e, checked) => this.setState({ signAll: !!checked })}
                                                    />
                                                </div>
                                            )}
                                        </div>
                                    </PivotItem>

                                    {canDraw && (
                                        <PivotItem headerText="Draw New" itemKey="draw">
                                            <div style={{ padding: '0 10px', textAlign: 'center' }}>
                                                <div style={{ border: '1px dashed #0078d4', display: 'block', position: 'relative', backgroundColor: 'white', maxWidth: '100%', overflow: 'hidden' }}>
                                                    <canvas
                                                        ref={this._onSignatureCanvasRef}
                                                        width={600}
                                                        height={400}
                                                        style={{ display: 'block', width: '100%', height: 'auto', maxWidth: '100%', border: '1px solid #e1dfdd', aspectRatio: '3/2' }}
                                                    />
                                                </div>
                                                <div style={{ marginTop: 15, display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', maxWidth: 600, margin: '15px auto 0' }}>
                                                    <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                                                        <Checkbox
                                                            label="Save as default (Update Specimen)"
                                                            checked={this.state.saveAsDefault}
                                                            onChange={(e?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => this.setState({ saveAsDefault: !!checked })}
                                                        />
                                                        <Checkbox
                                                            label={`Sign all my ${isInitial ? 'Initials' : 'Signatures'}`}
                                                            checked={this.state.signAll}
                                                            onChange={(e, checked) => this.setState({ signAll: !!checked })}
                                                        />
                                                    </div>
                                                    <DefaultButton text="Clear" iconProps={{ iconName: 'EraseTool' }} onClick={this._clearSignatureCanvas} />
                                                </div>

                                                <div style={{ fontSize: 12, color: '#666', marginTop: 10 }}>
                                                    {this.state.saveAsDefault ? "This will overwrite your existing specimen in SharePoint." : "This will be used for this session only."}
                                                </div>
                                            </div>
                                        </PivotItem>
                                    )}
                                </Pivot>
                                <DialogFooter>
                                    <PrimaryButton
                                        onClick={this._confirmSignature}
                                        text="Sign"
                                        disabled={
                                            this.state.isSignatureLoading ||
                                            (this.state.signatureDialogTab === 'draw' ? !this.state.newSignatureData : (!this.state.userSignatureData || !!this.state.signatureError))
                                        }
                                    />
                                    <DefaultButton onClick={this._closeSignatureDialog} text="Cancel" />
                                </DialogFooter>
                            </Dialog>
                        );
                    })()}

                    {/* Approve Dialog */}
                    <Dialog
                        hidden={!this.state.isApproveDialogOpen}
                        onDismiss={this._closeApproveDialog}
                        dialogContentProps={{
                            type: DialogType.normal,
                            title: 'Approve Document',
                            subText: 'Are you sure you want to APPROVE this document? This action cannot be undone.'
                        }}
                        modalProps={{
                            isBlocking: true,
                            styles: { main: { maxWidth: 500 } }
                        }}
                    >
                        <TextField
                            label="Comments (Optional)"
                            multiline
                            rows={3}
                            value={this.state.approveComment}
                            onChange={(e, val) => this.setState({ approveComment: val || '' })}
                            placeholder="Add optional comments..."
                        />
                        <DialogFooter>
                            <PrimaryButton
                                onClick={this._confirmApprove}
                                disabled={this.state.isSaving}
                            >
                                {this.state.isSaving ? <Spinner size={SpinnerSize.xSmall} styles={{ root: { marginRight: 8 } }} /> : null}
                                {this.state.isSaving ? "Approving..." : "Approve"}
                            </PrimaryButton>
                            <DefaultButton onClick={this._closeApproveDialog} text="Cancel" />
                        </DialogFooter>
                    </Dialog>

                    {/* Reject Dialog */}
                    <Dialog
                        hidden={!this.state.isRejectDialogOpen}
                        onDismiss={this._closeRejectDialog}
                        dialogContentProps={{
                            type: DialogType.normal,
                            title: 'Reject Document',
                            subText: 'Are you sure you want to REJECT this document? This action cannot be undone.'
                        }}
                        modalProps={{
                            isBlocking: true,
                            styles: { main: { maxWidth: 500 } }
                        }}
                    >
                        <TextField
                            label="Reason for Rejection"
                            multiline
                            rows={3}
                            value={this.state.rejectComment}
                            onChange={(e, val) => this.setState({ rejectComment: val || '', rejectCommentError: '' })}
                            placeholder="Please add comments..."
                            errorMessage={this.state.rejectCommentError}
                        />
                        <DialogFooter>
                            <PrimaryButton
                                onClick={this._submitReject}
                                disabled={(!this.state.rejectComment || this.state.rejectComment.trim().length === 0) || this.state.isSaving}
                                styles={{ root: { backgroundColor: '#d13438', borderColor: '#d13438' }, rootHovered: { backgroundColor: '#a4262c', borderColor: '#a4262c' } }}
                            >
                                {this.state.isSaving ? <Spinner size={SpinnerSize.xSmall} styles={{ root: { marginRight: 8 } }} /> : null}
                                {this.state.isSaving ? "Rejecting..." : "Reject"}
                            </PrimaryButton>
                            <DefaultButton onClick={this._closeRejectDialog} text="Cancel" disabled={this.state.isSaving} />
                        </DialogFooter>
                    </Dialog>

                    {/* Success Dialog */}
                    <Dialog
                        hidden={!this.state.isSuccessDialogOpen}
                        onDismiss={() => { /* Do nothing, must click Redirect */ }}
                        dialogContentProps={{
                            type: DialogType.normal,
                            title: 'Success',
                            subText: this.state.successMessage
                        }}
                        modalProps={{
                            isBlocking: true,
                            styles: { main: { maxWidth: 450 } }
                        }}
                    >
                        <DialogFooter>
                            <PrimaryButton onClick={() => { this._closeSuccessDialog(); window.location.href = `${TENANT_DOMAIN}/${SITE_REDIRECT}`; }} text="OK" />
                        </DialogFooter>
                    </Dialog>

                    {/* Blocking Access Dialog */}
                    <Dialog
                        hidden={!this.state.isBlockingDialogOpen}
                        onDismiss={() => { /* Do nothing, must click Redirect */ }}
                        dialogContentProps={{
                            type: DialogType.normal,
                            title: this.state.blockingDialogTitle,
                            subText: this.state.blockingDialogMessage,
                            styles: {
                                title: { color: '#a4262c' } // Dark red color for title
                            }
                        }}
                        modalProps={{
                            isBlocking: true,
                            styles: { main: { maxWidth: 450 } }
                        }}
                    >
                        <DialogFooter>
                            <PrimaryButton
                                text="OK"
                                onClick={() => {
                                    window.location.href = `${TENANT_DOMAIN}/${SITE_REDIRECT}`;
                                }}
                            />
                        </DialogFooter>
                    </Dialog>

                    {/* File Not Found Dialog */}
                    <Dialog
                        hidden={!this.state.isFileNotFoundDialogOpen}
                        onDismiss={() => { /* Blocking */ }}
                        dialogContentProps={{
                            type: DialogType.normal,
                            title: 'Document Not Found',
                            subText: 'The document file could not be found. It may have been renamed, moved or deleted. The process will be canceled.',
                            styles: { title: { color: '#a4262c' } }
                        }}
                        modalProps={{ isBlocking: true }}
                    >
                        <DialogFooter>
                            <PrimaryButton onClick={this._handleFileNotFoundCancel} text="OK" />
                        </DialogFooter>
                    </Dialog>

                    {/* Validation Error Dialog */}
                    <Dialog
                        hidden={!this.state.isValidationErrorOpen}
                        onDismiss={() => this.setState({ isValidationErrorOpen: false })}
                        dialogContentProps={{
                            type: DialogType.normal,
                            title: 'Incomplete Signature',
                            subText: 'There are unsigned fields, please complete all required signature before approving.',
                            styles: {
                                title: { color: '#d13438' } // Red title
                            }
                        }}
                        modalProps={{
                            isBlocking: true,
                            styles: { main: { maxWidth: 450 } }
                        }}
                    >
                        <DialogFooter>
                            <PrimaryButton
                                text="OK"
                                onClick={() => this.setState({ isValidationErrorOpen: false })}
                                styles={{
                                    root: { backgroundColor: '#d13438', borderColor: '#d13438' },
                                    rootHovered: { backgroundColor: '#a4262c', borderColor: '#a4262c' }
                                }}
                            />
                        </DialogFooter>
                    </Dialog>
                </div>
            </div>
        );
    }

    private _openLogDialog = (): void => {
        const fileRef = this.state.serverRelativeUrl;
        if (!fileRef) {
            console.error("No file ref available for log dialog");
            return;
        }

        const fileName = fileRef.split('/').pop() || "Document";
        const sp = getSP(this.props.context);

        const logDialog = new SignatureLogDialog(fileRef, sp, `Approval Log - ${fileName}`);
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        logDialog.show();
    }
}
