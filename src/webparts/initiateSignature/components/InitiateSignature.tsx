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

    useEffect(() => {
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
    }, [fileRef, context]);

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

                const newPlaceholders: ISignaturePlaceholder[] = [];

                for (let p = 1; p <= numPages; p++) {

                    newPlaceholders.push({
                        id: `${Date.now()}-${p}-${Math.random()}`,
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

                setPlaceholders([
                    ...placeholders,
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

                const newPlaceholders: ISignaturePlaceholder[] = [];

                for (
                    let p = fromPage;
                    p <= toPage;
                    p++
                ) {

                    newPlaceholders.push({
                        id: `${Date.now()}-${p}-${Math.random()}`,
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

                setPlaceholders([
                    ...placeholders,
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

    const updatePlaceholder = (id: string, updates: Partial<ISignaturePlaceholder>): void => {
        setPlaceholders(placeholders.map(p => p.id === id ? { ...p, ...updates } : p));
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
            let finalBlobToHash = fileBlob;
            if (embedQrCode) {
                try {
                    const qrContent = `${TENANT_DOMAIN}/${SITE_REDIRECT}/SitePages/Document-Tracker.aspx?code=${esinTitle}`;
                    const qrImageBytes = await generateQrCodeImageBytes(qrContent);

                    // 2. Manipulate PDF
                    const arrayBuffer = await fileBlob.arrayBuffer();
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
                    await spDoc.web.getFileByServerRelativePath(fileMetadata.fileRef).setContentChunked(finalBlobToHash as any, {
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
            const processSignature = approvers.map(approver => {
                const appPlaceholders = placeholders.filter(p => p.approverId === approver.id);

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
                fileRef: fileMetadata.fileRef,
                status: "Processing",
                ProcessSignature: processSignature
            };

            const jsonString = JSON.stringify(payload);
            const emailString = approvers.map(a => a.email).join(';');

            // Fetch Current User to get ID
            const currentUser = await spList.web.currentUser();

            // Save to LIST_PROCESS list
            await spList.web.lists.getByTitle(LIST_PROCESS).items.add({
                Title: payload.title,
                HashHex: payload.hashHex,
                FileRef0: payload.fileRef,
                RequestorId: currentUser.Id,
                Status: payload.status,
                ProcessData: jsonString,
                UserNotif: emailString,
                isDownloadable: reviewerCanDownload.toString()
            });

            // Save to LIST_ACTIVITY_LOG
            await spList.web.lists.getByTitle(LIST_ACTIVITY_LOG).items.add({
                Activity: "Initiated Signature Request",
                UserIdId: currentUser.Id,
                ProcessTitle: payload.title
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
                        onSave={handleSaveWorkflow}
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
            </div>
        </TouchDragProvider>
    );
};

export default InitiateSignature;
