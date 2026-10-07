import * as React from 'react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Stack, PrimaryButton, Label, Checkbox, IconButton, NormalPeoplePicker, IPersonaProps, Toggle } from '@fluentui/react';
import { IApprovalMapItem } from '../../../common/services/ApprovalService';
import { useTouchDrag, startTouchDrag } from './TouchDragContext';
import { ISignaturePlaceholder } from './PdfViewer';
import { getSP } from '../../../common/pnpjsConfig';
import { WhitelistService } from '../../../common/services/WhitelistService';
import { TENANT_DOMAIN, SITES_ESIGN } from '../../../common/constants';

export interface IApprover {
    id: number;
    displayName: string;
    email: string;
    jobTitle?: string; // Optional, might fetch from profile
    badgeNumber?: string; // Employee ID
    includeName: boolean;
    includeDate: boolean;
    includeBadge: boolean;
    placementMode: 'current' | 'all' | 'range' | 'selected';
    rangeFromPage?: number;
    rangeToPage?: number;
    selectedPages?: number[];
}

/**
 * Small muted helper text used under labels and options
 */
const HelpText: React.FC<{ className?: string }> = ({ children, className }) => (
    <div className={`text-muted ${className || ''}`} style={{ fontSize: 11, lineHeight: 1.4 }}>
        {children}
    </div>
);

interface ISidebarProps {
    context: WebPartContext;
    approvers: IApprover[];
    approvalMap: IApprovalMapItem[];
    placeholders: ISignaturePlaceholder[];
    onApproversChange: (approvers: IApprover[]) => void;
    onRemoveApprover: (id: number) => void;
    onSave: () => void;
    isSaving: boolean;
    fileLoaded: boolean;
    validationTriggered: boolean;
    reviewerCanDownload: boolean;
    onReviewerCanDownloadChange: (checked: boolean) => void;
    embedQrCode: boolean;
    onEmbedQrCodeChange: (checked: boolean) => void;
    placementMode: 'current' | 'all' | 'range' | 'selected';
    onPlacementModeChange: (mode: 'current' | 'all' | 'range') => void;
    onClearApproverPlaceholders: (approverId: number) => void;
    numPages: number;
    visiblePage: number;
    onPlaceAtCenter: (approver: IApprover, type: 'initial' | 'signature') => void;
}

const Sidebar: React.FC<ISidebarProps> = ({
    context,
    approvers,
    approvalMap,
    placeholders,
    onApproversChange,
    onRemoveApprover,
    onSave,
    isSaving,
    fileLoaded,
    validationTriggered,
    reviewerCanDownload,
    onReviewerCanDownloadChange,
    embedQrCode,
    onEmbedQrCodeChange,
    placementMode,
    onPlacementModeChange,
    onClearApproverPlaceholders,
    numPages,
    visiblePage,
    onPlaceAtCenter,
}): React.ReactElement => {

    // Temporary store for selected people picker items
    const [selectedPickerItems, setSelectedPickerItems] = React.useState<IPersonaProps[]>([]);

    const onResolveSuggestions = (filterText: string, currentPersonas?: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> => {
        if (!filterText) return [];

        const lowerFilter = filterText.toLowerCase();

        return approvalMap
            .filter(item =>
                item.Approver && (
                    item.Approver.Title.toLowerCase().indexOf(lowerFilter) !== -1 ||
                    item.Approver.EMail.toLowerCase().indexOf(lowerFilter) !== -1
                )
            )
            .map(item => ({
                key: item.Approver!.Id.toString(),
                text: item.Approver!.Title,
                secondaryText: item.Title, // Show Role as secondary text
                tertiaryText: item.Approver!.EMail,
                isValid: true
            }))
            .filter(persona => !currentPersonas?.some(curr => curr.text === persona.text)); // Filter out already selected in picker
    };

    const onChange = (items?: IPersonaProps[]): void => {
        setSelectedPickerItems(items || []);
    };

    const [addingApprover, setAddingApprover] = React.useState<boolean>(false);
    const [showGuide, setShowGuide] = React.useState<boolean>(true);

    const handleAddApprover = async (): Promise<void> => {
        if (!selectedPickerItems || selectedPickerItems.length === 0) return;
        setAddingApprover(true);

        try {
            // const client = await context.msGraphClientFactory.getClient('3'); // Removed to avoid delay
            const newApprovers: IApprover[] = [];

            // Process sequentially
            for (const item of selectedPickerItems) {
                const email = item.tertiaryText || "";
                if (approvers.some(a => a.email === email)) continue;

                const targetUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}`;
                const sp = getSP(context, targetUrl);

                const whitelistItem = await WhitelistService.getWhitelistItem(
                    sp,
                    email
                );

                const badgeNumber =
                    whitelistItem?.BadgeNumber || "";

                console.log("APPROVER EMAIL:", email);
                console.log("APPROVER WHITELIST ITEM:", whitelistItem);
                console.log("APPROVER BADGE FROM LIST:", badgeNumber);

                const newApprover: IApprover = {
                    id: parseInt(item.key as string) || 0,
                    displayName: item.text || "",
                    email: email,
                    jobTitle: item.secondaryText || "Approver",
                    badgeNumber: badgeNumber,
                    includeName: false,
                    includeDate: false,
                    includeBadge: false,
                    placementMode: 'current',
                    rangeFromPage: 1,
                    rangeToPage: 1,
                    selectedPages: []
                };
                newApprovers.push(newApprover);
            }

            if (newApprovers.length > 0) {
                onApproversChange([...approvers, ...newApprovers]);
                setSelectedPickerItems([]);
            }
        } catch (error) {
            console.error("Error adding approver:", error);
        } finally {
            setAddingApprover(false);
        }
    };

    // Touch drag context for mobile support
    const { setDragData, setIsDragging } = useTouchDrag();

    const onDragStart = (e: React.DragEvent<HTMLButtonElement>, approver: IApprover, type: 'initial' | 'signature'): void => {
        e.dataTransfer.setData("type", type);
        e.dataTransfer.setData("approverId", approver.id.toString());
        e.dataTransfer.setData("checklistName", approver.includeName.toString());
        e.dataTransfer.setData("checklistDate", approver.includeDate.toString());
        e.dataTransfer.setData("checklistBadge", approver.includeBadge.toString());
        e.dataTransfer.effectAllowed = "copy";
    };

    // Touch event handler for mobile - tap to select, then tap on PDF to place
    const handleTouchSelect = (approver: IApprover, type: 'initial' | 'signature'): void => {
        startTouchDrag(setDragData, setIsDragging, approver, type);
    };

    const handleApproverChange = (
        id: number,
        field: keyof IApprover,
        value: boolean
    ): void => {
        const updated = approvers.map(a => {
            if (a.id !== id) {
                return a;
            }

            if (field === "includeBadge" && !a.badgeNumber) {
                return {
                    ...a,
                    includeBadge: false
                };
            }

            return {
                ...a,
                [field]: value
            };
        });

        onApproversChange(updated);
    };

        const handlePlacementModeChange = (
            id: number,
            mode: 'current' | 'all' | 'range' | 'selected'
        ): void => {

        const updated = approvers.map(a =>
            a.id === id
                ? {
                    ...a,
                    placementMode: mode
                }
                : a
        );

        onApproversChange(updated);
    };

    const handleMoveUp = (index: number): void => {
        if (index === 0) return;
        const updated = [...approvers];
        [updated[index], updated[index - 1]] = [updated[index - 1], updated[index]];
        onApproversChange(updated);
    };

    const handleMoveDown = (index: number): void => {
        if (index === approvers.length - 1) return;
        const updated = [...approvers];
        [updated[index], updated[index + 1]] = [updated[index + 1], updated[index]];
        onApproversChange(updated);
    };



    return (
        <div className="p-3 bg-white border-start h-100" style={{ overflowY: 'auto' }}>
            <div className="d-flex justify-content-between align-items-center mb-2">
                <h5 className="mb-0">Workflow Setup</h5>
                <button
                    type="button"
                    className="btn btn-link btn-sm p-0"
                    style={{ fontSize: 12 }}
                    onClick={() => setShowGuide(!showGuide)}
                    aria-expanded={showGuide}
                >
                    {showGuide ? 'Hide guide' : 'Show guide'}
                </button>
            </div>

            {showGuide && (
                <div className="mb-3 p-2 rounded" style={{ backgroundColor: 'rgba(0, 120, 212, 0.06)', border: '1px solid rgba(0, 120, 212, 0.25)', fontSize: 12 }}>
                    <div className="fw-bold mb-1">How to set up a signing request</div>
                    <ol className="mb-0 ps-3">
                        <li>Add the approvers who need to sign, in signing order.</li>
                        <li>For each approver, choose the details to show and the Placement Mode.</li>
                        <li>Click <b>Initial</b> or <b>Signature</b> to add a box, then drag it to the right spot on the PDF.</li>
                        <li>Click <b>Process Workflow</b> to send the request.</li>
                    </ol>
                </div>
            )}

            <div className="mb-4">
                <Label>1. Add Approvers</Label>
                <HelpText className="mb-2">
                    Only people registered in the Approval Map can be added. You can select more than one person, then click Add to Workflow.
                </HelpText>

                <Label>Search Approver</Label>
                <NormalPeoplePicker
                    onResolveSuggestions={onResolveSuggestions}
                    getTextFromItem={(persona: IPersonaProps) => persona.text || ""}
                    pickerSuggestionsProps={{
                        suggestionsHeaderText: 'Suggested Approvers',
                        noResultsFoundText: 'No approvers found in Approval Map',
                    }}
                    selectedItems={selectedPickerItems}
                    onChange={onChange}
                    inputProps={{ placeholder: 'Search by name or role...' }}
                />
                <div className="mt-2">
                    <PrimaryButton
                        text={addingApprover ? "Adding..." : "Add to Workflow"}
                        onClick={handleAddApprover}
                        disabled={selectedPickerItems.length === 0 || addingApprover}
                        className="w-100"
                    />
                </div>
            </div>

            <div className="mb-4">
                <Label>2. Manage Approvers & Fields</Label>
                <HelpText className="mb-2">
                    Approvers sign in the order listed. Use the arrows to change the order. Every approver needs at least one Initial or Signature box.
                </HelpText>
                <Stack tokens={{ childrenGap: 12 }}>
                    {approvers.map((approver, index) => {
                        const sigCount = placeholders.filter(p => p.approverId === approver.id && p.type === 'signature').length;
                        const initCount = placeholders.filter(p => p.approverId === approver.id && p.type === 'initial').length;
                        const hasBox = sigCount + initCount > 0;
                        const isError = validationTriggered && !hasBox;

                        return (
                            <div key={approver.email} className={`p-3 bg-white border rounded shadow-sm ${isError ? 'border-danger' : ''}`}>
                                {/* Error Message */}
                                {isError && (
                                    <div className="text-danger mb-2" style={{ fontSize: '11px', fontWeight: 600 }}>
                                        Box Specimen belum di tambahkan
                                    </div>
                                )}

                                {/* Header */}
                                <div className="d-flex justify-content-between align-items-start mb-2 border-bottom pb-2">
                                    <div>
                                        <div className="fw-bold text-truncate" title={approver.displayName}>
                                            {index + 1}. {approver.displayName}
                                        </div>
                                        <div className="text-muted small">
                                            (Sig: {sigCount}, Init: {initCount})
                                        </div>
                                    </div>

                                    <div className="d-flex align-items-center">
                                        <IconButton
                                            iconProps={{ iconName: 'ChevronUp' }}
                                            title="Move Up"
                                            disabled={index === 0}
                                            onClick={() => handleMoveUp(index)}
                                            styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 12 } }}
                                        />
                                        <IconButton
                                            iconProps={{ iconName: 'ChevronDown' }}
                                            title="Move Down"
                                            disabled={index === approvers.length - 1}
                                            onClick={() => handleMoveDown(index)}
                                            styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 12 } }}
                                        />
                                        <IconButton
                                            iconProps={{ iconName: 'Delete' }}
                                            title="Remove"
                                            ariaLabel="Remove"
                                            onClick={() => onRemoveApprover(approver.id)}
                                            styles={{ root: { height: 24, width: 24, marginLeft: 4 }, icon: { fontSize: 12, color: 'red' } }}
                                        />
                                    </div>
                                </div>

                                {/* Checkboxes */}
                                <div className="mb-2">
                                    <HelpText className="mb-1">Shown with the signature when this approver signs:</HelpText>
                                    <Checkbox
                                        label="Include Automatic Full Name"
                                        checked={approver.includeName}
                                        onChange={(ev, checked) => handleApproverChange(approver.id, 'includeName', !!checked)}
                                        styles={{ root: { marginBottom: 4 }, label: { fontSize: 12 } }}
                                    />
                                    <Checkbox
                                        label="Include Date Time of Sign"
                                        checked={approver.includeDate}
                                        onChange={(ev, checked) => handleApproverChange(approver.id, 'includeDate', !!checked)}
                                        styles={{ root: { marginBottom: 4 }, label: { fontSize: 12 } }}
                                    />
                                    <Checkbox
                                        label="Include Badge Number"
                                        checked={!!approver.badgeNumber && approver.includeBadge}
                                        disabled={!approver.badgeNumber}
                                        onChange={(ev, checked) =>
                                            handleApproverChange(
                                                approver.id,
                                                'includeBadge',
                                                !!checked
                                            )
                                        }
                                        styles={{
                                            root: { marginBottom: 4 },
                                            label: { fontSize: 12 }
                                        }}
                                    />
                                    {!approver.badgeNumber && (
                                        <HelpText className="ms-4">No badge number registered for this approver.</HelpText>
                                    )}

                                    <div className="mt-2 mb-3">
                                        <Label>Placement Mode</Label>
                                        <HelpText className="mb-1">
                                            Choose this before clicking Initial or Signature. Changing it later does not affect boxes already added.
                                        </HelpText>

                                        <div className="form-check">
                                            <input
                                                className="form-check-input"
                                                type="radio"
                                                name={`placement-${approver.id}`}
                                                checked={approver.placementMode === 'current'}
                                                onChange={() =>
                                                    handlePlacementModeChange(
                                                        approver.id,
                                                        'current'
                                                    )
                                                }
                                            />
                                            <label className="form-check-label">
                                                Current Page
                                            </label>
                                            <HelpText>
                                                Adds one box on the page you are viewing now{numPages ? ` (page ${visiblePage} of ${numPages})` : ''}. Scroll the PDF to choose another page.
                                            </HelpText>
                                        </div>

                                        <div className="form-check">
                                            <input
                                                className="form-check-input"
                                                type="radio"
                                                name={`placement-${approver.id}`}
                                                checked={approver.placementMode === 'all'}
                                                onChange={() =>
                                                    handlePlacementModeChange(
                                                        approver.id,
                                                        'all'
                                                    )
                                                }
                                            />
                                            <label className="form-check-label">
                                                All Pages
                                            </label>
                                            <HelpText>
                                                Adds the box on every page{numPages ? ` (${numPages} pages)` : ''}. Moving or resizing one box updates it on all pages.
                                            </HelpText>
                                        </div>

                                        <div className="form-check">
                                            <input
                                                className="form-check-input"
                                                type="radio"
                                                name={`placement-${approver.id}`}
                                                checked={approver.placementMode === 'range'}
                                                onChange={() =>
                                                    handlePlacementModeChange(
                                                        approver.id,
                                                        'range'
                                                    )
                                                }
                                            />
                                            <label className="form-check-label">
                                                Page Range
                                            </label>
                                            <HelpText>
                                                Adds the box on each page from the From page to the To page, for example pages 2 to 5. Moving or resizing one box updates them all.
                                            </HelpText>
                                        </div>

                                        <div className="form-check">
                                            <input
                                                className="form-check-input"
                                                type="radio"
                                                name={`placement-${approver.id}`}
                                                checked={
                                                    approver.placementMode ===
                                                    'selected'
                                                }
                                                onChange={() =>
                                                    handlePlacementModeChange(
                                                        approver.id,
                                                        'selected'
                                                    )
                                                }
                                            />
                                            <label className="form-check-label">
                                                Selected Pages
                                            </label>
                                            <HelpText>
                                                Adds the box only on the pages you tick, for example pages 1, 3 and 7. Moving or resizing one box updates them all.
                                            </HelpText>
                                        </div>

                                        {
                                            approver.placementMode === 'selected' && (
                                                <div
                                                    className="mt-2 border rounded p-2"
                                                    style={{
                                                        maxHeight: '150px',
                                                        overflowY: 'auto'
                                                    }}
                                                >
                                                    {
                                                        Array.from(
                                                            { length: numPages },
                                                            (_, index) => {

                                                                const page =
                                                                    index + 1;

                                                                return (
                                                                    <div
                                                                        key={page}
                                                                        className="form-check"
                                                                    >
                                                                        <input
                                                                            className="form-check-input"
                                                                            type="checkbox"

                                                                            checked={
                                                                                approver.selectedPages?.includes(page) || false
                                                                            }

                                                                            onChange={(e) => {

                                                                                const updated = approvers.map(a => {

                                                                                    if (a.id !== approver.id) {
                                                                                        return a;
                                                                                    }

                                                                                    const currentPages =
                                                                                        a.selectedPages || [];

                                                                                    return {
                                                                                        ...a,

                                                                                        selectedPages:
                                                                                            e.target.checked
                                                                                                ? [
                                                                                                    ...currentPages,
                                                                                                    page
                                                                                                ]
                                                                                                : currentPages.filter(
                                                                                                    p => p !== page
                                                                                                )
                                                                                    };
                                                                                });

                                                                                onApproversChange(updated);
                                                                            }}
                                                                        />

                                                                        <label
                                                                            className="form-check-label"
                                                                        >
                                                                            Page {page}
                                                                        </label>
                                                                    </div>
                                                                );
                                                            }
                                                        )
                                                    }
                                                    <div className="small text-primary mt-2">
                                                        Selected:
                                                        {
                                                            approver.selectedPages?.join(', ')
                                                            || ' None'
                                                        }
                                                    </div>
                                                </div>
                                            )
                                        }

                                        {
                                            approver.placementMode === 'range' && (
                                                <div className="mt-2">

                                                    <label className="form-label">
                                                        From
                                                    </label>

                                                    <input
                                                        type="number"
                                                        min={1}
                                                        value={approver.rangeFromPage || 1}
                                                        className="form-control mb-2"
                                                        onChange={(e) => {

                                                            const updated =
                                                                approvers.map(a =>
                                                                    a.id === approver.id
                                                                        ? {
                                                                            ...a,
                                                                            rangeFromPage:
                                                                                Number(
                                                                                    e.target.value
                                                                                )
                                                                        }
                                                                        : a
                                                                );

                                                            onApproversChange(updated);
                                                        }}
                                                    />

                                                    <label className="form-label">
                                                        To
                                                    </label>

                                                    <input
                                                        type="number"
                                                        min={1}
                                                        value={approver.rangeToPage || 1}
                                                        className="form-control"
                                                        onChange={(e) => {

                                                            const updated =
                                                                approvers.map(a =>
                                                                    a.id === approver.id
                                                                        ? {
                                                                            ...a,
                                                                            rangeToPage:
                                                                                Number(
                                                                                    e.target.value
                                                                                )
                                                                        }
                                                                        : a
                                                                );

                                                            onApproversChange(updated);
                                                        }}
                                                    />

                                                </div>
                                            )
                                        }
                                                                            </div>
                                </div>

                                {/* Drag Buttons - tap on mobile, drag on desktop */}
                                <HelpText className="mb-2">
                                    <b>Initial</b> adds a box for the approver&apos;s initials (paraf). <b>Signature</b> adds a box for the full signature. Each click adds a box in the middle of the page, based on the Placement Mode above.
                                    Drag a box to move it, drag its bottom-right corner to resize it, and click × to remove it from that page.
                                    <b> Clear</b> removes all boxes for this approver.
                                </HelpText>
                                <div className="d-flex flex-column gap-2">
                                    <div className="d-flex gap-2">
                                        <button
                                            onClick={() => onPlaceAtCenter(approver, 'initial')}
                                            type="button"
                                            className="btn btn-sm btn-outline-primary flex-grow-1"
                                            style={{ touchAction: 'manipulation' }}
                                        >
                                            Initial
                                        </button>

                                        <button
                                            onClick={() => onPlaceAtCenter(approver, 'signature')}
                                            type="button"
                                            className="btn btn-sm btn-outline-success flex-grow-1"
                                            style={{ touchAction: 'manipulation' }}
                                        >
                                            Signature
                                        </button>
                                    </div>

                                    <button
                                        type="button"
                                        className="btn btn-sm btn-outline-danger w-100"
                                        onClick={() => onClearApproverPlaceholders(approver.id)}
                                    >
                                        Clear
                                    </button>
                                </div>
                            </div>
                        );
                    })}

                    {approvers.length === 0 && <div className="text-muted small fst-italic text-center py-3">No approvers added yet. Add someone above!</div>}
                </Stack>
            </div>

            <div className="mt-auto pt-3 border-top">
                <div className="mb-3">
                    <Toggle
                        label="Reviewers Can Download"
                        checked={reviewerCanDownload}
                        onChange={(ev, checked) => onReviewerCanDownloadChange(!!checked)}
                        onText="Yes"
                        offText="No"
                        styles={{ root: { marginBottom: 0 }, label: { fontSize: 12, fontWeight: 600 } }}
                    />
                    <HelpText className="mb-2">
                        Yes: approvers can download the PDF while reviewing. No: approvers can only view it.
                    </HelpText>
                    <Toggle
                        label="Embed QR Code on Document"
                        checked={embedQrCode}
                        onChange={(ev, checked) => onEmbedQrCodeChange(!!checked)}
                        onText="Yes"
                        offText="No"
                        styles={{ root: { marginBottom: 0 }, label: { fontSize: 12, fontWeight: 600 } }}
                    />
                    <p className="text-muted small mb-0" style={{ fontSize: 11, marginTop: 2 }}>QR code will be placed at the bottom-right corner of last page.</p>
                </div>
                <PrimaryButton
                    text={isSaving ? "Processing..." : "Process Workflow"}
                    onClick={onSave}
                    disabled={isSaving || !fileLoaded || approvers.length === 0}
                    className="w-100 mb-2"
                />
            </div>
        </div>
    );
};

export default Sidebar;
