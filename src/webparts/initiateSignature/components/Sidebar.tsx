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
    placementMode: 'current' | 'all';
}

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
    placementMode: 'current' | 'all';
    onPlacementModeChange: (mode: 'current' | 'all') => void;
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
                    placementMode: 'current'
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
        mode: 'current' | 'all'
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
            <h5 className="mb-3">Workflow Setup</h5>

            <div className="mb-4">
                <Label>1. Add Approvers</Label>

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

            {/* <div className="mb-4">
                <Label>Placement Mode</Label>

                <div className="form-check">
                    <input
                        className="form-check-input"
                        type="radio"
                        name="placementMode"
                        checked={placementMode === 'current'}
                        onChange={() => onPlacementModeChange('current')}
                    />
                    <label className="form-check-label">
                        Current Page
                    </label>
                </div>

                <div className="form-check">
                    <input
                        className="form-check-input"
                        type="radio"
                        name="placementMode"
                        checked={placementMode === 'all'}
                        onChange={() => onPlacementModeChange('all')}
                    />
                    <label className="form-check-label">
                        All Pages
                    </label>
                </div>
            </div> */}

            <div className="mb-4">
                <Label>2. Manage Approvers & Fields</Label>
                <p className="text-muted small">Configure settings and drag fields onto the PDF.</p>
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

                                    <div className="mt-2 mb-3">
                                        <Label>Placement Mode</Label>

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
                                        </div>
                                    </div>
                                </div>

                                {/* Drag Buttons - tap on mobile, drag on desktop */}
                                <div className="d-flex gap-2">
                                    <button
                                        draggable
                                        onDragStart={(e) => onDragStart(e, approver, 'initial')}
                                        onTouchStart={() => handleTouchSelect(approver, 'initial')}
                                        type="button"
                                        className="btn btn-sm btn-outline-primary flex-grow-1"
                                        style={{ touchAction: 'manipulation', userSelect: 'none', WebkitUserSelect: 'none' }}
                                    >
                                        Initial
                                    </button>
                                    <button
                                        draggable
                                        onDragStart={(e) => onDragStart(e, approver, 'signature')}
                                        onTouchStart={() => handleTouchSelect(approver, 'signature')}
                                        type="button"
                                        className="btn btn-sm btn-outline-success flex-grow-1"
                                        style={{ touchAction: 'manipulation', userSelect: 'none', WebkitUserSelect: 'none' }}
                                    >
                                        Signature
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
                        styles={{ root: { marginBottom: 8 }, label: { fontSize: 12, fontWeight: 600 } }}
                    />
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
