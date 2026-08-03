import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
    DialogContent,
    Spinner,
    SpinnerSize,
    MessageBar,
    MessageBarType,
    DetailsList,
    IColumn,
    SelectionMode,
    Text,
    TooltipHost,
    Icon,
    Stack,
    IconButton
} from '@fluentui/react';
import { SPFI } from "@pnp/sp";
import { EsignDataService } from '../../../common/services/EsignDataService';

interface ISignatureLogDialogContentProps {
    fileRef: string;
    sp: SPFI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    context: any; // Added context prop
    close: () => void;
    title?: string;
}

interface ISignatureLogDialogContentState {
    loading: boolean;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    data: any;
    error: string;
}

export class SignatureLogDialogContent extends React.Component<ISignatureLogDialogContentProps, ISignatureLogDialogContentState> {

    constructor(props: ISignatureLogDialogContentProps) {
        super(props);
        this.state = {
            loading: true,
            data: null,
            error: ""
        };
    }

    public componentDidMount(): void {
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this._fetchData();
    }

    private _fetchData = async (): Promise<void> => {
        try {
            const data = await EsignDataService.getSignatureLogData(this.props.fileRef, this.props.sp, this.props.context);
            if (data) {
                this.setState({ loading: false, data: data });
            } else {
                this.setState({ loading: false, error: "No approval process found for this document." });
            }
        } catch (err) {
            console.error(err);
            this.setState({ loading: false, error: "Failed to load signature log." });
        }
    }

    private _formatDate(dateStr: string): string {
        if (!dateStr) return "";
        const date = new Date(dateStr);
        return date.toLocaleString('en-GB', { day: '2-digit', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    }

    private _renderStatusBadge(status: string): JSX.Element {
        const styles = {
            backgroundColor: '#f3f2f1',
            color: '#201f1e',
            padding: '2px 8px',
            borderRadius: '12px',
            fontSize: '12px',
            fontWeight: 600,
            display: 'inline-block'
        };

        switch (status) {
            case 'Processing':
                styles.backgroundColor = 'rgba(0, 120, 212, 0.1)'; // Light Blue
                styles.color = '#005a9e';
                break;
            case 'Approved':
                styles.backgroundColor = 'rgba(16, 124, 16, 0.1)'; // Light Green
                styles.color = '#107c10';
                break;
            case 'Rejected':
            case 'Canceled':
                styles.backgroundColor = 'rgba(168, 0, 0, 0.1)'; // Light Red
                styles.color = '#a80000';
                break;
            case 'Expired':
                styles.backgroundColor = 'rgba(255, 170, 0, 0.1)'; // Light Orange
                styles.color = '#da3b01';
                break;
            case 'Pending':
                styles.backgroundColor = 'rgba(0, 120, 212, 0.1)'; // Light Blue
                styles.color = '#106fb8d5';
                break;
        }

        return <span style={styles}>{status}</span>;
    }

    public render(): JSX.Element {
        const { loading, error, data } = this.state;

        const columns: IColumn[] = [
            {
                key: 'assignedTo', name: 'Assigned To', fieldName: 'assignedTo', minWidth: 150, maxWidth: 150, isMultiline: true,
                onRender: (item) => {
                    const assignees = item.assignedTo || [];
                    if (assignees.length === 0) return <span>Unassigned</span>;

                    const firstAssignee = assignees[0].title;
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const others = `Delegation: ${assignees.slice(1).map((a: any) => a.title).join(', ')}`;

                    if (assignees.length > 1) {
                        return (
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 5 }}>
                                <span>{firstAssignee}</span>
                                <TooltipHost content={others}>
                                    <Icon iconName="People" style={{ fontSize: '14px', cursor: 'pointer', color: '#0078d4' }} />
                                </TooltipHost>
                            </Stack>
                        );
                    } else {
                        return <span>{firstAssignee}</span>;
                    }
                }
            },
            {
                key: 'step', name: 'Step', fieldName: 'step', minWidth: 40, maxWidth: 40,
                onRender: (item) => {
                    if (item.step && item.totalSteps) {
                        return <span style={{ fontWeight: 600 }}>{item.step} of {item.totalSteps}</span>;
                    }
                    return <span>-</span>;
                }
            },
            {
                key: 'status', name: 'Status', fieldName: 'status', minWidth: 80, maxWidth: 80,
                onRender: (item) => this._renderStatusBadge(item.status)
            },
            {
                key: 'createdOn', name: 'Created on', fieldName: 'createdOn', minWidth: 130, maxWidth: 130, isMultiline: true,
                onRender: (item) => <span>{this._formatDate(item.createdOn)}</span>
            },
            {
                key: 'completedOn', name: 'Completed On', fieldName: 'completedOn', minWidth: 130, maxWidth: 130, isMultiline: true,
                onRender: (item) => <span>{this._formatDate(item.completedOn)}</span>
            },
            { key: 'comments', name: 'Comments', fieldName: 'comments', minWidth: 100, maxWidth: 200, isMultiline: true }
        ];

        return (
            <DialogContent
                title={this.props.title || "Approval Log"}
                onDismiss={this.props.close}
                showCloseButton={true}
                styles={{
                    content: { width: '900px', maxWidth: '95vw' },
                    title: {
                        fontSize: '18px',
                        overflow: 'hidden',
                        whiteSpace: 'nowrap',
                        textOverflow: 'ellipsis',
                        display: 'block'
                    }
                }} // Widen the dialog
            >
                {loading && <Spinner size={SpinnerSize.large} label="Loading Log..." />}

                {!loading && error && (
                    <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
                )}

                {!loading && !error && data && (
                    <Stack tokens={{ childrenGap: 0 }}>
                        {/* Header Info */}
                        {data.isHashMatch === false && (
                            <MessageBar messageBarType={MessageBarType.severeWarning} styles={{ root: { marginBottom: 10 } }}>
                                <strong>File Mismatch:</strong> The file does not match the approval record. The document may have been modified.
                            </MessageBar>
                        )}
                        {data.isFilenameChanged && (
                            <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: 10 } }}>
                                <strong>Filename Changed:</strong> The document content matches a record with a different name.
                            </MessageBar>
                        )}
                        <div style={{ textAlign: 'center', marginBottom: '10px' }}>
                            <Stack horizontal horizontalAlign="center" verticalAlign="center" tokens={{ childrenGap: 5 }}>
                                <Text variant="xLarge" style={{ fontWeight: 700, color: '#0078d4' }}>
                                    {data.process.Title}
                                </Text>
                                <TooltipHost content="Copy Code">
                                    <IconButton
                                        iconProps={{ iconName: 'Copy' }}
                                        styles={{ root: { color: '#0078d4', height: 24 } }}
                                        onClick={() => {
                                            navigator.clipboard.writeText(data.process.Title).catch(console.error);
                                        }}
                                    />
                                </TooltipHost>
                            </Stack>
                        </div>
                        <div style={{
                            display: 'grid',
                            gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))',
                            gap: '10px 30px',
                            backgroundColor: '#f3f2f1',
                            padding: '15px',
                            borderRadius: '5px',
                            marginBottom: 0
                        }}>
                            <div>
                                <Text block variant="small" style={{ fontWeight: 600 }}>Origin Filename</Text>
                                <Text block style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={data.process.OriginFilename}>
                                    {data.process.OriginFilename}
                                </Text>
                            </div>
                            <div>
                                <Text block variant="small" style={{ fontWeight: 600 }}>Initiated On</Text>
                                <Text block>{this._formatDate(data.process.Created)}</Text>
                            </div>
                            <div>
                                <Text block variant="small" style={{ fontWeight: 600 }}>Status</Text>
                                {this._renderStatusBadge(data.process.Status)}
                            </div>
                            <div>
                                <Text block variant="small" style={{ fontWeight: 600 }}>Initiated By</Text>
                                <Text block>{data.process.Requestor ? data.process.Requestor.Title : "Unknown"}</Text>
                            </div>
                        </div>
                        {/* Tasks Table */}
                        <div>
                            {/* Removed minWidth to let DetailsList adapt, reducing scroll chance on desktop */}
                        <div style={{ maxHeight: '40vh', overflowY: 'auto', overflowX: 'auto', width: '100%' }}>
                            <DetailsList
                                items={data.tasks || []}
                                columns={columns}
                                selectionMode={SelectionMode.none}
                                layoutMode={0} // Fixed columns
                                compact={true}
                            />
                            {(!data.tasks || data.tasks.length === 0) && (
                                <div style={{ textAlign: 'center', padding: '10px', color: '#666' }}>No approval history found.</div>
                            )}
                        </div>
                        </div>
                    </Stack>
                )}
            </DialogContent>
        );
    }
}

export default class SignatureLogDialog extends BaseDialog {
    public fileRef: string;
    public sp: SPFI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public context: any; // Added context
    public title: string;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    constructor(fileRef: string, sp: SPFI, context: any, title: string = "Approval Log") {
        super();
        this.fileRef = fileRef;
        this.sp = sp;
        this.context = context; // Store context
        this.title = title;
    }

    public render(): void {
        ReactDOM.render(
            <SignatureLogDialogContent
                fileRef={this.fileRef}
                sp={this.sp}
                context={this.context} // Pass context
                close={this._close}
                title={this.title}
            />,
            this.domElement
        );
    }

    public getConfig(): IDialogConfiguration {
        return { isBlocking: false };
    }

    private _close = (): void => {
        this.close().catch(console.error);
    }

    protected onAfterClose(): void {
        super.onAfterClose();
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
}
