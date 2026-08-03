import * as React from 'react';
import { 
    Stack, 
    Text, 
    Icon, 
    DetailsList, 
    IColumn, 
    SelectionMode, 
    Spinner, 
    SpinnerSize, 
    MessageBar, 
    MessageBarType,
    TooltipHost
} from '@fluentui/react';
import { SPFI } from "@pnp/sp";
import { EsignDataService } from '../../../common/services/EsignDataService';
import styles from './DocumentTrackerEsign.module.scss';

export interface IProcessHistoryItemProps {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    process: any;
    sp: SPFI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    context: any;
}

export interface IProcessHistoryItemState {
    isExpanded: boolean;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    tasks: any[];
    loading: boolean;
    error: string;
}

export class ProcessHistoryItem extends React.Component<IProcessHistoryItemProps, IProcessHistoryItemState> {
    
    constructor(props: IProcessHistoryItemProps) {
        super(props);
        this.state = {
            isExpanded: false,
            tasks: [],
            loading: false,
            error: ""
        };
    }

    private _toggleExpand = (): void => {
        this.setState((prevState) => {
            const newExpanded = !prevState.isExpanded;
            // Fetch tasks if expanding and not yet loaded
            if (newExpanded && prevState.tasks.length === 0 && !prevState.loading) {
                // eslint-disable-next-line @typescript-eslint/no-floating-promises
                this._fetchTasks();
            }
            return { isExpanded: newExpanded };
        });
    }

    private _fetchTasks = async (): Promise<void> => {
        this.setState({ loading: true, error: "" });
        try {
            const tasks = await EsignDataService.getTasksByProcessId(this.props.process.Id, this.props.sp);
            
            // Transform tasks to match display requirements (similar to SignatureLogDialog)
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const formattedTasks = tasks.map((t: any) => {
                let assignees = [];
                if (Array.isArray(t.AssignedTo)) {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    assignees = t.AssignedTo.map((u: any) => ({ title: u.Title, id: u.Id }));
                } else if (t.AssignedTo) {
                    assignees = [{ title: t.AssignedTo.Title, id: t.AssignedTo.Id }];
                }

                return {
                    assignedTo: assignees,
                    step: t.StepApprover,
                    totalSteps: t.TotalApprover,
                    status: t.Status,
                    createdOn: t.Created,
                    completedOn: (t.Status === 'Approved' || t.Status === 'Rejected') ? t.Modified : null,
                    comments: t.Comments
                };
            });

            this.setState({ tasks: formattedTasks, loading: false });
        } catch (err) {
            console.error("Error fetching tasks", err);
            this.setState({ loading: false, error: "Failed to load approval history." });
        }
    }

    private _formatDate(dateStr: string): string {
        if (!dateStr) return "-";
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
        const { process } = this.props;
        const { isExpanded, loading, error, tasks } = this.state;

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
                key: 'createdOn', name: 'Created On', fieldName: 'createdOn', minWidth: 130, maxWidth: 130, isMultiline: true,
                onRender: (item) => <span>{this._formatDate(item.createdOn)}</span>
            },
            { 
                key: 'completedOn', name: 'Completed On', fieldName: 'completedOn', minWidth: 130, maxWidth: 130, isMultiline: true,
                onRender: (item) => <span>{this._formatDate(item.completedOn)}</span>
            },
            { 
                key: 'comments', name: 'Comments', fieldName: 'comments', minWidth: 100, maxWidth: 175, isMultiline: true,
                onRender: (item) => <span title={item.comments}>{item.comments || "-"}</span>
            }
        ];

        return (
            <div style={{ 
                border: '1px solid #edebe9', 
                borderRadius: '4px', 
                backgroundColor: 'white',
                marginBottom: '10px',
                boxShadow: isExpanded ? '0 2px 4px rgba(0,0,0,0.1)' : 'none'
            }}>
                {/* Header (Clickable) */}
                {/* Header (Clickable) */}
                <div 
                    className={styles.historyItemHeader}
                    style={{ 
                        backgroundColor: isExpanded ? '#f3f2f1' : 'transparent',
                        borderBottom: isExpanded ? '1px solid #edebe9' : 'none'
                    }}
                    onClick={this._toggleExpand}
                >
                    <div className={styles.historyItemHeaderLeft}>
                        <Icon iconName={isExpanded ? "ChevronDown" : "ChevronRight"} styles={{ root: { fontSize: 14, color: '#605e5c' } }} />
                        <Stack>
                            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{process.Title}</Text>
                            <Text variant="small" style={{ color: '#605e5c' }}>Created: {this._formatDate(process.Created)}</Text>
                        </Stack>
                    </div>
                    
                    <div className={styles.historyItemHeaderRight}>
                         <div style={{minWidth: 100}}>
                             <Text variant="small" style={{ fontWeight: 600, display: 'block' }}>Status</Text>
                             {this._renderStatusBadge(process.Status)}
                         </div>
                         <div style={{minWidth: 150}}>
                             <Text variant="small" style={{ fontWeight: 600, display: 'block' }}>Requestor</Text>
                             <Text variant="small">{process.Requestor ? process.Requestor.Title : "Unknown"}</Text>
                         </div>
                    </div>
                </div>

                {/* Expanded Content */}
                {isExpanded && (
                    <div style={{ padding: '15px' }}>
                         <Stack tokens={{ childrenGap: 10 }}>
                             {loading && <Spinner size={SpinnerSize.medium} label="Loading history..." />}
                             
                             {!loading && error && (
                                 <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
                             )}

                             {!loading && !error && (
                                 <div>
                                     <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>Process History</Text>
                                     <Text variant="small" style={{ display: 'block', marginBottom: 10, color: '#605e5c' }}>
                                         File: {process.FileRef0 ? process.FileRef0.split('/').pop() : 'Unknown'}
                                     </Text>
                                     <div className={styles.tableContainer}>
                                         <DetailsList
                                             items={tasks}
                                             columns={columns}
                                             selectionMode={SelectionMode.none}
                                             layoutMode={0}
                                             compact={true}
                                         />
                                     </div>
                                     {tasks.length === 0 && (
                                          <div style={{ padding: 10, textAlign: 'center', color: '#605e5c' }}>No history records found.</div>
                                     )}
                                 </div>
                             )}
                         </Stack>
                    </div>
                )}
            </div>
        );
    }
}
