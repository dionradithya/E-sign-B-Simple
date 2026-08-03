import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './EsignTaskList.module.scss';
import type { IEsignTaskListProps } from './IEsignTaskListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { getSP } from '../../../common/pnpjsConfig';
import { EsignDataService } from '../../../common/services/EsignDataService';
import { TENANT_DOMAIN, SITES_ESIGN, LIST_TASKS } from '../../../common/constants';
import { getSiteUrlFromPath } from '../../../common/utils';
import {
    DetailsList,
    IColumn,
    SelectionMode,
    Spinner,
    SpinnerSize,
    Text,
    Stack,
    MessageBar,
    MessageBarType,
    PrimaryButton
} from '@fluentui/react';

interface IProcessingRequest {
    Id: number;
    Title: string;
    Status: string;
    Created: string;
    FileRef0?: string;
    Requestor?: {
        Id: number;
        Title: string;
    };
}

const EsignTaskList: React.FC<IEsignTaskListProps> = (props) => {
    const {
        hasTeamsContext,
    } = props;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const [tasks, setTasks] = useState<any[]>([]);
    const [processingRequests, setProcessingRequests] = useState<IProcessingRequest[]>([]);
    const [loading, setLoading] = useState<boolean>(true);
    const [currentUserId, setCurrentUserId] = useState<number | null>(null);

    useEffect(() => {
        const fetchData = async (): Promise<void> => {
            try {
                const requestUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}`;
                const sp = getSP(props.context, requestUrl);

                // Parallel Fetch: Tasks, Processing Requests and Current User
                const [user, data, requests] = await Promise.all([
                    sp.web.currentUser(),
                    EsignDataService.getPendingTasksForCurrentUser(sp),
                    EsignDataService.getMyProcessingRequests(sp)
                ]);

                setCurrentUserId(user.Id);
                setTasks(data);
                setProcessingRequests(requests);

            } catch (err) {
                console.error("Error fetching tasks or user", err);
            } finally {
                setLoading(false);
            }
        };

        fetchData().catch(console.error);
    }, [props.context]);

    const handleOpenDocument = (fileRef: string, isDownloadableStr?: string | boolean, fileGuid?: string): void => {
        const origin = window.location.origin;
        // Check if fileRef starts with /
        const formattedRef = fileRef.startsWith('/') ? fileRef : `/${fileRef}`;

        const isDownloadable = isDownloadableStr !== 'false' && isDownloadableStr !== false;

        let url = "";

        if (!isDownloadable && fileGuid) {
            // Construct view-only embedded link using UniqueId (GUID)
            // Example: https://kpccoid.sharepoint.com/sites/DMS/_layouts/15/Embed.aspx?UniqueId=4cfd894a-55b4-4f36-99d6-bd59588ca120
            const siteUrl = getSiteUrlFromPath(fileRef);
            url = `${siteUrl}/_layouts/15/Embed.aspx?UniqueId=${fileGuid}`;
        } else {
            // Standard open (allows downloads) or fallback if GUID is missing
            const encodedRef = formattedRef.split('/').map(s => encodeURIComponent(s)).join('/');
            url = origin + encodedRef;
        }

        window.open(url, '_blank');
    };

    const handleOpenFolder = (fileRef: string): void => {
        const origin = window.location.origin;
        const formattedRef = fileRef.startsWith('/') ? fileRef : `/${fileRef}`;

        const lastSlashIndex = formattedRef.lastIndexOf('/');
        const folderRef = lastSlashIndex > -1 ? formattedRef.substring(0, lastSlashIndex) : formattedRef;

        // Encode segments. Special case: Spacing in folder names needs careful handling or just standard URI processing.
        // SharePoint folders usually handled well with encodeURI or split/encodeURIComponent.
        // Let's stick to the proven split/join method.
        const encodedRef = folderRef.split('/').map(s => encodeURIComponent(s)).join('/');

        window.open(origin + encodedRef, '_blank');
    };

    const getStatusBadgeStyle = (status: string): React.CSSProperties => {
        const styles: React.CSSProperties = {
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
        return styles;
    };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const getActionButtonText = (assignedTo: any[]): string => {
        if (!currentUserId || !assignedTo || assignedTo.length === 0) return "Open";

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const userIndex = assignedTo.findIndex((u: any) => u.Id === currentUserId);

        if (userIndex === 0) {
            return "Execute Task";
        } else if (userIndex > 0) {
            return "Execute Delegation";
        }

        return "Open";
    };

    const requestColumns: IColumn[] = [
        {
            key: 'no',
            name: 'No',
            fieldName: 'no',
            minWidth: 30,
            maxWidth: 30,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onRender: (item: any, index?: number) => <span>{(index || 0) + 1}</span>
        },
        {
            key: 'processCode',
            name: 'Process Code',
            fieldName: 'Title',
            minWidth: 150,
            maxWidth: 170,
            isResizable: true,
        },
        {
            key: 'requestor',
            name: 'Requestor',
            fieldName: 'requestor',
            minWidth: 100,
            maxWidth: 100,
            isResizable: true,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onRender: (item: any) => <span>{item.Requestor?.Title || "-"}</span>
        },
        {
            key: 'fileName',
            name: 'File Name',
            minWidth: 150,
            maxWidth: 250,
            isResizable: true,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onRender: (item: any) => <span>{item.FileRef0 ? item.FileRef0.split('/').pop() : "-"}</span>
        },
        {
            key: 'status',
            name: 'Status',
            fieldName: 'Status',
            minWidth: 80,
            maxWidth: 80,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onRender: (item: any) => <span style={getStatusBadgeStyle(item.Status)}>{item.Status}</span>
        },
        {
            key: 'created',
            name: 'Created',
            fieldName: 'Created',
            minWidth: 100,
            maxWidth: 100,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onRender: (item: any) => <span>{new Date(item.Created).toLocaleDateString()} {new Date(item.Created).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
        },
        {
            key: 'action',
            name: 'Action',
            minWidth: 100,
            maxWidth: 150,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onRender: (item: any) => {
                if (!item.FileRef0) return null;
                return (
                    <PrimaryButton
                        text="View Document"
                        onClick={() => handleOpenFolder(item.FileRef0)}
                        styles={{ root: { height: 26, padding: '0 12px', fontSize: '12px', backgroundColor: '#f3f2f1', color: '#323130', border: '1px solid #8a8886' } }}
                    />
                );
            }
        }
    ];

    const columns: IColumn[] = [
        {
            key: 'no',
            name: 'No',
            fieldName: 'no',
            minWidth: 30,
            maxWidth: 30,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onRender: (item: any, index?: number) => <span>{(index || 0) + 1}</span>
        },
        {
            key: 'processCode',
            name: 'Process Code',
            fieldName: 'processCode',
            minWidth: 150,
            maxWidth: 170,
            isResizable: true,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onRender: (item: any) => <span>{item.ProcessID?.Title || "-"}</span>
        },
        {
            key: 'requestor',
            name: 'Requestor',
            fieldName: 'requestor',
            minWidth: 100,
            maxWidth: 100,
            isResizable: true,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onRender: (item: any) => <span>{item.ProcessID?.Requestor?.Title || "-"}</span>
        },
        {
            key: 'taskTitle',
            name: 'Task Title',
            fieldName: 'Title',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
        },
        {
            key: 'status',
            name: 'Status',
            fieldName: 'Status',
            minWidth: 80,
            maxWidth: 80,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onRender: (item: any) => <span style={getStatusBadgeStyle(item.Status)}>{item.Status}</span>
        },
        {
            key: 'created',
            name: 'Created',
            fieldName: 'Created',
            minWidth: 100,
            maxWidth: 100,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onRender: (item: any) => <span>{new Date(item.Created).toLocaleDateString()} {new Date(item.Created).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
        },
        {
            key: 'action',
            name: 'Action',
            minWidth: 200,
            maxWidth: 250,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            onRender: (item: any) => {
                const actionText = getActionButtonText(item.AssignedTo);
                const executeUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}/Lists/${LIST_TASKS}/EditForm.aspx?ID=${item.Id}`;

                return (
                    <Stack horizontal tokens={{ childrenGap: 8 }}>
                        {item.ProcessID?.FileRef0 && (
                            <PrimaryButton
                                text="View Document"
                                onClick={() => handleOpenDocument(item.ProcessID.FileRef0, item.ProcessID.isDownloadable, item.ProcessID.GUID)}
                                styles={{ root: { height: 26, padding: '0 12px', fontSize: '12px', backgroundColor: '#f3f2f1', color: '#323130', border: '1px solid #8a8886' } }}
                            />
                        )}
                        <PrimaryButton
                            text={actionText}
                            onClick={() => window.open(executeUrl, '_blank')}
                            styles={{ root: { height: 26, padding: '0 12px', fontSize: '12px' } }}
                        />
                    </Stack>
                );
            }
        }
    ];

    return (
        <section className={`${styles.esignTaskList} ${hasTeamsContext ? styles.teams : ''}`}>
            <Stack tokens={{ childrenGap: 20 }} style={{ maxWidth: '100%', margin: '0 auto', padding: 20 }}>
                <div>
                    <Text variant="xLarge" style={{ fontWeight: 600, color: '#0078d4' }}>Welcome, {escape(props.userDisplayName)}</Text>
                    <Text block variant="medium" style={{ color: '#605e5c', marginTop: 5 }}>
                        Track your active requests and manage pending sign-off tasks.
                    </Text>
                </div>

                {loading && <Spinner size={SpinnerSize.large} label="Loading data..." />}

                {/* 1. My Processing Requests Section */}
                {!loading && (
                    <div style={{ marginBottom: 20 }}>
                        <Text variant="large" style={{ fontWeight: 600, color: '#005a9e', display: 'block', marginBottom: 10 }}>My Active Requests</Text>
                        {processingRequests.length === 0 ? (
                            <MessageBar messageBarType={MessageBarType.info}>
                                No active requests found.
                            </MessageBar>
                        ) : (
                            <div style={{ boxShadow: '0 2px 5px rgba(0,0,0,0.1)', borderRadius: 2, overflow: 'hidden' }}>
                                <DetailsList
                                    items={processingRequests}
                                    columns={requestColumns}
                                    selectionMode={SelectionMode.none}
                                    layoutMode={0} // Fixed columns
                                />
                            </div>
                        )}
                    </div>
                )}

                {/* 2. My Pending Tasks Section */}
                {!loading && (
                    <div>
                        <Text variant="large" style={{ fontWeight: 600, color: '#005a9e', display: 'block', marginBottom: 10 }}>My Pending Tasks</Text>
                        {tasks.length === 0 ? (
                            <MessageBar messageBarType={MessageBarType.info}>
                                No pending tasks found for <strong>{escape(props.userDisplayName)}</strong>.
                            </MessageBar>
                        ) : (
                            <div style={{ boxShadow: '0 2px 5px rgba(0,0,0,0.1)', borderRadius: 2, overflow: 'hidden' }}>
                                <DetailsList
                                    items={tasks}
                                    columns={columns}
                                    selectionMode={SelectionMode.none}
                                    layoutMode={0} // Fixed columns
                                />
                            </div>
                        )}
                    </div>
                )}
            </Stack>
        </section>
    );
};

export default EsignTaskList;
