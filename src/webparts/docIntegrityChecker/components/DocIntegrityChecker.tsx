import * as React from 'react';
import { useState } from 'react';
import styles from './DocIntegrityChecker.module.scss';
import type { IDocIntegrityCheckerProps } from './IDocIntegrityCheckerProps';
import { 
    Stack, 
    TextField, 
    PrimaryButton, 
    MessageBar, 
    MessageBarType, 
    Text,
    Spinner,
    SpinnerSize,
    FontIcon,
    mergeStyles,
    DetailsList,
    IColumn,
    SelectionMode,
    TooltipHost,
    Icon
} from '@fluentui/react';
import { getSP } from '../../../common/pnpjsConfig';
import { EsignDataService } from '../../../common/services/EsignDataService';
import { generatePDFHash } from '../../../common/utils/HashHelper';
import { TENANT_DOMAIN, SITES_ESIGN } from '../../../common/constants';

const iconClass = mergeStyles({
    fontSize: 50,
    height: 50,
    width: 50,
    margin: '0 auto 10px auto',
    color: '#0078d4'
});

const DocIntegrityChecker: React.FC<IDocIntegrityCheckerProps> = (props) => {
  const {
    hasTeamsContext,
  } = props;

  const [processCode, setProcessCode] = useState<string>('');
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [isDragOver, setIsDragOver] = useState<boolean>(false);
  const [result, setResult] = useState<{ type: MessageBarType, message: string } | null>(null);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [processItem, setProcessItem] = useState<any | null>(null);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const [tasks, setTasks] = useState<any[]>([]);
  
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
      const files = event.target.files;
      if (files && files.length > 0) {
          if (files[0].type === 'application/pdf') {
            setFile(files[0]);
            setResult(null);
            setProcessItem(null);
            setTasks([]);
          } else {
            setResult({ type: MessageBarType.error, message: 'Invalid file type. Please upload a PDF file.' });
          }
      }
  };

  const onDragOver = (event: React.DragEvent<HTMLDivElement>): void => {
      event.preventDefault();
      setIsDragOver(true);
  };

  const onDragLeave = (event: React.DragEvent<HTMLDivElement>): void => {
      event.preventDefault();
      setIsDragOver(false);
  };

  const onDrop = (event: React.DragEvent<HTMLDivElement>): void => {
      event.preventDefault();
      setIsDragOver(false);
      const files = event.dataTransfer.files;
      if (files && files.length > 0) {
          if (files[0].type === 'application/pdf') {
             setFile(files[0]);
             setResult(null);
             setProcessItem(null);
             setTasks([]);
          } else {
             setResult({ type: MessageBarType.error, message: 'Invalid file type. Please upload a PDF file.' });
          }
      }
  };

  const handleClickUpload = (): void => {
      if (fileInputRef.current) {
          fileInputRef.current.click();
      }
  };

  // Helpers for DetailsList
  const formatDate = (dateStr: string): string => {
      if (!dateStr) return "-";
      const date = new Date(dateStr);
      return date.toLocaleString('en-GB', { day: '2-digit', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  };

  const renderStatusBadge = (status: string): JSX.Element => {
      const badgeStyles = {
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
              badgeStyles.backgroundColor = 'rgba(0, 120, 212, 0.1)'; 
              badgeStyles.color = '#005a9e';
              break;
          case 'Approved':
              badgeStyles.backgroundColor = 'rgba(16, 124, 16, 0.1)'; 
              badgeStyles.color = '#107c10';
              break;
          case 'Rejected':
          case 'Canceled':
              badgeStyles.backgroundColor = 'rgba(168, 0, 0, 0.1)'; 
              badgeStyles.color = '#a80000';
              break;
          case 'Expired':
              badgeStyles.backgroundColor = 'rgba(255, 170, 0, 0.1)'; 
              badgeStyles.color = '#da3b01';
              break;
          case 'Pending':
              badgeStyles.backgroundColor = 'rgba(0, 120, 212, 0.1)'; 
              badgeStyles.color = '#106fb8d5';
              break;
      }
      return <span style={badgeStyles}>{status}</span>;
  };

  const columns: IColumn[] = [
      { 
          key: 'assignedTo', name: 'Assigned To', fieldName: 'assignedTo', minWidth: 140, maxWidth: 140, isMultiline: true,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          onRender: (item: any) => {
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
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          onRender: (item: any) => {
              if (item.step && item.totalSteps) {
                  return <span style={{ fontWeight: 600 }}>{item.step} of {item.totalSteps}</span>;
              }
              return <span>-</span>;
          }
      },
      { 
          key: 'status', name: 'Status', fieldName: 'status', minWidth: 80, maxWidth: 80, 
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          onRender: (item: any) => renderStatusBadge(item.status)
      },
      { 
          key: 'createdOn', name: 'Created On', fieldName: 'createdOn', minWidth: 120, maxWidth: 120, isMultiline: true,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          onRender: (item: any) => <span>{formatDate(item.createdOn)}</span>
      },
      { 
          key: 'completedOn', name: 'Completed On', fieldName: 'completedOn', minWidth: 120, maxWidth: 120, isMultiline: true,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          onRender: (item: any) => <span>{formatDate(item.completedOn)}</span>
      },
      { 
          key: 'comments', name: 'Comments', fieldName: 'comments', minWidth: 80, maxWidth: 130, isMultiline: true,
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          onRender: (item: any) => <span title={item.comments}>{item.comments || "-"}</span>
      }
  ];

  const handleVerify = async (): Promise<void> => {
      // Only file is required now
      if (!file) {
          setResult({ type: MessageBarType.error, message: 'Please upload a PDF file to verify.' });
          return;
      }

      setLoading(true);
      setResult(null);
      setProcessItem(null);
      setTasks([]);

      try {
          // 1. Setup PnP Context
          const requestUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}`;
          const sp = getSP(props.context, requestUrl);

          // 2. Generate Hash of Uploaded File
          const generatedHash = await generatePDFHash(file);

          // 3. Search Strategy: Process Code (if provided) OR Hash
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          let foundProcess: any = null;
          let storedHash: string | undefined = undefined;

          if (processCode.trim()) {
              // If process code provided, search by code first
              storedHash = await EsignDataService.getProcessHashByCode(processCode.trim(), sp);
              
              if (!storedHash) {
                  setResult({ 
                      type: MessageBarType.error, 
                      message: `No process found with code: ${processCode}` 
                  });
                  setLoading(false);
                  return;
              }

              // Compare with generated hash
              if (generatedHash === storedHash) {
                  foundProcess = await EsignDataService.getProcessByUniqueCode(processCode.trim(), sp);
              }
          } else {
              // No process code - search directly by hash
              foundProcess = await EsignDataService.getProcessByHash(generatedHash, sp);
              if (foundProcess) {
                  storedHash = foundProcess.HashHex;
              }
          }

          // 4. Evaluate Result
          if (foundProcess && generatedHash === storedHash) {
              // SUCCESS: Hash matches
              setProcessItem(foundProcess);

              // Fetch Tasks History
              const fetchedTasks = await EsignDataService.getTasksByProcessId(foundProcess.Id, sp);
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const formattedTasks = fetchedTasks.map((t: any) => {
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
              setTasks(formattedTasks);

              setResult({ 
                  type: MessageBarType.success, 
                  message: `SUCCESS: The document integrity is verified. This file is authentic and matches process ${foundProcess.Title}.` 
              });
          } else if (processCode.trim() && storedHash && generatedHash !== storedHash) {
              // MISMATCH: Process code found but hash doesn't match
              setResult({ 
                  type: MessageBarType.severeWarning, 
                  message: 'MISMATCH: The document content does not match the stored record. The file may have been modified.' 
              });
          } else {
              // NOT FOUND: No matching process
              setResult({ 
                  type: MessageBarType.warning, 
                  message: 'NOT FOUND: No matching eSign process found for this document. It may not have been signed through this system.' 
              });
          }

      } catch (err) {
          console.error("Verification Error:", err);
          setResult({ type: MessageBarType.error, message: 'An error occurred during verification. Check console for details.' });
      } finally {
          setLoading(false);
      }
  };

  return (
    <section className={`${styles.docIntegrityChecker} ${hasTeamsContext ? styles.teams : ''}`}>
       <Stack tokens={{ childrenGap: 20 }} style={{ maxWidth: 800, margin: '0 auto', padding: 20 }}>
            <div style={{ textAlign: 'center' }}>
                 <Text variant="xLarge" style={{ fontWeight: 600, color: '#0078d4' }}>Document Integrity Checker</Text>
                 <Text block variant="medium" style={{ color: '#605e5c', marginTop: 5 }}>
                    Verify the authenticity of a signed document by comparing its digital signature.
                 </Text>
            </div>

            <Stack tokens={{ childrenGap: 15 }}>
                {/* File Upload - Primary */}
                <div>
                     <Text block style={{ fontWeight: 600, marginBottom: 5 }}>Upload PDF Document <span style={{ color: '#a80000' }}>*</span></Text>
                     
                     <div 
                        onDragOver={onDragOver}
                        onDragLeave={onDragLeave}
                        onDrop={onDrop}
                        onClick={handleClickUpload}
                        style={{
                            border: `2px dashed ${isDragOver ? '#0078d4' : '#c8c6c4'}`,
                            borderRadius: '4px',
                            backgroundColor: isDragOver ? '#eff6fc' : '#f3f2f1',
                            padding: '30px',
                            textAlign: 'center',
                            cursor: 'pointer',
                            transition: 'all 0.2s ease-in-out'
                        }}
                     >
                        <input 
                            type="file" 
                            accept=".pdf" 
                            ref={fileInputRef}
                            onChange={handleFileChange}
                            style={{ display: 'none' }}
                        />
                        
                        {!file ? (
                            <>
                                <FontIcon iconName="CloudUpload" className={iconClass} />
                                <Text block variant="mediumPlus" style={{ fontWeight: 600, color: '#323130' }}>
                                    Drag & Drop PDF here
                                </Text>
                                <Text block variant="small" style={{ color: '#605e5c', marginTop: 5 }}>
                                    or click to browse
                                </Text>
                            </>
                        ) : (
                            <>
                                <FontIcon iconName="PDF" className={iconClass} style={{ color: '#e81123' }} />
                                <Text block variant="mediumPlus" style={{ fontWeight: 600, color: '#323130' }}>
                                    {file.name}
                                </Text>
                                <Text block variant="small" style={{ color: '#0078d4', marginTop: 5 }}>
                                    Click to change file
                                </Text>
                            </>
                        )}
                     </div>
                </div>

                {/* Process Code - Optional */}
                <TextField 
                    label="Process Code (Optional)" 
                    placeholder="Enter process code (e.g., REQ-ESIGN-XXXX)" 
                    value={processCode}
                    onChange={(e, val) => setProcessCode(val || '')}
                    description="Leave empty to automatically search by document content"
                />

                <PrimaryButton 
                    text="Verify Integrity" 
                    onClick={handleVerify} 
                    disabled={loading || !file}
                    styles={{ root: { marginTop: 10 } }}
                />
            </Stack>

            {loading && <Spinner size={SpinnerSize.large} label="Verifying document..." />}

            {result && (
                <MessageBar messageBarType={result.type} isMultiline={true}>
                    {result.message}
                </MessageBar>
            )}

            {/* Display History Log on Success */}
            {result && result.type === MessageBarType.success && processItem && (
                <div style={{ marginTop: 20 }}>
                    <div style={{ marginBottom: 10 }}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                             <Text variant="large" style={{ fontWeight: 600, display: 'block' }}>{processItem.Title}</Text>
                             {renderStatusBadge(processItem.Status)}
                        </Stack>
                        
                        <Text variant="small" style={{ color: '#605e5c' }}>Created: {formatDate(processItem.Created)}</Text>
                        <div style={{ marginTop: 5 }}>
                            <Text variant="small" style={{ fontWeight: 600 }}>Requestor: </Text>
                            <Text variant="small">{processItem.Requestor ? processItem.Requestor.Title : "Unknown"}</Text>
                        </div>
                    </div>
                    
                    <Text variant="medium" style={{ fontWeight: 600, display: 'block', marginBottom: 10 }}>Process Activity Log</Text>
                    
                    <div style={{ border: '1px solid #edebe9', borderRadius: '4px' }}>
                        <DetailsList
                            items={tasks}
                            columns={columns}
                            selectionMode={SelectionMode.none}
                            layoutMode={0}
                            compact={true}
                            isHeaderVisible={true}
                        />
                    </div>
                    
                    {tasks.length === 0 && (
                        <div style={{ padding: 10, textAlign: 'center', color: '#605e5c' }}>No history records found.</div>
                    )}
                </div>
            )}
       </Stack>
    </section>
  );
};

export default DocIntegrityChecker;
