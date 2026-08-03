import * as React from 'react';
import styles from './DocumentTrackerEsign.module.scss';
import { IDocumentTrackerEsignProps } from './IDocumentTrackerEsignProps';
import { 
  Stack, 
  TextField, 
  PrimaryButton,
  DefaultButton,
  MessageBar, 
  MessageBarType,
  Spinner,
  SpinnerSize,
  Text,
  Pivot,
  PivotItem,
  Icon
} from '@fluentui/react';
import { EsignDataService } from '../../../common/services/EsignDataService';
import { ProcessHistoryItem } from './ProcessHistoryItem';
import jsQR from 'jsqr';

export interface IDocumentTrackerEsignState {
  // Code tab
  searchQuery: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  processes: any[];
  loading: boolean;
  searched: boolean;
  error: string;

  // QR tab
  activeTab: string;
  qrCode: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  qrProcesses: any[];
  qrLoading: boolean;
  qrSearched: boolean;
  qrError: string;
  qrFileName: string;
  codeFromUrl: boolean;
}

export default class DocumentTrackerEsign extends React.Component<IDocumentTrackerEsignProps, IDocumentTrackerEsignState> {
  
  private _fileInputRef: React.RefObject<HTMLInputElement>;

  constructor(props: IDocumentTrackerEsignProps) {
    super(props);

    // Check for ?code= in URL
    const urlParams = new URLSearchParams(window.location.search);
    const codeFromUrl = urlParams.get('code') || '';

    this._fileInputRef = React.createRef<HTMLInputElement>();

    this.state = {
      // Code tab
      searchQuery: "",
      processes: [],
      loading: false,
      searched: false,
      error: "",

      // QR tab
      activeTab: codeFromUrl ? 'qr' : 'code',
      qrCode: codeFromUrl,
      qrProcesses: [],
      qrLoading: false,
      qrSearched: false,
      qrError: "",
      qrFileName: "",
      codeFromUrl: !!codeFromUrl
    };
  }

  public componentDidMount(): void {
    // Auto-search if code is provided in URL
    if (this.state.qrCode) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this._handleQrSearch(this.state.qrCode);
    }
  }

  // ============================
  // Code Tab Handlers
  // ============================
  private _onSearchChange = (_ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ searchQuery: newValue || "" });
  }

  private _handleSearch = async (): Promise<void> => {
    const { searchQuery } = this.state;
    if (!searchQuery.trim()) {
      this.setState({ error: "Please enter a tracking code.", processes: [], searched: true });
      return;
    }

    this.setState({ loading: true, error: "", processes: [], searched: true });

    try {
      const foundProcess = await EsignDataService.getProcessByUniqueCode(searchQuery.trim(), this.props.sp);

      if (!foundProcess) {
        this.setState({ loading: false, processes: [], error: "No process found with this ID." });
        return;
      }

      if (!foundProcess.FileRef0) {
        this.setState({ loading: false, processes: [foundProcess] });
        return;
      }

      const allProcesses = await EsignDataService.getRelatedProcesses(
          foundProcess.FileRef0, 
          foundProcess.HashHex, 
          this.props.sp
      );
      
      this.setState({ 
        processes: allProcesses, 
        loading: false 
      });

    } catch (err) {
      console.error("Error searching document history:", err);
      this.setState({ loading: false, error: "An error occurred while searching. Please try again." });
    }
  }

  private _handleKeyDown = (event: React.KeyboardEvent<HTMLInputElement>): void => {
    if (event.key === 'Enter') {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this._handleSearch();
    }
  }

  // ============================
  // QR Tab Handlers
  // ============================
  private _handleQrSearch = async (code: string): Promise<void> => {
    if (!code.trim()) {
      this.setState({ qrError: "No code found in QR image.", qrProcesses: [], qrSearched: true, qrLoading: false });
      return;
    }

    this.setState({ qrLoading: true, qrError: "", qrProcesses: [], qrSearched: true });

    try {
      const foundProcess = await EsignDataService.getProcessByUniqueCode(code.trim(), this.props.sp);

      if (!foundProcess) {
        this.setState({ qrLoading: false, qrProcesses: [], qrError: `No process found for code: ${code}` });
        return;
      }

      if (!foundProcess.FileRef0) {
        this.setState({ qrLoading: false, qrProcesses: [foundProcess] });
        return;
      }

      const allProcesses = await EsignDataService.getRelatedProcesses(
          foundProcess.FileRef0, 
          foundProcess.HashHex, 
          this.props.sp
      );
      
      this.setState({ 
        qrProcesses: allProcesses, 
        qrLoading: false 
      });

    } catch (err) {
      console.error("Error searching by QR code:", err);
      this.setState({ qrLoading: false, qrError: "An error occurred while searching. Please try again." });
    }
  }

  private _handleUploadClick = (): void => {
    if (this._fileInputRef.current) {
      this._fileInputRef.current.click();
    }
  }

  private _handleFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const file = event.target.files?.[0];
    if (!file) return;

    this.setState({ qrFileName: file.name, qrError: "", qrProcesses: [], qrSearched: false });

    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        // Draw image to canvas and extract ImageData
        const canvas = document.createElement('canvas');
        canvas.width = img.width;
        canvas.height = img.height;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          this.setState({ qrError: "Failed to process image." });
          return;
        }
        ctx.drawImage(img, 0, 0);
        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        
        // Decode QR code
        const qrResult = jsQR(imageData.data, imageData.width, imageData.height);
        
        if (qrResult && qrResult.data) {
          // Try to extract code from URL or use as-is
          const extractedCode = this._extractCodeFromQrData(qrResult.data);
          this.setState({ qrCode: extractedCode });
          // eslint-disable-next-line @typescript-eslint/no-floating-promises
          this._handleQrSearch(extractedCode);
        } else {
          this.setState({ qrError: "Could not read QR code from the uploaded image. Please try a clearer image.", qrSearched: true });
        }
      };
      img.onerror = () => {
        this.setState({ qrError: "Failed to load image file." });
      };
      img.src = e.target?.result as string;
    };
    reader.readAsDataURL(file);

    // Reset input so the same file can be re-selected
    event.target.value = '';
  }

  /**
   * Extracts the code from QR data.
   * QR might contain a full URL like https://domain/page?code=REQ-ESIGN-XXX
   * or just the code itself.
   */
  private _extractCodeFromQrData(data: string): string {
    try {
      const url = new URL(data);
      const code = url.searchParams.get('code');
      if (code) return code;
    } catch {
      // Not a URL, use as-is
    }
    return data;
  }

  private _onTabChange = (item?: PivotItem): void => {
    if (item) {
      this.setState({ activeTab: item.props.itemKey || 'code' });
    }
  }

  // ============================
  // Render
  // ============================
  public render(): React.ReactElement<IDocumentTrackerEsignProps> {
    const { processes, loading, error, searched, searchQuery, activeTab } = this.state;
    const { qrProcesses, qrLoading, qrError, qrSearched, qrCode, qrFileName } = this.state;

    return (
      <section className={styles.documentTrackerEsign}>
         <Stack tokens={{ childrenGap: 20 }} style={{ maxWidth: 900, margin: '0 auto', padding: 20 }}>
            <div>
              <Text variant="xLarge" style={{ fontWeight: 600, color: '#0078d4' }}>Document Tracker</Text>
              <Text block variant="medium" style={{ color: '#605e5c', marginTop: 5 }}>
                Track your document approval status by code or QR scan.
              </Text>
            </div>

            <Pivot
              selectedKey={activeTab}
              onLinkClick={this._onTabChange}
              styles={{ root: { marginBottom: 10 } }}
            >
              {/* ===== CODE TAB ===== */}
              <PivotItem headerText="Code" itemKey="code" itemIcon="Search">
                <Stack tokens={{ childrenGap: 15 }} style={{ paddingTop: 15 }}>
                  <div className={styles.searchContainer}>
                     <TextField 
                        placeholder="Enter Process ID (e.g., REQ-ESIGN-XXXXX)" 
                        value={searchQuery}
                        onChange={this._onSearchChange}
                        onKeyDown={this._handleKeyDown}
                        styles={{ root: { flexGrow: 1 } }}
                     />
                     <PrimaryButton 
                        text="Search" 
                        iconProps={{ iconName: 'Search' }} 
                        onClick={this._handleSearch} 
                        disabled={loading}
                     />
                  </div>

                  {loading && (
                    <Spinner size={SpinnerSize.large} label="Searching records..." />
                  )}

                  {!loading && error && (
                    <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
                  )}

                  {!loading && !error && searched && processes.length === 0 && (
                     <MessageBar messageBarType={MessageBarType.info}>No records found.</MessageBar>
                  )}

                  {!loading && !error && processes.length > 0 && (
                     <Stack tokens={{ childrenGap: 15 }}>
                        <Text variant="large" style={{ fontWeight: 600, marginBottom: 10 }}>
                           Found {processes.length} related process{processes.length > 1 ? 'es' : ''}
                        </Text>
                        
                        {processes.map((process) => (
                          <ProcessHistoryItem 
                            key={process.Id} 
                            process={process} 
                            sp={this.props.sp}
                            context={this.props.context}
                          />
                        ))}
                     </Stack>
                  )}
                </Stack>
              </PivotItem>

              {/* ===== QR CHECK TAB ===== */}
              <PivotItem headerText="QR Check" itemKey="qr" itemIcon="QRCode">
                <Stack tokens={{ childrenGap: 15 }} style={{ paddingTop: 15 }}>
                  
                  {/* Hidden file input (always present for upload functionality) */}
                  <input 
                    type="file" 
                    accept="image/*" 
                    ref={this._fileInputRef}
                    onChange={this._handleFileChange}
                    style={{ display: 'none' }}
                  />

                  {/* When code comes from URL: show document info card */}
                  {this.state.codeFromUrl && qrCode && !qrLoading && qrProcesses.length > 0 && (
                    <div style={{
                      border: '1px solid #edebe9',
                      borderRadius: 8,
                      padding: '20px',
                      backgroundColor: '#faf9f8'
                    }}>
                      <Stack tokens={{ childrenGap: 12 }}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                          <Icon iconName="QRCode" style={{ fontSize: 28, color: '#0078d4' }} />
                          <Text variant="large" style={{ fontWeight: 600, color: '#0078d4' }}>QR Code Verified</Text>
                        </Stack>
                        <div style={{ display: 'grid', gridTemplateColumns: '140px 1fr', gap: '8px 12px', marginTop: 8 }}>
                          <Text style={{ fontWeight: 600, color: '#605e5c', fontSize: 13 }}>Tracking Code</Text>
                          <Text style={{ fontSize: 13 }}>{qrCode}</Text>
                          
                          <Text style={{ fontWeight: 600, color: '#605e5c', fontSize: 13 }}>File Name</Text>
                          <Text style={{ fontSize: 13 }}>{qrProcesses[0].FileRef0 ? qrProcesses[0].FileRef0.split('/').pop() : 'Unknown'}</Text>
                          
                          <Text style={{ fontWeight: 600, color: '#605e5c', fontSize: 13 }}>Requestor</Text>
                          <Text style={{ fontSize: 13 }}>{qrProcesses[0].Requestor ? qrProcesses[0].Requestor.Title : 'Unknown'}</Text>
                          
                          <Text style={{ fontWeight: 600, color: '#605e5c', fontSize: 13 }}>Status</Text>
                          <Text style={{ fontSize: 13 }}>{qrProcesses[0].Status}</Text>
                        </div>
                      </Stack>
                    </div>
                  )}

                  {/* Show upload area: when no URL code, OR when URL code search failed */}
                  {(!this.state.codeFromUrl || (this.state.codeFromUrl && !qrLoading && (qrError || (qrSearched && qrProcesses.length === 0)))) && (
                    <div className={styles.qrUploadArea} onClick={this._handleUploadClick}>
                      <Icon iconName="QRCode" style={{ fontSize: 48, color: '#0078d4', marginBottom: 10 }} />
                      <Text variant="mediumPlus" style={{ fontWeight: 600, color: '#323130' }}>
                        Upload QR Code Image
                      </Text>
                      <Text variant="small" style={{ color: '#605e5c', marginTop: 4 }}>
                        Click to select an image file containing a QR code
                      </Text>
                      {qrFileName && (
                        <Text variant="small" style={{ color: '#0078d4', marginTop: 8, fontWeight: 600 }}>
                          <Icon iconName="Attach" style={{ marginRight: 4 }} />
                          {qrFileName}
                        </Text>
                      )}
                    </div>
                  )}

                  {/* Upload another button (only for upload mode, not URL mode) */}
                  {!this.state.codeFromUrl && !qrLoading && qrSearched && qrProcesses.length > 0 && (
                    <DefaultButton 
                      text="Upload Another QR Code" 
                      iconProps={{ iconName: 'Upload' }}
                      onClick={this._handleUploadClick}
                    />
                  )}

                  {qrLoading && (
                    <Spinner size={SpinnerSize.large} label="Searching records..." />
                  )}

                  {!qrLoading && qrError && (
                    <MessageBar messageBarType={MessageBarType.error}>{qrError}</MessageBar>
                  )}

                  {!qrLoading && !qrError && qrSearched && qrProcesses.length === 0 && (
                     <MessageBar messageBarType={MessageBarType.info}>No records found.</MessageBar>
                  )}

                  {!qrLoading && !qrError && qrProcesses.length > 0 && (
                     <Stack tokens={{ childrenGap: 15 }}>
                        <Text variant="large" style={{ fontWeight: 600, marginBottom: 10 }}>
                           Found {qrProcesses.length} related process{qrProcesses.length > 1 ? 'es' : ''}
                        </Text>
                        
                        {qrProcesses.map((process) => (
                          <ProcessHistoryItem 
                            key={process.Id} 
                            process={process} 
                            sp={this.props.sp}
                            context={this.props.context}
                          />
                        ))}
                     </Stack>
                  )}
                </Stack>
              </PivotItem>
            </Pivot>
         </Stack>
      </section>
    );
  }
}
