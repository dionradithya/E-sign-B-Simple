import * as React from 'react';
import { Stack, IconButton, DefaultButton, IButtonStyles, TooltipHost, IStyle, Callout, DirectionalHint } from '@fluentui/react';


export type AnnotationMode = 'view' | 'pen' | 'text';

// Available pen colors
const PEN_COLORS = [
    { color: '#ff0000', name: 'Red' },
    { color: '#0000ff', name: 'Blue' },
    { color: '#00aa00', name: 'Green' },
    { color: '#000000', name: 'Black' },
    { color: '#ff6600', name: 'Orange' },
    { color: '#9900ff', name: 'Purple' }
];

// Available pen sizes
const PEN_SIZES = [
    { size: 1, name: 'Fine' },
    { size: 2, name: 'Medium' },
    { size: 4, name: 'Thick' },
    { size: 6, name: 'Bold' }
];

// Available text sizes
const TEXT_SIZES = [
    { size: 8, name: 'Tiny' },
    { size: 10, name: 'Small' },
    { size: 12, name: 'Normal' },
    { size: 14, name: 'Medium' },
    { size: 18, name: 'Large' },
    { size: 24, name: 'X-Large' }
];

interface IAnnotationToolbarProps {
  currentMode: AnnotationMode;
  onModeChange: (mode: AnnotationMode) => void;
  onRotate: () => void;
  onUndo?: () => void;
  canUndo?: boolean;
  textCount?: number;
  signedCount?: number;
  totalPlaceholders?: number;
  onClear: () => void;
  onReject: () => void;
  onApprove: () => void;
  onViewLog: () => void;
  disabled?: boolean;
  // Pen settings
  penColor?: string;
  penSize?: number;
  onPenColorChange?: (color: string) => void;
  onPenSizeChange?: (size: number) => void;
  // Text settings
  textSize?: number;
  onTextSizeChange?: (size: number) => void;
  // Download
  isDownloadable?: boolean;
  onDownload?: () => void;

  isSecretaryOnly?: boolean;
}

const commonButtonRootStyle: IStyle = {
    backgroundColor: '#fff',
    border: '1px solid #c8c6c4',
    borderRadius: 2,
    margin: '0 4px',
    height: 32,
    width: 32,
    selectors: {
        '@media (max-width: 640px)': {
            width: '100%',
            margin: '4px 0',
            height: 40 // Taller for touch
        }
    }
};

const commonButtonStyles: IButtonStyles = {
  root: commonButtonRootStyle,
  rootHovered: {
    backgroundColor: '#f3f2f1',
    selectors: {
        '@media (max-width: 640px)': {
             width: '100%' 
        }
    }
  },
  rootChecked: {
    backgroundColor: '#c7e0f4',
    borderColor: '#0078d4',
    selectors: {
        '@media (max-width: 640px)': {
             width: '100%' 
        }
    }
  },
  rootDisabled: {
      backgroundColor: '#f3f2f1',
      borderColor: '#e1dfdd',
      pointerEvents: 'none',
      opacity: 0.6
  }
};

export const AnnotationToolbar: React.FunctionComponent<IAnnotationToolbarProps> = (props) => {
  const [isColorCalloutOpen, setIsColorCalloutOpen] = React.useState(false);
  const [isSizeCalloutOpen, setIsSizeCalloutOpen] = React.useState(false);
  const [isTextSizeCalloutOpen, setIsTextSizeCalloutOpen] = React.useState(false);
  const colorButtonRef = React.useRef<HTMLDivElement>(null);
  const sizeButtonRef = React.useRef<HTMLDivElement>(null);
  const textSizeButtonRef = React.useRef<HTMLDivElement>(null);

  const currentColor = props.penColor || '#ff0000';
  const currentSize = props.penSize || 2;
  const currentTextSize = props.textSize || 14;

  return (
    <Stack 
        horizontal 
        wrap 
        verticalAlign="center" 
        tokens={{ childrenGap: 8 }}
        styles={{ 
        root: { 
            padding: 10, 
            background: '#f3f2f1', 
            borderBottom: '1px solid #e1dfdd',
            selectors: {
                '@media (max-width: 640px)': {
                    flexDirection: 'column',
                    alignItems: 'stretch'
                }
            }
        } 
    }}>
      
      {/* Draw (Pen) Button */}
      <TooltipHost content="Draw (Pen)">
        <IconButton 
            iconProps={{ iconName: 'Edit' }} 
            styles={props.currentMode === 'pen' ? 
                { ...commonButtonStyles, root: { ...commonButtonRootStyle, backgroundColor: '#c7e0f4', borderColor: '#0078d4' } } 
                : commonButtonStyles}
            onClick={() => props.onModeChange(props.currentMode === 'pen' ? 'view' : 'pen')}
            disabled={props.disabled}
        />
      </TooltipHost>

      {/* Pen Color Picker - Show when pen mode is active */}
      {props.currentMode === 'pen' && (
        <>
          <div ref={colorButtonRef}>
            <TooltipHost content="Pen Color">
              <div 
                onClick={() => setIsColorCalloutOpen(!isColorCalloutOpen)}
                style={{
                  width: 28,
                  height: 28,
                  backgroundColor: currentColor,
                  border: '2px solid #c8c6c4',
                  borderRadius: 4,
                  cursor: 'pointer',
                  boxShadow: 'inset 0 0 0 2px white'
                }}
              />
            </TooltipHost>
          </div>
          {isColorCalloutOpen && (
            <Callout
              target={colorButtonRef.current}
              onDismiss={() => setIsColorCalloutOpen(false)}
              directionalHint={DirectionalHint.bottomCenter}
              isBeakVisible={true}
            >
              <div style={{ padding: 12 }}>
                <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 8 }}>Pen Color</div>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', maxWidth: 140 }}>
                  {PEN_COLORS.map(c => (
                    <div
                      key={c.color}
                      onClick={() => {
                        props.onPenColorChange?.(c.color);
                        setIsColorCalloutOpen(false);
                      }}
                      style={{
                        width: 28,
                        height: 28,
                        backgroundColor: c.color,
                        border: currentColor === c.color ? '3px solid #0078d4' : '2px solid #c8c6c4',
                        borderRadius: 4,
                        cursor: 'pointer',
                        boxShadow: currentColor === c.color ? '0 0 4px rgba(0,120,212,0.5)' : 'none'
                      }}
                      title={c.name}
                    />
                  ))}
                </div>
              </div>
            </Callout>
          )}

          {/* Pen Size Picker */}
          <div ref={sizeButtonRef}>
            <TooltipHost content="Pen Size">
              <div 
                onClick={() => setIsSizeCalloutOpen(!isSizeCalloutOpen)}
                style={{
                  width: 28,
                  height: 28,
                  backgroundColor: '#fff',
                  border: '2px solid #c8c6c4',
                  borderRadius: 4,
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center'
                }}
              >
                <div style={{
                  width: currentSize * 3,
                  height: currentSize * 3,
                  maxWidth: 20,
                  maxHeight: 20,
                  backgroundColor: currentColor,
                  borderRadius: '50%'
                }} />
              </div>
            </TooltipHost>
          </div>
          {isSizeCalloutOpen && (
            <Callout
              target={sizeButtonRef.current}
              onDismiss={() => setIsSizeCalloutOpen(false)}
              directionalHint={DirectionalHint.bottomCenter}
              isBeakVisible={true}
            >
              <div style={{ padding: 12 }}>
                <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 8 }}>Pen Size</div>
                <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  {PEN_SIZES.map(s => (
                    <div
                      key={s.size}
                      onClick={() => {
                        props.onPenSizeChange?.(s.size);
                        setIsSizeCalloutOpen(false);
                      }}
                      style={{
                        width: 32,
                        height: 32,
                        backgroundColor: currentSize === s.size ? '#c7e0f4' : '#fff',
                        border: currentSize === s.size ? '2px solid #0078d4' : '2px solid #c8c6c4',
                        borderRadius: 4,
                        cursor: 'pointer',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center'
                      }}
                      title={s.name}
                    >
                      <div style={{
                        width: s.size * 3,
                        height: s.size * 3,
                        backgroundColor: currentColor,
                        borderRadius: '50%'
                      }} />
                    </div>
                  ))}
                </div>
              </div>
            </Callout>
          )}
        </>
      )}

      <TooltipHost content="Add Text">
         <IconButton 
            iconProps={{ iconName: 'TextField' }} 
            styles={props.currentMode === 'text' ? 
                { ...commonButtonStyles, root: { ...commonButtonRootStyle, backgroundColor: '#c7e0f4', borderColor: '#0078d4' } } 
                : commonButtonStyles}
            onClick={() => props.onModeChange(props.currentMode === 'text' ? 'view' : 'text')}
            disabled={props.disabled}
        />
      </TooltipHost>

      {/* Text Size Picker - Show when text mode is active */}
      {props.currentMode === 'text' && (
        <>
          <div ref={textSizeButtonRef}>
            <TooltipHost content="Text Size">
              <div 
                onClick={() => setIsTextSizeCalloutOpen(!isTextSizeCalloutOpen)}
                style={{
                  minWidth: 40,
                  height: 28,
                  backgroundColor: '#fff',
                  border: '2px solid #c8c6c4',
                  borderRadius: 4,
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  padding: '0 8px',
                  fontSize: 12,
                  fontWeight: 600,
                  color: '#323130'
                }}
              >
                {currentTextSize}px
              </div>
            </TooltipHost>
          </div>
          {isTextSizeCalloutOpen && (
            <Callout
              target={textSizeButtonRef.current}
              onDismiss={() => setIsTextSizeCalloutOpen(false)}
              directionalHint={DirectionalHint.bottomCenter}
              isBeakVisible={true}
            >
              <div style={{ padding: 12 }}>
                <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 8 }}>Text Size</div>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                  {TEXT_SIZES.map(s => (
                    <div
                      key={s.size}
                      onClick={() => {
                        props.onTextSizeChange?.(s.size);
                        setIsTextSizeCalloutOpen(false);
                      }}
                      style={{
                        minWidth: 50,
                        padding: '6px 10px',
                        backgroundColor: currentTextSize === s.size ? '#c7e0f4' : '#fff',
                        border: currentTextSize === s.size ? '2px solid #0078d4' : '2px solid #c8c6c4',
                        borderRadius: 4,
                        cursor: 'pointer',
                        textAlign: 'center',
                        fontSize: 12,
                        fontWeight: currentTextSize === s.size ? 600 : 400
                      }}
                      title={s.name}
                    >
                      {s.size}px
                    </div>
                  ))}
                </div>
              </div>
            </Callout>
          )}
        </>
      )}

      <div className="hidden-sm" style={{ width: 1, height: 24, backgroundColor: '#c8c6c4', margin: '0 8px' }} />
      
       <style>{`
          @media (max-width: 640px) {
             .toolbar-separator { display: none !important; }
          }
       `}</style>
      <div className="toolbar-separator" style={{ width: 1, height: 24, backgroundColor: '#c8c6c4', margin: '0 8px' }} />


      <TooltipHost content="Undo Last Stroke">
          <IconButton
              iconProps={{ iconName: 'Undo' }}
              styles={commonButtonStyles}
              onClick={props.onUndo}
              disabled={!props.canUndo || props.disabled}
          />
      </TooltipHost>

      <TooltipHost content="Clear All Annotations">
        <IconButton 
            iconProps={{ iconName: 'Delete' }} 
            styles={commonButtonStyles}
            onClick={props.onClear}
            disabled={props.disabled}
        />
      </TooltipHost>

      {props.isDownloadable !== false && (
        <TooltipHost content="Download PDF">
          <IconButton 
              iconProps={{ iconName: 'Download' }} 
              styles={{
                root: { ...commonButtonRootStyle, borderColor: '#0078d4', backgroundColor: '#eff6fc' },
                rootHovered: { backgroundColor: '#deecf9', borderColor: '#0078d4' },
                rootDisabled: { backgroundColor: '#f3f2f1', borderColor: '#e1dfdd', pointerEvents: 'none', opacity: 0.6 },
                icon: { color: '#0078d4' }
              }}
              onClick={props.onDownload}
              disabled={props.disabled}
          />
        </TooltipHost>
      )}
      
      <div style={{ padding: '0 8px', display: 'flex', alignItems: 'center', fontSize: 12, fontWeight: 600, color: '#0078d4', justifyContent: 'center' }}>
          Comments: {props.textCount || 0}
      </div>

      <div style={{ padding: '0 8px', display: 'flex', alignItems: 'center', fontSize: 12, fontWeight: 600, color: '#107c10', justifyContent: 'center' }}>
          Signed: {props.signedCount || 0}/{props.totalPlaceholders || 0}
      </div>

      <div style={{ flexGrow: 1 }} /> 
      
      <DefaultButton 
          text="Approval Log"
          disabled={props.disabled}
          styles={{
              root: { ...commonButtonRootStyle, borderColor: '#0078d4', backgroundColor: '#eff6fc', minWidth: 100, width: 'auto', borderWidth: 2, 
                  selectors: { 
                      '@media (max-width: 640px)': { width: '100%', margin: '4px 0' } 
                  } 
              },
              rootHovered: { backgroundColor: '#deecf9',  borderColor: '#0078d4',
                  selectors: { 
                      '@media (max-width: 640px)': { width: '100%' } 
                  } 
              },
              rootDisabled: {
                  backgroundColor: '#f3f2f1',
                  borderColor: '#e1dfdd',
              },
              label: { color: props.disabled ? '#a19f9d' : '#0078d4', fontWeight: 'bold' }
          }}
          onClick={props.onViewLog}
      /> 

      {!props.isSecretaryOnly && (
        <DefaultButton 
            text="APPROVE"
            disabled={props.disabled}
            styles={{
                root: { ...commonButtonRootStyle, borderColor: 'green', backgroundColor: '#e6ffec', minWidth: 80, width: 'auto', borderWidth: 2, 
                    selectors: { 
                        '@media (max-width: 640px)': { width: '100%', margin: '4px 0' } 
                    } 
                },
                rootHovered: { backgroundColor: '#dff6dd',  borderColor: 'green',
                    selectors: { 
                        '@media (max-width: 640px)': { width: '100%' } 
                    } 
                },
                rootDisabled: {
                    backgroundColor: '#f3f2f1',
                    borderColor: '#e1dfdd',
                },
                label: { color: props.disabled ? '#a19f9d' : 'green', fontWeight: 'bold' }
            }}
            onClick={props.onApprove}
        />
        )}

      {!props.isSecretaryOnly && (
        <DefaultButton 
            text="REJECT"
            disabled={props.disabled}
            styles={{
                root: { ...commonButtonRootStyle, borderColor: 'red', backgroundColor: '#fde7e9', minWidth: 80, width: 'auto', borderWidth: 2, 
                    selectors: { 
                        '@media (max-width: 640px)': { width: '100%', margin: '4px 0' } 
                    } 
                },
                rootHovered: { backgroundColor: '#fdd3d6',  borderColor: 'red',
                    selectors: { 
                        '@media (max-width: 640px)': { width: '100%' } 
                    } 
                },
                rootDisabled: {
                    backgroundColor: '#f3f2f1',
                    borderColor: '#e1dfdd',
                },
                label: { color: props.disabled ? '#a19f9d' : 'red', fontWeight: 'bold' }
            }}
            onClick={props.onReject}
        />
        )}

    </Stack>
  );
};
