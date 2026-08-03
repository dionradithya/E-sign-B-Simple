import { FormDisplayMode } from "@microsoft/sp-core-library";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";

/**
 * Annotation mode type - re-exported for convenience
 */
export type AnnotationMode = 'view' | 'pen' | 'text';

/**
 * Point coordinates for annotations
 */
export interface IPoint {
    x: number;
    y: number;
}

/**
 * Drawing helper for managing drawing state
 */
export interface IDrawingHelper {
    isDrawing: boolean;
    currentPath: { x: number; y: number }[];
    x: number;
    y: number;
}

/**
 * Path annotation data structure
 */
export interface IPathAnnotation {
    page: number;
    points: IPoint[];
    color: string;
    thickness: number;
}

/**
 * Text annotation data structure
 */
export interface ITextAnnotation {
    id: string;
    page: number;
    x: number;
    y: number;
    text: string;
    color: string;
    size: number;
}

/**
 * Signature placeholder configuration
 */
export interface ISignaturePlaceholder {
    page: number;
    x: number;
    y: number;
    width: number;
    height: number;
    type: 'initial' | 'signature';
    approverName: string;
    signatureData?: string;
    includeName?: boolean;
    includeDate?: boolean;
    includeBadge?: boolean;
    badgeNumber?: string;
}

/**
 * Props for InitiateEsignForm component
 */
export interface IInitiateEsignFormProps {
    context: FormCustomizerContext;
    displayMode: FormDisplayMode;
    onSave: () => void;
    onClose: () => void;
}

/**
 * State for InitiateEsignForm component
 */
export interface IInitiateEsignFormState {
    fileUrl: string | undefined;
    numPages: number | undefined;
    currentWidth: number;
    loading: boolean;
    error: string | undefined;
    placeholders: ISignaturePlaceholder[];
    
    // Annotation State
    annotationMode: AnnotationMode;
    rotation: number;
    paths: IPathAnnotation[];
    texts: ITextAnnotation[];
    penColor: string;
    penSize: number;
    textSize: number;

    // Reject Dialog State
    isRejectDialogOpen?: boolean;
    rejectComment?: string;
    rejectCommentError?: string;
    isSuccessDialogOpen?: boolean;

    // Blocking Dialog State
    isBlockingDialogOpen?: boolean;
    blockingDialogTitle?: string;
    blockingDialogMessage?: string;

    isSaving: boolean;
    
    // Signature Dialog State
    isSignatureDialogOpen?: boolean;
    activePlaceholderIndex?: number;
    userSignatureData?: string | undefined;
    isSignatureLoading?: boolean;
    signatureError?: string;
    
    // New Signature Pad State
    signatureDialogTab?: string; // 'saved' | 'draw'
    newSignatureData?: string | undefined; // Data URL of drawn signature
    saveAsDefault?: boolean;
    signAll?: boolean;

    // Text Interaction State
    editingTextId: string | undefined;
    draggingTextId: string | undefined;

    serverRelativeUrl?: string;

    processId?: number;
    processTitle?: string;

    isValidationErrorOpen?: boolean;
    successMessage?: string;
    currentUserId?: string;
    signingUserId?: number;
    signingUserDisplayName?: string;

    isApproveDialogOpen?: boolean;
    approveComment?: string;
    isLogDialogOpen?: boolean;

    // File Not Found Dialog State
    isFileNotFoundDialogOpen?: boolean;
    fileNotFoundProcessId?: number;

    // Download
    isDownloadable?: boolean;

    // Secretary 
    isSecretaryOnly?: boolean;
}
