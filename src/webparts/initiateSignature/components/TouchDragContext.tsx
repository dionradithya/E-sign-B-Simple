import * as React from 'react';
import { IApprover } from './Sidebar';

// Interface for touch drag data
export interface ITouchDragData {
    type: 'initial' | 'signature';
    approverId: number;
    checklistName: boolean;
    checklistDate: boolean;
    checklistBadge: boolean;
}

interface ITouchDragContextType {
    dragData: ITouchDragData | undefined;
    setDragData: (data: ITouchDragData | undefined) => void;
    isDragging: boolean;
    setIsDragging: (value: boolean) => void;
}

const TouchDragContext = React.createContext<ITouchDragContextType | undefined>(undefined);

export const TouchDragProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    const [dragData, setDragData] = React.useState<ITouchDragData | undefined>(undefined);
    const [isDragging, setIsDragging] = React.useState<boolean>(false);

    return (
        <TouchDragContext.Provider value={{ dragData, setDragData, isDragging, setIsDragging }}>
            {children}
        </TouchDragContext.Provider>
    );
};

export const useTouchDrag = (): ITouchDragContextType => {
    const context = React.useContext(TouchDragContext);
    if (!context) {
        throw new Error('useTouchDrag must be used within a TouchDragProvider');
    }
    return context;
};

// Helper to start touch drag from a button
export const startTouchDrag = (
    setDragData: (data: ITouchDragData | undefined) => void,
    setIsDragging: (value: boolean) => void,
    approver: IApprover,
    type: 'initial' | 'signature'
): void => {
    setDragData({
        type,
        approverId: approver.id,
        checklistName: approver.includeName,
        checklistDate: approver.includeDate,
        checklistBadge: approver.includeBadge
    });
    setIsDragging(true);
};

export default TouchDragContext;
