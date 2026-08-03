import * as React from "react";
import { useState } from "react";
import { Persona, PersonaSize, PersonaPresence } from "@fluentui/react/lib/Persona";
import { Toggle } from "@fluentui/react/lib/Toggle";
import { Stack, IStackTokens, IStackStyles } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { TextField } from "@fluentui/react/lib/TextField";
import { IconButton } from "@fluentui/react/lib/Button";
import { initializeIcons } from "@fluentui/react/lib/Icons";

initializeIcons();

export interface IIDCardProps {
    displayName: string;
    email: string;
    badgeNumber: string;
    phoneNumber?: string;
    waStatus: boolean;
    onWaStatusChange: (checked: boolean) => void;
    onPhoneNumberChange?: (phoneNumber: string) => Promise<void>;
    onBadgeNumberChange?: (badgeNumber: string) => Promise<void>;
    onCheck: () => void;
    loading: boolean;
    checkLoading?: boolean;
    isWhitelisted?: boolean;
}

const stackTokens: IStackTokens = { childrenGap: 10 };
const cardStyles: IStackStyles = {
    root: {
        backgroundColor: "white",
        padding: 20,
        borderRadius: 8,
        boxShadow: "0 4px 8px rgba(0,0,0,0.1)",
        maxWidth: 400,
        minWidth: 300,
        border: "1px solid #e1e1e1",
    },
};

const isValidPhoneNumber = (phone: string | undefined): boolean => {
    return !!phone && phone !== "-" && phone.toLowerCase() !== "not found" && phone.trim() !== "";
};

export const IDCard: React.FC<IIDCardProps> = ({
    displayName,
    email,
    badgeNumber,
    phoneNumber,
    waStatus,
    onWaStatusChange,
    onPhoneNumberChange,
    onBadgeNumberChange,
    onCheck,
    loading,
    checkLoading = false,
    isWhitelisted = false,
}) => {
    const hasPhone = isValidPhoneNumber(phoneNumber);
    const [isEditingPhone, setIsEditingPhone] = useState(false);
    const [editPhoneValue, setEditPhoneValue] = useState(phoneNumber || "+628");
    const [isSavingPhone, setIsSavingPhone] = useState(false);
    const [phoneError, setPhoneError] = useState<string>("");

    const [isEditingBadge, setIsEditingBadge] =useState(false);
    const [editBadgeValue, setEditBadgeValue] = useState(badgeNumber || "");
    const [isSavingBadge, setIsSavingBadge] = useState(false);

    // Validate phone number format: must start with +628 and only contain numbers after +
    const validatePhoneNumber = (phone: string): string => {
        if (!phone.startsWith("+628")) {
            return "Nomor harus diawali dengan +628";
        }
        // Check if all characters after + are digits
        const numberPart = phone.substring(1);
        if (!/^\d+$/.test(numberPart)) {
            return "Hanya boleh berisi angka setelah +";
        }
        if (phone.length < 12) {
            return "Nomor terlalu pendek (min. 12 karakter)";
        }
        if (phone.length > 15) {
            return "Nomor terlalu panjang (max. 15 karakter)";
        }
        return "";
    };

    const handleEditPhone = (): void => {
        setEditPhoneValue(phoneNumber || "+628");
        setPhoneError("");
        setIsEditingPhone(true);
    };

    const handleEditBadge = (): void => {

        setEditBadgeValue(badgeNumber || "");
        setIsEditingBadge(true);
    };

    const handleCancelBadge = (): void => {
        setIsEditingBadge(false);
        setEditBadgeValue(badgeNumber || "");
    };

    const handleSaveBadge = async (): Promise<void> => {
        if (!onBadgeNumberChange) {
            return;
        }
        
        setIsSavingBadge(true);
        
        try {
            await onBadgeNumberChange(editBadgeValue);
            setIsEditingBadge(false);
        
        } catch (err) {
            console.error(
                "Failed to save badge number:",
                err
            );
        
        } finally {
            setIsSavingBadge(false);
        }
    };


    const handleCancelEdit = (): void => {
        setIsEditingPhone(false);
        setEditPhoneValue(phoneNumber || "+628");
        setPhoneError("");
    };

    const handlePhoneInputChange = (value: string): void => {
        // Ensure it always starts with +628
        let newValue = value;
        if (!newValue.startsWith("+")) {
            newValue = "+" + newValue.replace(/^\+/, "");
        }
        // Remove any non-digit characters except the leading +
        newValue = "+" + newValue.substring(1).replace(/\D/g, "");
        
        setEditPhoneValue(newValue);
        setPhoneError(validatePhoneNumber(newValue));
    };

    const handleSavePhone = async (): Promise<void> => {
        if (!onPhoneNumberChange || !editPhoneValue.trim()) return;
        
        const error = validatePhoneNumber(editPhoneValue);
        if (error) {
            setPhoneError(error);
            return;
        }
        
        setIsSavingPhone(true);
        try {
            await onPhoneNumberChange(editPhoneValue.trim());
            setIsEditingPhone(false);
            setPhoneError("");
        } catch (err) {
            console.error("Failed to save phone number:", err);
            alert("Gagal menyimpan nomor telepon. Silakan coba lagi.");
        } finally {
            setIsSavingPhone(false);
        }
    };

    
    return (
        <Stack styles={cardStyles} tokens={stackTokens}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
                <Persona
                    text={displayName}
                    secondaryText={email}
                    size={PersonaSize.size72}
                    presence={waStatus ? PersonaPresence.online : PersonaPresence.dnd}
                    hidePersonaDetails={false}
                />
            </Stack>

            <div style={{ borderTop: "1px solid #f3f2f1", margin: "10px 0" }} />

            <Stack tokens={stackTokens}>
                <Stack
                horizontal
                horizontalAlign="space-between"
                verticalAlign="center">
                    
                    <Text 
                    variant="small"
                    style={{ color: "#666" }}>
                        Badge Number
                    </Text>
                    
                    {isEditingBadge ? (

        <Stack
            horizontal
            verticalAlign="center"
            tokens={{ childrenGap: 4 }}
        >

            <TextField
                value={editBadgeValue}
                onChange={(_, val) =>
                    setEditBadgeValue(val || "")
                }
                styles={{
                    root: {
                        width: 120
                    }
                }}
            />

            <IconButton
                iconProps={{
                    iconName: "Accept"
                }}
                onClick={
                    handleSaveBadge
                }
                disabled={
                    isSavingBadge
                }
            />

            <IconButton
                iconProps={{
                    iconName: "Cancel"
                }}
                onClick={
                    handleCancelBadge
                }
            />

        </Stack>

    ) : (

        <Stack
            horizontal
            verticalAlign="center"
        >

            <Text
                variant="medium"
                style={{
                    fontWeight: 600
                }}
            >
                {badgeNumber || "-"}
            </Text>

            {isWhitelisted && (
            <IconButton
                iconProps={{
                    iconName: "Edit"
                }}
                styles={{
                    root: {
                        width: 24,
                        height: 24
                    },
                    icon: {
                        fontSize: 12,
                        color: "#666"
                    }
            }}
                
                title="Edit badge number"
                onClick={
                    handleEditBadge
                }
            />
            )}

        </Stack>

    )}

</Stack>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Text variant="small" style={{ color: "#666" }}>Phone Number</Text>
                    {isWhitelisted ? (
                        isEditingPhone ? (
                            <Stack tokens={{ childrenGap: 2 }}>
                                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                                    <TextField
                                        value={editPhoneValue}
                                        onChange={(_, val) => handlePhoneInputChange(val || "")}
                                        placeholder="+628xxxxxxxxx"
                                        styles={{ 
                                            root: { width: 130 },
                                            fieldGroup: { 
                                                height: 28,
                                                borderColor: phoneError ? "#f44336" : undefined
                                            }
                                        }}
                                        disabled={isSavingPhone}
                                    />
                                    <IconButton
                                        iconProps={{ iconName: "Accept" }}
                                        title="Save"
                                        onClick={handleSavePhone}
                                        disabled={isSavingPhone || !editPhoneValue.trim() || !!phoneError}
                                        styles={{ 
                                            root: { width: 24, height: 24 },
                                            icon: { fontSize: 12, color: phoneError ? "#ccc" : "#4CAF50" }
                                        }}
                                    />
                                    <IconButton
                                        iconProps={{ iconName: "Cancel" }}
                                        title="Cancel"
                                        onClick={handleCancelEdit}
                                        disabled={isSavingPhone}
                                        styles={{ 
                                            root: { width: 24, height: 24 },
                                            icon: { fontSize: 12, color: "#f44336" }
                                        }}
                                    />
                                </Stack>
                                {phoneError && (
                                    <Text variant="xSmall" style={{ color: "#f44336", fontSize: 10 }}>
                                        {phoneError}
                                    </Text>
                                )}
                            </Stack>
                        ) : (
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                                <Text variant="medium" style={{ fontWeight: 600 }}>
                                    {hasPhone ? phoneNumber : "-"}
                                </Text>
                                {onPhoneNumberChange && (
                                    <IconButton
                                        iconProps={{ iconName: "Edit" }}
                                        title="Edit phone number"
                                        onClick={handleEditPhone}
                                        styles={{ 
                                            root: { width: 24, height: 24 },
                                            icon: { fontSize: 12, color: "#666" }
                                        }}
                                    />
                                )}
                            </Stack>
                        )
                    ) : (
                        <Text
                            variant="medium"
                            style={{ fontWeight: 600 }}
                        >
                            -
                        </Text>
                    )}
                </Stack>
            </Stack>

            {!isWhitelisted && (
                <Stack
                    horizontalAlign="center"
                    style={{ marginTop: 10 }}
                >
                    <button
                        className="btn btn-sm btn-outline-primary"
                        style={{
                            minWidth: "120px"
                        }}
                        onClick={onCheck}
                        disabled={checkLoading}
                    >
                        {checkLoading
                            ? "Loading..."
                            : "Checking"}
                    </button>
                </Stack>
            )}

            <div style={{ borderTop: "1px solid #f3f2f1", margin: "10px 0" }} />

            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Text variant="medium">WhatsApp Notification</Text>
                <Toggle
                    checked={waStatus}
                    onChange={(_, checked) => onWaStatusChange(!!checked)}
                    disabled={loading || !hasPhone}
                    onText="Active"
                    offText="Inactive"
                />
            </Stack>
            {isWhitelisted && !hasPhone && (
                <Text variant="small" style={{ color: "#666", marginTop: "5px", display: "block", fontStyle: "italic" }}>
                    *Klik icon edit untuk update nomor telepon
                </Text>
            )}
        </Stack>
    );
};
