import * as React from "react";
import { useRef, useState, useEffect } from "react";
import { EraserRegular } from "@fluentui/react-icons";
import SignaturePad from "signature_pad";

type SignatureComponentProps = {
  onBack: (val: boolean) => void;
  onSave: (val: string) => void;
  loading: boolean;
};

const SignatureComponent: React.FC<SignatureComponentProps> = ({
  onBack,
  onSave,
  loading,
}) => {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const signaturePadRef = useRef<SignaturePad | null>(null);
  const [isEmpty, setIsEmpty] = useState(true);

  useEffect(() => {
    if (!canvasRef.current) return;

    const pad = new SignaturePad(canvasRef.current, {
      penColor: "blue",
      minWidth: 1,
      maxWidth: 2.5,
    });

    // Update state when user finishes a stroke
    pad.addEventListener("endStroke", () => {
      setIsEmpty(pad.isEmpty());
    });

    signaturePadRef.current = pad;
  }, [signaturePadRef, canvasRef]);

  const clear = (): void => {
    signaturePadRef.current?.clear();
    setIsEmpty(true);
  };

  const getDataURL = (): string => {
    return signaturePadRef.current?.toDataURL("image/png") ?? "";
  };

  const handleBack: () => void = () => onBack(false);
  const handleSave: () => void = () => {
    if (signaturePadRef.current?.isEmpty()) {
      alert("❌ Lengkapi data sebelum menyimpan!");
      return;
    }
    onSave(getDataURL());
  };

  return (
    <div className="d-flex flex-column align-items-center w-100">
      {/* Canvas Area */}
      <div
        className="position-relative mb-2 mb-md-3 w-100"
        style={{
          maxWidth: "240px",
          minWidth: "160px",
          aspectRatio: "3/2",
          backgroundColor: "#fafafa",
          borderRadius: "12px",
          padding: "8px",
          border: isEmpty ? "2px dashed #e0e0e0" : "2px solid #4CAF50",
          transition: "border-color 0.3s ease",
          boxSizing: "border-box",
          display: "flex",
          justifyContent: "center",
          alignItems: "center"
        }}
      >
        <canvas
          ref={canvasRef}
          width={220}
          height={140}
          style={{
            display: "block",
            borderRadius: "8px",
            backgroundColor: "white",
            cursor: "crosshair",
            width: "100%",
            height: "100%",
            maxWidth: "220px",
            maxHeight: "140px"
          }}
        />
        <button
          onClick={clear}
          className="position-absolute btn btn-light btn-sm"
          style={{
            top: "8px",
            right: "8px",
            borderRadius: "50%",
            width: "28px",
            height: "28px",
            padding: 0,
            boxShadow: "0 2px 4px rgba(0,0,0,0.1)"
          }}
          title="Clear"
        >
          <EraserRegular fontSize={14} color="#666" />
        </button>
        {isEmpty && (
          <div
            className="position-absolute text-muted small"
            style={{
              bottom: "16px",
              left: "50%",
              transform: "translateX(-50%)",
              pointerEvents: "none",
              opacity: 0.5,
              fontSize: "clamp(0.65rem, 1.5vw, 0.875rem)"
            }}
          >
            Draw here
          </div>
        )}
      </div>

      {/* Action Buttons */}
      <div className="d-flex w-100" style={{ maxWidth: "240px", gap: "8px" }}>
        <button
          className="btn flex-fill"
          onClick={handleBack}
          style={{
            backgroundColor: "#f5f5f5",
            color: "#666",
            border: "none",
            borderRadius: "8px",
            padding: "8px 12px",
            fontWeight: 500,
            fontSize: "clamp(0.75rem, 2vw, 0.875rem)"
          }}
        >
          Cancel
        </button>
        <button
          className="btn flex-fill d-flex justify-content-center align-items-center"
          onClick={handleSave}
          disabled={loading || isEmpty}
          style={{
            backgroundColor: loading || isEmpty ? "#ccc" : "#2196F3",
            color: "white",
            border: "none",
            borderRadius: "8px",
            padding: "8px 12px",
            fontWeight: 500,
            gap: "6px",
            fontSize: "clamp(0.75rem, 2vw, 0.875rem)"
          }}
        >
          {loading && (
            <div className="spinner-border spinner-border-sm" role="status">
              <span className="visually-hidden">Loading...</span>
            </div>
          )}
          Save
        </button>
      </div>
    </div>
  );
};

export default SignatureComponent;
