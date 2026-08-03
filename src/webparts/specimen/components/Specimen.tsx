import * as React from "react";
import { useState, useEffect } from "react";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import { GraphFI } from "@pnp/graph";
import type { ISpecimenProps } from "./ISpecimenProps";
import { useSpecimens } from "../hooks/useSpecimens";
import SignatureComponent from "./SignatureComponent";
import { dataURLToBlob, formatBadgeNumber } from "../../../common/utils/helper";
import { ColorLineRegular } from "@fluentui/react-icons";
import { IDCard } from "./IDCard";
import { WhitelistService, IWhitelistItem } from "../../../common/services/WhitelistService";
import { getSP } from "../../../common/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import { getGraph } from "../../../common/pnpjsConfig";
import { TENANT_DOMAIN, SITES_ESIGN } from "../../../common/constants";

const Specimen: React.FC<ISpecimenProps> = ({ userDisplayName, context }) => {
  if (!context) {
    return <div className="alert alert-danger">WebPart Context is missing/undefined.</div>;
  }
  const {
    items: initialUrl,
    userId,
    updateFile: updateFileInitial,
    loading: initialLoading,
  } = useSpecimens(context, "initial");

  const {
    items: signatureUrl,
    updateFile: updateFileSignature,
    loading: signatureLoading,
  } = useSpecimens(context, "signature");

  const [editInitial, setEditInitial] = useState<boolean>(false);
  const [editSignature, setEditSignature] = useState<boolean>(false);
  // ID Card State
  const [badgeNumber, setBadgeNumber] = useState<string>("");
  const [phoneNumber, setPhoneNumber] = useState<string>("");
  const [waStatus, setWaStatus] = useState<boolean>(false);
  const [whitelistItem, setWhitelistItem] = useState<IWhitelistItem | undefined>(undefined);
  const [cardLoading, setCardLoading] = useState<boolean>(true);
  const [checkLoading, setCheckLoading] = useState<boolean>(false);


  useEffect(() => {
    const fetchData = async (): Promise<void> => {
      setCardLoading(true);
      try {
        const targetUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}`;
        const sp: SPFI = getSP(context, targetUrl);
        const email = context.pageContext.user.email;

        const item = await WhitelistService.getWhitelistItem(
            sp,
            email
        );

        if (item) {
            setWhitelistItem(item);
            setWaStatus(item.Status === "Active");
            setPhoneNumber(item.PhoneNumber);

            setBadgeNumber(
              formatBadgeNumber(
                item.BadgeNumber || ""
              )
          );
        }
        else {

            setWaStatus(false);

        }

      } catch (err) {
        console.error("Error fetching ID Card data:", err);
      } finally {
        setCardLoading(false);
      }
    };

    fetchData().catch(console.error);
  }, [context]);

  const handleWaStatusChange = async (checked: boolean): Promise<void> => {
    if (!whitelistItem) return;

    // Optimistic UI update
    setWaStatus(checked);
    const originalStatus = waStatus;

    try {
      const targetUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}`;
      const sp: SPFI = getSP(context, targetUrl);
      await WhitelistService.updateWhitelistStatus(sp, whitelistItem.Id, checked ? "Active" : "Inactive");

      // Update local item state
      setWhitelistItem({ ...whitelistItem, Status: checked ? "Active" : "Inactive" });

    } catch (err) {
      console.error("Failed to update WA status:", err);
      setWaStatus(originalStatus); // Revert on failure
      alert("Failed to update status. Please try again.");
    }
  };


  const handleSave = (val: string, type: string): void => {
    if (!val) {
      alert("❌ Lengkapi data sebelum menyimpan!");
      return;
    }

    (type === "initial" ? updateFileInitial : updateFileSignature)(
      `${userId}-${type}.png`,
      dataURLToBlob(val),
    )
      .then(() => {
        if (type === "initial") {
          setEditInitial(false);
        } else if (type === "signature") {
          setEditSignature(false);
        }
      })
      .catch((err) => {
        console.error("❌ Gagal mengupdate file:", err);
      });
  };

  const handleCheck = async (): Promise<void> => {
    if (!context.pageContext.user.email) return;
    setCheckLoading(true);
    try {
      const targetUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}`;
      const sp: SPFI = getSP(context, targetUrl);
      const graph: GraphFI = getGraph(context);
      await WhitelistService.addWANotification(sp, graph, context.pageContext.user.email);

      // Update state without reload
      const item = await WhitelistService.getWhitelistItem(sp, context.pageContext.user.email);
      if (item) {
        setWhitelistItem(item);
        setWaStatus(item.Status === "Active");
        setPhoneNumber(item.PhoneNumber);
        setBadgeNumber(formatBadgeNumber(item.BadgeNumber || ""));
      }

    } catch (err) {
      console.error("Failed to send check request:", err);
      alert("Gagal mengirim permintaan pengecekan.");
    } finally {
      setCheckLoading(false);
    }
  };

  const handlePhoneNumberChange = async (newPhoneNumber: string): Promise<void> => {
    if (!whitelistItem) return;

    const originalPhone = phoneNumber;
    setPhoneNumber(newPhoneNumber); // Optimistic update

    try {
      const targetUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}`;
      const sp: SPFI = getSP(context, targetUrl);
      await WhitelistService.updatePhoneNumber(sp, whitelistItem.Id, newPhoneNumber);

      // Update local item state
      setWhitelistItem({ ...whitelistItem, PhoneNumber: newPhoneNumber });

    } catch (err) {
      console.error("Failed to update phone number:", err);
      setPhoneNumber(originalPhone); // Revert on failure
      throw err;
    }
  };

  const handleBadgeNumberChange =
    async (
        newBadgeNumber: string
    ): Promise<void> => {

    if (!whitelistItem) return;

    const targetUrl =
        `${TENANT_DOMAIN}/${SITES_ESIGN}`;

    const sp: SPFI =
        getSP(context, targetUrl);

    await WhitelistService.updateBadgeNumber(
        sp,
        whitelistItem.Id,
        newBadgeNumber
    );

    setBadgeNumber(
        newBadgeNumber
    );

    setWhitelistItem({
        ...whitelistItem,
        BadgeNumber:
            newBadgeNumber
    });
};

  // Responsive styles for specimen boxes
  const specimenBoxStyle = {
    width: "100%",
    maxWidth: "240px",
    minWidth: "160px",
    aspectRatio: "3/2",
    backgroundColor: "#fafafa",
    border: "2px dashed #e0e0e0",
    borderRadius: "12px"
  };

  const specimenBoxActiveStyle = {
    ...specimenBoxStyle,
    border: "2px solid #4CAF50",
    padding: "8px"
  };

  return (
    <section className="container mt-2 mt-md-4 px-2 px-md-3">
      {/* Unified Card Layout: ID Card (Left) + Specimens (Right) */}
      <div className="card shadow-sm border-0" style={{ borderRadius: "16px", overflow: "hidden" }}>
        <div className="row g-0">
          {/* Left Column: ID Card */}
          <div 
            className="col-lg-4 col-md-5 col-12 d-flex align-items-stretch" 
            style={{ backgroundColor: "#f8f9fa" }}
          >
            <div className="p-2 p-md-4 w-100 d-flex justify-content-center align-items-center">
              <IDCard
                displayName={userDisplayName}
                email={context.pageContext.user.email}
                badgeNumber={badgeNumber}
                phoneNumber={phoneNumber}
                waStatus={waStatus}
                onWaStatusChange={handleWaStatusChange}
                onPhoneNumberChange={handlePhoneNumberChange}
                onBadgeNumberChange={handleBadgeNumberChange}
                onCheck={handleCheck}
                loading={cardLoading}
                checkLoading={checkLoading}
                isWhitelisted={!!whitelistItem}
              />
            </div>
          </div>

          {/* Right Column: Specimens */}
          <div className="col-lg-8 col-md-7 col-12">
            <div className="p-2 p-md-4">
              {/* Header */}
              <div className="text-center mb-3 mb-md-4">
                <h5 className="mb-1 mb-md-2 fw-bold" style={{ fontSize: "clamp(1rem, 2.5vw, 1.25rem)" }}>
                  Digital Specimens
                </h5>
                <p className="text-muted small mb-0 d-none d-sm-block">
                  Manage your initial and signature specimens for e-signature.
                </p>
              </div>

              {/* Specimens Row */}
              <div className="row g-3 g-md-1">
                <div className="col-12 col-md-6">
                  <div className="text-center h-100">
                    <h6 className="fw-bold mb-1 mb-md-2" style={{ fontSize: "clamp(0.75rem, 2vw, 1rem)" }}>
                      Initial
                    </h6>
                    <div className="d-flex flex-column align-items-center h-100">
                      {editInitial ? (
                        <div className="my-1 my-md-2 w-100 d-flex justify-content-center">
                          <SignatureComponent
                            onBack={(val) => setEditInitial(val)}
                            onSave={(val) => handleSave(val, "initial")}
                            loading={initialLoading}
                          />
                        </div>
                      ) : (
                        <div className="my-1 my-md-2 d-flex justify-content-center w-100">
                          {initialLoading ? (
                            <div
                              className="d-flex justify-content-center align-items-center rounded"
                              style={specimenBoxStyle}
                            >
                              <div className="spinner-border spinner-border-sm text-primary" role="status">
                                <span className="visually-hidden">Loading...</span>
                              </div>
                            </div>
                          ) : !initialUrl ? (
                            <div
                              className="d-flex flex-column justify-content-center align-items-center rounded"
                              style={specimenBoxStyle}
                            >
                              <p className="small text-muted mb-0" style={{ fontSize: "clamp(0.65rem, 1.5vw, 0.875rem)" }}>
                                No initial set
                              </p>
                            </div>
                          ) : (
                            <div
                              className="d-flex justify-content-center align-items-center rounded"
                              style={specimenBoxActiveStyle}
                            >
                              <img
                                src={initialUrl}
                                alt="Initial"
                                className="img-fluid"
                                style={{ maxHeight: "100%", maxWidth: "100%", borderRadius: "8px", objectFit: "contain" }}
                              />
                            </div>
                          )}
                        </div>
                      )}
                      {!editInitial && (
                        <button
                          type="button"
                          className="btn btn-outline-primary btn-sm mt-1 mt-md-2 px-2 px-md-3"
                          style={{ fontSize: "clamp(0.65rem, 1.5vw, 0.875rem)", whiteSpace: "nowrap" }}
                          onClick={() => {
                            setEditInitial(true);
                          }}
                        >
                          <ColorLineRegular fontSize={14} className="me-1" /> 
                          <span className="d-none d-sm-inline">Update </span>Initial
                        </button>
                      )}
                    </div>
                  </div>
                </div>

                <div className="col-12 col-md-6">
                  <div className="text-center h-100">
                    <h6 className="fw-bold mb-1 mb-md-2" style={{ fontSize: "clamp(0.75rem, 2vw, 1rem)" }}>
                      Signature
                    </h6>
                    <div className="d-flex flex-column align-items-center h-100">
                      {editSignature ? (
                        <div className="my-1 my-md-2 w-100 d-flex justify-content-center">
                          <SignatureComponent
                            onBack={(val) => setEditSignature(val)}
                            onSave={(val) => handleSave(val, "signature")}
                            loading={signatureLoading}
                          />
                        </div>
                      ) : (
                        <div className="my-1 my-md-2 d-flex justify-content-center w-100">
                          {signatureLoading ? (
                            <div
                              className="d-flex justify-content-center align-items-center rounded"
                              style={specimenBoxStyle}
                            >
                              <div className="spinner-border spinner-border-sm text-primary" role="status">
                                <span className="visually-hidden">Loading...</span>
                              </div>
                            </div>
                          ) : !signatureUrl ? (
                            <div
                              className="d-flex flex-column justify-content-center align-items-center rounded"
                              style={specimenBoxStyle}
                            >
                              <p className="small text-muted mb-0" style={{ fontSize: "clamp(0.65rem, 1.5vw, 0.875rem)" }}>
                                No signature set
                              </p>
                            </div>
                          ) : (
                            <div
                              className="d-flex justify-content-center align-items-center rounded"
                              style={specimenBoxActiveStyle}
                            >
                              <img
                                src={signatureUrl}
                                alt="Signature"
                                className="img-fluid"
                                style={{ maxHeight: "100%", maxWidth: "100%", borderRadius: "8px", objectFit: "contain" }}
                              />
                            </div>
                          )}
                        </div>
                      )}
                      {!editSignature && (
                        <button
                          type="button"
                          className="btn btn-outline-primary btn-sm mt-1 mt-md-2 px-2 px-md-3"
                          style={{ fontSize: "clamp(0.65rem, 1.5vw, 0.875rem)", whiteSpace: "nowrap" }}
                          onClick={() => {
                            setEditSignature(true);
                          }}
                        >
                          <ColorLineRegular fontSize={14} className="me-1" /> 
                          <span className="d-none d-sm-inline">Update </span>Signature
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              {/* Mobile Notice */}
              <div className="text-center mt-3 mt-md-4">
                <p className="text-muted small mb-0" style={{ fontSize: "clamp(0.65rem, 1.5vw, 0.8rem)" }}>
                  <span role="img" aria-label="mobile">📱</span> Update easier with mobile device
                </p>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
};

export default Specimen;
