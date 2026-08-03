// hooks/useSpecimens.ts
import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { getSP } from "../../../common/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import { TENANT_DOMAIN, SITES_ESIGN, DATABASE_SPECIMEN } from "../../../common/constants";

// eslint-disable-next-line @typescript-eslint/explicit-module-boundary-types
export const useSpecimens = (context: WebPartContext | undefined, type?: string): { items: string; loading: boolean; getSpecimen: (uid?: number) => Promise<void>; userId: number | undefined; updateFile: (fileName: string, content: Blob) => Promise<void>; } => {
  const baseUrl = TENANT_DOMAIN;
  const sp: SPFI | null = useMemo(
    () => (context ? getSP(context, `${baseUrl}/${SITES_ESIGN}`) : null),
    [context]
  );

  const [items, setItems] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);
  const [userId, setUserId] = useState<number | undefined>(undefined);
  
  // Ref to hold the current blob URL to revoke it when it changes or on unmount
  const itemUrlRef = useRef<string | null>(null);

  const cleanupUrl = (): void => {
    if (itemUrlRef.current) {
      URL.revokeObjectURL(itemUrlRef.current);
      itemUrlRef.current = null;
    }
  };

  // Ensure we revoke the URL when the component unmounts
  useEffect(() => {
    return () => {
      cleanupUrl();
    };
  }, []);

  const getSpecimen = useCallback(async (uid?: number): Promise<void> => {
    // If no uid provided, try to use state (though usually we pass it or rely on it being set)
    const targetId = uid ?? userId;
    if (!targetId) return;

    setLoading(true);
    try {
      const filePath = `/${SITES_ESIGN}/${DATABASE_SPECIMEN}/${targetId}-${type}.png`;
      if (!sp) return;
      const file = sp.web.getFileByServerRelativePath(filePath);
      const blob = await file.getBlob();

      cleanupUrl(); // Cleanup previous URL
      const blobUrl = URL.createObjectURL(blob);
      itemUrlRef.current = blobUrl;
      setItems(blobUrl);
    } catch {
      // It is normal to fail if the file doesn't exist yet (new user)
      // console.warn("Specimen not found or error:", err);
      // Ensure we clear items if fetch failed (e.g. user deleted signature)
      if (items) {
           cleanupUrl();
           setItems("");
      }
    } finally {
      setLoading(false);
    }
  }, [sp, type, userId, items]);

  useEffect(() => {
    let isMounted = true;
    const init = async (): Promise<void> => {
        // Avoid double loading if already loading or set? 
        // But for strict mode or re-mounts, we should be careful.
        // We'll just fetch user then specimen.
        try {
            if (!sp) return;
            const user = await sp.web.ensureUser(context!.pageContext.user.email);
            if (!isMounted) return;
            setUserId(user.Id);
            // Pass user.Id directly to avoid stale state issues in closure
            // We call the logic of getSpecimen but we can't call the callback 
            // easily if it depends on userId state which isn't updated yet.
            // So we copy logic or accept parameters in getSpecimen.
            
            // Re-implementing fetch logic here to ensure sequential execution with correct ID
            setLoading(true);
            try {
                if (!sp) return;
                const filePath = `/${SITES_ESIGN}/${DATABASE_SPECIMEN}/${user.Id}-${type}.png`;
                const file = sp.web.getFileByServerRelativePath(filePath);
                const blob = await file.getBlob();
                if (!isMounted) return;
                
                cleanupUrl();
                const blobUrl = URL.createObjectURL(blob);
                itemUrlRef.current = blobUrl;
                setItems(blobUrl);
            } catch {
                // Ignore 404
            } finally {
                if (isMounted) setLoading(false);
            }

        } catch (err) {
            console.error("❌ Error initializing user:", err);
        }
    };

    init().catch(err => console.error("Init failed", err));

    return () => {
        isMounted = false;
    };
    // We only want this to run on mount effectively, or if context/sp changes
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sp, type]);


  const updateFile = async (
    fileName: string,
    content: Blob
  ): Promise<void> => {
    setLoading(true);
    try {
      // Use current userId from state
      if (!userId) throw new Error("User ID not loaded");
      if (!sp) throw new Error("SP context undefined");

      await sp.web
        .getFolderByServerRelativePath(`/${SITES_ESIGN}/${DATABASE_SPECIMEN}`)
        .files.addUsingPath(fileName, content, { Overwrite: true });

      // Small delay for list item creation
      await new Promise<void>((resolve) => setTimeout(resolve, 500));

      const relativePath = `/${SITES_ESIGN}/${DATABASE_SPECIMEN}/${fileName}`;
      const item = await sp.web
        .getFileByServerRelativePath(relativePath)
        .getItem();

      await item.update({
        Specimen_x0020_Type: type === "initial" ? "Initial" : "Signature",
        Specimen_x0020_OwnerId: userId,
        Format: "Formatted",
      });

      cleanupUrl();
      const blobUrl = URL.createObjectURL(content);
      itemUrlRef.current = blobUrl;
      setItems(blobUrl);
    } catch (err) {
      console.error("❌ Gagal upload atau update metadata:", err);
    } finally {
      setLoading(false);
    }
  };

  return {
    items,
    loading,
    getSpecimen,
    userId,
    updateFile,
  };
};
