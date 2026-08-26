import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { LIST_APPROVAL_MAP } from "../constants";

export interface IApprovalMapItem {
  Id: number;
  Title: string; // Role / Description
  Approver: {
    Id: number;
    Title: string;
    EMail: string;
  };
  AuthorizedApprovers: string;
  Secretaries: string;
}

export class ApprovalService {
  private _sp: SPFI;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  // Translate UPN guest (user_domain.com#EXT#@tenant.onmicrosoft.com) jadi email asli
  private upnToEmail(upn: string): string {
    const match = upn.match(/^(.+)#ext#@/i);
    if (!match) return upn;
    const raw = match[1];
    const i = raw.lastIndexOf("_");
    return i !== -1 ? raw.substring(0, i) + "@" + raw.substring(i + 1) : raw;
  }

  // Ambil email dari person field: EMail kalau ada (internal), else parse dari Name (guest)
  private resolveEmail(user: { EMail?: string; Name?: string } | undefined): string {
    if (!user) return "";
    if (user.EMail) return user.EMail;
    if (!user.Name) return "";
    const claim = user.Name.split("|").pop() ?? "";
    return this.upnToEmail(claim);
  }

  public async getApprovalMap(): Promise<IApprovalMapItem[]> {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const items: any[] = await this._sp.web.lists
        .getByTitle(LIST_APPROVAL_MAP)
        .items.select(
          "Id",
          "Title",
          "Official/Id",
          "Official/Title",
          "Official/EMail",
          "Official/Name",
          "onBehalf/EMail",
          "onBehalf/Name",
          "Secretary/EMail",
          "Secretary/Name",
        )
        .expand("Official", "onBehalf", "Secretary")
        .top(5000)();

      console.log("APPROVAL MAP RAW DATA:", items);

      return items
        .filter((item) => item.Official && item.Official.Id) // cukup cek Id, jangan EMail
        .map((item) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const onBehalfEmails = item.onBehalf
            ? item.onBehalf.map((u: any) => this.resolveEmail(u)).filter(Boolean).join(";")
            : "";
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const secretaryEmails = item.Secretary
            ? item.Secretary.map((u: any) => this.resolveEmail(u)).filter(Boolean).join(";")
            : "";
          return {
            Id: item.Id,
            Title: item.Title,
            Approver: {
              Id: item.Official.Id,
              Title: item.Official.Title,
              EMail: this.resolveEmail(item.Official),
            },
            AuthorizedApprovers: onBehalfEmails,
            Secretaries: secretaryEmails,
          };
        });
    } catch (err) {
      console.error("Error fetching Approval Map:", err);
      return [];
    }
  }
}