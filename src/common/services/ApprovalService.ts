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

  public async getApprovalMap(): Promise<IApprovalMapItem[]> {
    try {
      // Fetch Title and User field (expanding User to get details)
      // Internal Name of 'Email' column is 'User'
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const items: any[] = await this._sp.web.lists
        .getByTitle(LIST_APPROVAL_MAP)
        // For Official field: .items.select("Id", "Title", "Official/Id", "Official/Title", "Official/EMail", "onBehalf/EMail")
        // For Official field: .expand("Official", "onBehalf")
        .items.select("Id", "Title", "Official/Id", "Official/Title", "Official/EMail", "onBehalf/EMail",  "Secretary/EMail")
        .expand("Official", "onBehalf", "Secretary")
        .top(5000)();

      // Mapped to interface
      return items
        // For Official field: .filter(item => item.Official && item.Official.Id && item.Official.EMail)
        .filter(item => item.Official && item.Official.Id && item.Official.EMail) // Ensure User (Person) exists
        .map(item => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const onBehalfEmails = item.onBehalf ? item.onBehalf.map((u: any) => u.EMail).join(";") : "";
          const secretaryEmails = item.Secretary ? item.Secretary.map((u: { EMail: string }) => u.EMail).join(";") : "";
          return {
            Id: item.Id,
            Title: item.Title,
            Approver: {
              // For Official field: Id: item.Official.Id, Title: item.Official.Title, EMail: item.Official.EMail
              Id: item.Official.Id,
              Title: item.Official.Title,
              EMail: item.Official.EMail
            },
            AuthorizedApprovers: onBehalfEmails,
            Secretaries: secretaryEmails
          };
        });
    } catch (err) {
      console.error("Error fetching Approval Map:", err);
      return [];
    }
  }
}
