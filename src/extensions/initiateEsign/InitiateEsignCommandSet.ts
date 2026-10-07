import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";
import { getSP } from "../../common/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import ConfirmationDialog from "./components/ConfirmationDialog";
import SignatureLogDialog from "./components/SignatureLogDialog";
import { TENANT_DOMAIN, SITES_ESIGN, LIST_PROCESS, LIST_TASKS, LIST_ACTIVE_SITES, PAGE_INITIATE_ESIGN } from "../../common/constants";

interface IActiveSiteConfig {
  Title: string; // Site relative URL/path
  statusSites: string; // "Active" | "Inactive"
  allDirectory: string; // "True" | "False" (Choice)
  directoryContainers: string; // Long text: "Lib1; Lib2"
}

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IInitiateEsignCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = "InitiateEsignCommandSet";

export default class InitiateEsignCommandSet extends BaseListViewCommandSet<IInitiateEsignCommandSetProperties> {
  private _activeSites: IActiveSiteConfig[] = [];

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized InitiateEsignCommandSet");

    // // initial state of the command's visibility
    // const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
    // compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(
      this,
      this._onListViewStateChanged
    );

    // Fetch Active Sites Configuration
    const esignSiteUrl = `${TENANT_DOMAIN}/${SITES_ESIGN}`;
    const spEsign = getSP(this.context, esignSiteUrl);

    return spEsign.web.lists.getByTitle(LIST_ACTIVE_SITES).items
      .select("Title,statusSites,allDirectory,directoryContainers")
      .filter("statusSites eq 'Active'")()
      .then((items: IActiveSiteConfig[]) => {
        this._activeSites = items;

        this._updateCommandVisibility(); // Re-run logic immediately after load
        this.raiseOnChange();
      })
      .catch((err) => {
        console.error(LOG_SOURCE, "Failed to load Active Sites Config", err);
      });
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "COMMAND_1": {
        // Prevent execution if still checking
        const command = this.tryGetCommand("COMMAND_1");
        if (command && command.title === "Checking...") {
          return;
        }

        const selectedItem = event.selectedRows[0];
        if (!selectedItem) {
          console.log("Error: Tidak ada file yang dipilih.");
          return;
        }
        const fileRef = selectedItem.getValueByName("FileRef");
        // const webUrl = this.context.pageContext.web.absoluteUrl;
        const webUrlInitiateEsign = `${TENANT_DOMAIN}/${SITES_ESIGN}`;
        const webUrlNow = this.context.pageContext.web.absoluteUrl;

        if (fileRef) {
          // Assuming the page is in the same site's SitePages
          window.location.href = `${webUrlInitiateEsign}/SitePages/${PAGE_INITIATE_ESIGN}?fileRef=${encodeURIComponent(fileRef)}&source=${encodeURIComponent(webUrlNow)}`;
        } else {
          console.error("Could not retrieve file information");
        }
        break;
      }
      case "COMMAND_2": {
        const selectedItem = event.selectedRows[0];
        if (!selectedItem) return;
        const fileRef = selectedItem.getValueByName("FileRef");
        const fileName = selectedItem.getValueByName("FileLeafRef");

        const sp = getSP(this.context, `${TENANT_DOMAIN}/${SITES_ESIGN}`);
        const logDialog = new SignatureLogDialog(fileRef, sp, this.context, `Approval Log - ${fileName}`);
        logDialog.show().catch(console.error);
        break;
      }
      case "COMMAND_3": {
        const selectedItem = event.selectedRows[0];
        if (!selectedItem) return;
        const fileRef = selectedItem.getValueByName("FileRef");
        const currentUserEmail = this.context.pageContext.user.email;

        const sp = getSP(this.context, `${TENANT_DOMAIN}/${SITES_ESIGN}`);

        sp.web.lists.getByTitle(LIST_PROCESS).items
          .select("Id", "Requestor/EMail", "Requestor/Title", "Status")
          .expand("Requestor")
          .filter(`FileRef0 eq '${fileRef.replace(/'/g, "''")}' and Status eq 'Processing'`)()
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          .then((items: any[]) => {
            if (items.length > 0) {
              const item = items[0];
              const requestorEmail = item.Requestor ? item.Requestor.EMail : "";
              const requestorName = item.Requestor ? item.Requestor.Title : "Unknown";

              if (requestorEmail.toLowerCase() === currentUserEmail.toLowerCase()) {

                const fileName = selectedItem.getValueByName("FileLeafRef");

                const confirmDialog = new ConfirmationDialog(
                  `Are you sure you want to cancel the request for "${fileName}"? This action cannot be undone.`,
                  "Cancel Request Confirmation"
                );

                confirmDialog.show().then(() => {
                  if (confirmDialog.isConfirmed) {
                    // Update to Canceled
                    sp.web.lists.getByTitle(LIST_PROCESS).items.getById(item.Id).update({
                      Status: "Canceled"
                    }).then(() => this._cancelPendingTasks(sp, item.Id)).then(() => {
                      Dialog.alert("Requests successfully canceled.").then(() => {
                        window.location.reload();
                      }).catch(console.error);
                    }).catch((err) => {
                      console.error(err);
                      Dialog.alert("Failed to cancel request.").catch(console.error);
                    });
                  }
                }).catch(console.error);
              } else {
                Dialog.alert(`You are not authorized to cancel this request. Current Requestor is: ${requestorName}`).catch(console.error);
              }
            } else {
              Dialog.alert("No active workflow found to cancel.").catch(console.error);
            }
          }).catch(err => {
            console.error(err);
            Dialog.alert("Error checking workflow status.").catch(console.error);
          });
        break;
      }
      default:
        throw new Error("Unknown command");
    }
  }

  private _onListViewStateChanged = (
    args: ListViewStateChangedEventArgs
  ): void => {
    Log.info(LOG_SOURCE, "List view state changed");
    this._updateCommandVisibility();
  };

  private _updateCommandVisibility(): void {
    const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
    const compareTwoCommand: Command = this.tryGetCommand("COMMAND_2");
    const compareThreeCommand: Command = this.tryGetCommand("COMMAND_3");
    const isSelectedRowCount = this.context.listView.selectedRows?.length ?? 0;
    const selectedRow = this.context.listView.selectedRows![0];
    const fileName = selectedRow ? selectedRow.getValueByName("FileLeafRef") : "";
    const listTitle = this.context.pageContext.list?.title || "";

    // Show only if:
    // 1. Matches "Active Sites" configuration
    // 2. 1 item is selected
    // 3. It is a PDF

    let isAllowedLocation = false;
    const currentWebUrl = this.context.pageContext.web.serverRelativeUrl.toLowerCase();
    const currentListTitleLower = listTitle.toLowerCase();

    // Normalization Logic
    // If starts with /sites/, take 2 segments. Else take 1 segment.
    const urlParts = currentWebUrl.split('/').filter(p => p);
    let siteUrlMatches = "";
    if (urlParts.length > 0 && urlParts[0] === "sites") {
      siteUrlMatches = "/" + urlParts.slice(0, 2).join("/");
    } else if (urlParts.length > 0) {
      siteUrlMatches = "/" + urlParts[0];
    }

    for (const siteConfig of this._activeSites) {
      // Rule 1: Site Match
      // Compare normalized URL with Config Title (case-insensitive)
      if (siteUrlMatches === siteConfig.Title.toLowerCase()) {

        // Rule 2: All Directory
        // Since it is a Choice column (Option), the value is a string "True" or "False"
        if (siteConfig.allDirectory === "True") {
          isAllowedLocation = true;
          break;
        }

        // Rule 3: Directory Containers
        if (siteConfig.allDirectory === "False" && siteConfig.directoryContainers) {
          const allowedLibs = siteConfig.directoryContainers.split(';').map(s => s.trim().toLowerCase()).filter(s => s);

          // Check if currentListTitle contains ANY of the allowedLib strings
          const isMatch = allowedLibs.some(lib => currentListTitleLower.includes(lib));

          if (isMatch) {
            isAllowedLocation = true;
            break;
          }
        }
      }
    }

    const shouldShow =
      isAllowedLocation &&
      isSelectedRowCount === 1 &&
      fileName &&
      fileName.toLowerCase().endsWith(".pdf");

    // Initially show commands if conditions met
    if (shouldShow) {
      // 1. Show "Checking..." state using Command 1
      if (compareOneCommand) {
        compareOneCommand.visible = true;
        compareOneCommand.title = "Checking...";
        // Use a spinner icon if possible. Since we can't easily use Fabric names in CommandSet without hacking,
        // we'll stick to text or use a generic 'sync' base64 if we had one.
        // For now, text "Checking..." is the clearest indicator.
        // Optional: We could set a temporary icon if we had a data URI for a spinner.
      }
      if (compareTwoCommand) compareTwoCommand.visible = true;
      if (compareThreeCommand) compareThreeCommand.visible = false;

      this.raiseOnChange();

      // 2. Perform Async Check
      const fileRef = selectedRow.getValueByName("FileRef");
      this._checkExistingWorkflow(fileRef).then((exists) => {
        if (exists) {
          // Workflow exists: Hide Request (Cmd 1), Show Cancel (Cmd 3)
          if (compareOneCommand) compareOneCommand.visible = false;
          if (compareThreeCommand) compareThreeCommand.visible = true;
        } else {
          // No workflow: Show Request (Cmd 1), Hide Cancel (Cmd 3)
          if (compareOneCommand) {
            compareOneCommand.visible = true;
            compareOneCommand.title = "Request Signature"; // Restore Title
          }
          if (compareThreeCommand) compareThreeCommand.visible = false;
        }
        this.raiseOnChange();
      }).catch((error) => {
        console.error("Error checking existing workflow:", error);
        // Fallback: Show Request button on error
        if (compareOneCommand) {
          compareOneCommand.visible = true;
          compareOneCommand.title = "Request Signature";
        }
        this.raiseOnChange();
      });
    } else {
      // Hide all if selection invalid
      if (compareOneCommand) compareOneCommand.visible = false;
      if (compareTwoCommand) compareTwoCommand.visible = false;
      if (compareThreeCommand) compareThreeCommand.visible = false;
      this.raiseOnChange();
    }
  }

  /**
   * Marks all Pending approver tasks of a process as Canceled so they no longer
   * appear as work to do. Failures are logged only: the approval form also checks
   * the process status, so approvers stay blocked even if this update fails.
   */
  private async _cancelPendingTasks(sp: SPFI, processId: number): Promise<void> {
    try {
      const tasks = await sp.web.lists.getByTitle(LIST_TASKS).items
        .select("Id", "ProcessID/Id")
        .expand("ProcessID")
        .filter(`ProcessID/Id eq ${processId} and Status eq 'Pending'`)();

      await Promise.all(tasks.map((task: { Id: number }) =>
        sp.web.lists.getByTitle(LIST_TASKS).items.getById(task.Id).update({
          Status: "Canceled",
          Comments: "Canceled by requestor"
        })
      ));
    } catch (e) {
      console.warn(LOG_SOURCE, "Failed to cancel pending tasks", e);
    }
  }

  private async _checkExistingWorkflow(fileRef: string): Promise<boolean> {
    try {
      const sp: SPFI = getSP(this.context, `${TENANT_DOMAIN}/${SITES_ESIGN}`);
      const items = await sp.web.lists.getByTitle(LIST_PROCESS).items
        .select("FileRef0", "HashHex", "Status")
        .filter(`FileRef0 eq '${fileRef.replace(/'/g, "''")}' and Status eq 'Processing'`)
        ();
      return items.length > 0;
    } catch (e) {
      console.error("Failed to check existing workflow", e);
      return false;
    }
  }
}
