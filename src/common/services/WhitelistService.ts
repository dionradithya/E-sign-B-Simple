import { SPFI } from "@pnp/sp";
import { GraphFI } from "@pnp/graph";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { LIST_WHITELIST_WA } from "../constants";

// Define interface locally to handle potential PnPJS version differences safely
interface ILocalUserInfo {
    Id: number;
    data?: {
        Id: number;
    };
}

export interface IWhitelistItem {
    Id: number;
    Title: string;
    Email: { EMail: string };
    Status: "Active" | "Inactive";
    PhoneNumber: string;
    BadgeNumber: string;
}

export class WhitelistService {
    /**
     * Fetches the whitelist item for a given user email.
     * @param sp The SPFI instance
     * @param userEmail The email of the user to search for
     * @returns The whitelist item or null if not found
     */
    public static async getWhitelistItem(
        sp: SPFI,
        userEmail: string
    ): Promise<IWhitelistItem | undefined> {
        try {
            const items = await sp.web.lists
                .getByTitle(LIST_WHITELIST_WA)
                .items.select("Id", "Title", "Email/EMail", "Status", "PhoneNumber", "BadgeNumber")
                .expand("Email")
                .filter(`Email/EMail eq '${userEmail}'`)
                .top(1)();

            return items.length > 0 ? (items[0] as IWhitelistItem) : undefined;
        } catch (err) {
            console.error(
                `[WhitelistService] Error fetching whitelist item for email ${userEmail}:`,
                err
            );
            return undefined;
        }
    }

    /**
     * Updates the status of a whitelist item.
     * @param sp The SPFI instance
     * @param itemId The ID of the item to update
     * @param status The new status ('Active' or 'Inactive')
     */
    public static async updateWhitelistStatus(
        sp: SPFI,
        itemId: number,
        status: "Active" | "Inactive"
    ): Promise<void> {
        try {
            await sp.web.lists
                .getByTitle(LIST_WHITELIST_WA)
                .items.getById(itemId)
                .update({
                    Status: status,
                });
        } catch (err) {
            console.error(`Error updating whitelist status for item ${itemId}:`, err);
            throw err;
        }
    }

    /**
     * Updates the phone number of a whitelist item.
     * @param sp The SPFI instance
     * @param itemId The ID of the item to update
     * @param phoneNumber The new phone number
     */
    public static async updatePhoneNumber(
        sp: SPFI,
        itemId: number,
        phoneNumber: string
    ): Promise<void> {
        try {
            await sp.web.lists
                .getByTitle(LIST_WHITELIST_WA)
                .items.getById(itemId)
                .update({
                    PhoneNumber: phoneNumber,
                });
        } catch (err) {
            console.error(`Error updating phone number for item ${itemId}:`, err);
            throw err;
        }
    }

    public static async updateBadgeNumber(
        sp: SPFI,
        itemId: number,
        badgeNumber: string
    ): Promise<void> {
        try {
            await sp.web.lists
                .getByTitle(LIST_WHITELIST_WA)
                .items.getById(itemId)
                .update({
                    BadgeNumber: badgeNumber,
                });
        } catch (err) {
            console.error(
                `Error updating badge number for item ${itemId}:`,
                err
            );
            throw err;
        }
    }

    /**
     * Adds a new item to the WA Notification list.
     * @param sp The SPFI instance
     * @param graph The GraphFI instance
     * @param email The user email
     */
    public static async addWANotification(
        sp: SPFI,
        graph: GraphFI,
        email: string
    ): Promise<void> {
        try {
            const userResponse = await sp.web.ensureUser(email);

            const user = userResponse as unknown as ILocalUserInfo;
            const userId = user.data ? user.data.Id : user.Id;

            // 1. Check if user already exists in Whitelist WA
            const existingItems = await sp.web.lists
                .getByTitle(LIST_WHITELIST_WA)
                .items
                .filter(`EmailId eq ${userId}`)
                .top(1)();

            // 2. If not exists, create item in Whitelist WA
            if (existingItems.length === 0) {
                const graphUser = await graph.me
                    .select(
                        "mobilePhone",
                        "employeeId",
                        "onPremisesExtensionAttributes"
                    )();

                console.log("GRAPH USER", graphUser);
                console.log("EMPLOYEE ID", graphUser?.employeeId);

                const mobilePhone =
                    graphUser?.mobilePhone || "Not Found";

                const badgeNumber =
                    graphUser?.employeeId ||
                    graphUser?.onPremisesExtensionAttributes?.extensionAttribute1 ||
                    "";

                await sp.web.lists
                    .getByTitle(LIST_WHITELIST_WA)
                    .items
                    .add({
                        EmailId: userId,
                        PhoneNumber: mobilePhone,
                        BadgeNumber: badgeNumber,
                        Status: "Inactive",
                    });

                console.log(
                    `User ${email} added to Whitelist WA.`
                );
            } else {
                console.log(
                    `User ${email} already exists in Whitelist WA.`
                );
            }

            // 3. Check if user already exists in Approval Map
            const existingApprovalMap = await sp.web.lists
                .getByTitle("Approval Map")
                .items
                .filter(`OfficialId eq ${userId}`)
                .top(1)();

            // 4. If not exists, create item in Approval Map
            if (existingApprovalMap.length === 0) {
                await sp.web.lists
                    .getByTitle("Approval Map")
                    .items
                    .add({
                        Title: email,
                        OfficialId: userId
                    });

                console.log(
                    `User ${email} added to Approval Map.`
                );
            } else {
                console.log(
                    `User ${email} already exists in Approval Map.`
                );
            }

        } catch (err) {
            const errorObj = err as { message: string };

            if (
                errorObj?.message?.indexOf("SPDuplicateValuesFoundException") !== -1 ||
                JSON.stringify(errorObj).indexOf("SPDuplicateValuesFoundException") !== -1
            ) {
                console.warn(
                    `User ${email} duplicate detected via exception.`
                );
                return;
            }

            console.error(
                `Error adding WA notification / Approval Map for ${email}:`,
                err
            );

            throw err;
        }
    }
}
