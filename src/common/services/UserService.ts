import { GraphFI } from "@pnp/graph";
import "@pnp/graph/users";

/**
 * Service for handling user-related operations
 */
export class UserService {
    /**
     * Fetches the current user's Employee ID from Microsoft Graph (Entra ID).
     * Returns "-" if not found.
     * 
     * @param graph - PnPjs GraphFI instance
     * @returns Promise resolving to the employee ID or "-"
     */
    public static async getCurrentUserEmployeeId(graph: GraphFI): Promise<string> {
        try {
            console.log("Fetching Employee ID via Graph...");
            const user = await graph.me.select("displayName", "employeeId", "onPremisesExtensionAttributes")();
            console.log("FULL GRAPH USER", user);

            return (
                user.employeeId || 
                user.onPremisesExtensionAttributes?.extensionAttribute1 ||
                "-"
            );

        } catch (error) {
            console.error("Failed to fetch employee ID from Graph", error);
            return "-";
        }
    }

    public static async getUserEmployeeIdByEmail(
    graph: GraphFI,
    email: string
    ): Promise<string> {
        try {

            const user = await graph.users
                .getById(email)
                .select(
                    "displayName",
                    "employeeId",
                    "onPremisesExtensionAttributes"
                )();

                console.log("LOOKUP USER:", email);
                console.log("GRAPH APPROVER USER:", user);

            return (
                user.employeeId ||
                user.onPremisesExtensionAttributes?.extensionAttribute1 ||
                ""
            );

        } catch (error) {
            console.error(
                `Failed to fetch Employee ID for ${email}`,
                error
            );
            return "";
        }
    }

    public static async searchUsers(
    graph: GraphFI,
    filterText: string
    ): Promise<any[]> {

        try {

            const users = await graph.users();

            return users.filter(
                u =>
                    u.displayName
                        ?.toLowerCase()
                        .includes(filterText.toLowerCase())
            );

        } catch (error) {

            console.error(error);

            return [];
        }
    }
}
