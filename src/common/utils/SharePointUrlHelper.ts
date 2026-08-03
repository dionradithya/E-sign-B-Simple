import { TENANT_DOMAIN } from "../constants";

/**
 * Helper function to extract site URL from a server-relative file path.
 * Handles different SharePoint path structures:
 * - /sites/SiteName/... -> TENANT_DOMAIN/sites/SiteName (managed path with subsite)
 * - /gensys/... -> TENANT_DOMAIN/gensys (direct site under tenant root)
 * - Everything else -> TENANT_DOMAIN (root site)
 * 
 * @param fileRef - Server relative URL of the file
 * @returns Full site URL
 */
export function getSiteUrlFromPath(fileRef: string): string {
    const pathParts = fileRef.split('/');
    // pathParts[0] = '', pathParts[1] = first segment, pathParts[2] = second segment...

    const firstSegment = pathParts[1]?.toLowerCase();

    // /sites/xxx/... -> managed path with subsite
    if (firstSegment === 'sites' && pathParts[2]) {
        return `${TENANT_DOMAIN}/sites/${pathParts[2]}`;
    }

    // For any other path (e.g., /gensys/..., /teams/...) -> direct site under tenant root
    return firstSegment ? `${TENANT_DOMAIN}/${firstSegment}` : TENANT_DOMAIN;
}
