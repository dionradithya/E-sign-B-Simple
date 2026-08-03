# 📝 eSign-SPFX

## 📋 Summary

This project is a SharePoint Framework (SPFx) solution designed to handle electronic signature workflows. It includes a Form Customizer extension for initiating and handling signatures on PDF documents directly within SharePoint lists, and likely a Web Part for initiating the process.

## 🛠 Prerequisites

- SharePoint Online Tenant
- **SPFx Version**: v1.22.2
- **Node.js**: v22.14.0 (Recommended: >=22.14.0 < 23.0.0)
- **PnP JS**: v4.14.0
- **Yeoman** (yo) & **SharePoint Generator** (@microsoft/generator-sharepoint)
- **PnP PowerShell** (required for associated list forms)

## 🚀 Installation

1.  **Clone the repository**:
    ```bash
    git clone <repository-url>
    cd esign-spfx
    ```

2.  **Install dependencies**:
    ```bash
    npm install
    ```

3.  **Trust user-created certificate (for local development)**:
    ```bash
    npm run trust-dev-cert # If configured, otherwise typically 'gulp trust-dev-cert' is legacy. 
    # For Heft projects, this is often handled differently or via 'heft trust-dev-cert' if available, 
    # but strictly speaking 'npm install' might handle certificates if part of postinstall.
    # Note: If 'gulp' is completely removed, certificate trust might be:
    node node_modules/@microsoft/sp-build-web/lib/cert/cert.js 
    # OR simply proceed if already trusted.
    ```

## 🛠️ Build & Development Commands (Heft & NPM)

This project uses **Heft** for building and bundling, not Gulp. We have configured **NPM aliases** for convenience.

| Task | NPM Command | Underlying Heft Command | Description |
| :--- | :--- | :--- | :--- |
| **Serve** | `npm run serve` | `heft start --clean` | Starts the local dev server. |
| **Build (Dev)** | `npm run build` | `heft build --clean` | Builds the project in dev mode. |
| **Build (Prod)** | `npm run build:ship` | `heft build --clean --production` | Builds optimized assets for release. |
| **Package** | `npm run package` | `heft build ... && heft package-solution ...` | Builds and creates the `.sppkg` file. |
| **Clean** | `npm run clean` | `heft clean` | Cleans build artifacts. |
| **Test** | `npm run test` | `heft test --clean` | Runs unit tests. |

## ⚙️ Configuration

Before running or deploying, ensure the constants match your SharePoint environment.
Update **`src/shared/constants.ts`** with your specific URLs and list names:

```typescript
export const TENANT_DOMAIN = "https://yourtenant.sharepoint.com";
export const SITES_ESIGN = "sites/YourSite";
export const LIST_PROCESS = "Approvals Process";
export const LIST_TASKS = "Tasks";
export const LIST_ACTIVITY_LOG = "Activity Log";
export const SITE_REDIRECT = "sites/YourSite";
export const DATABASE_SPECIMEN = "Specimen";
export const LIST_APPROVAL_MAP = "Approval Map";
export const LIST_ACTIVE_SITES = "Active Sites";
export const PAGE_INITIATE_ESIGN = "RequestSignatures.aspx";
export const LIST_WHITELIST_WA = "Whitelist WA Notification";
```

## 🗄 SharePoint Setup
 
You need to create the following Lists and Libraries in your SharePoint site.
 
### 1. 📂 Document Library: `Specimen`
Stores user signature and initial images.
*   **Name**: `Specimen` (URL: `/Specimen`)
*   **Columns**: Standard Document Library columns.
*   **Content**: Images should be named format: `{UserId}-signature.png` or `{UserId}-initial.png`.
 
### 2. 📋 List: `Approvals Process`
Tracks the overall approval workflow for a document.
*   **Name**: `Approvals Process` (or match your `LIST_PROCESS` constant)
*   **Columns**:
    *   `Title` (Text): Unique Code/Process Title.
    *   `FileRef0` (Text): Server Relative URL of the PDF file.
    *   `ProcessData` (Multiple lines of text - Plain text): JSON string containing signature placeholder coordinates and signer details.
    *   `HashHex` (Text): Hash of the original PDF document for integrity checks.
    *   `Status` (Text): Workflow status (e.g., "Pending", "Canceled").
    *   `Notes` (Multiple lines of text): System or admin notes.
    *   `Requestor` (Person or Group): The user who initiated the request.
 
### 3. ✅ List: `Tasks`
Contains individual approval tasks assigned to users. This is where the Form Customizer is typically applied.
*   **Name**: `Tasks` (or match your `LIST_TASKS` constant)
*   **Columns**:
    *   `Title` (Text): Task Title.
    *   `ProcessID` (Lookup): Lookup to **Approvals Process** list (Select Title/ID).
    *   `AssignedTo` (Person or Group): The approver(s) for this task.
    *   `Status` (Text): Task status (e.g., "Pending", "Approved", "Rejected", "Expired", "Canceled").
    *   `StepApprover` (Number): Step sequence number.
    *   `TotalApprover` (Number): Total number of steps/approvers.
    *   `Comments` (Multiple lines of text): Approver's comments.
    *   `ResponseData` (Multiple lines of text): JSON data containing client info (IP, Agent).
    *   `executedById` (Person or Group): The actual user who executed the action (in case of delegation or admin override).
 
### 4. 📜 List: `Activity Log`
Audit log for all actions taken.
*   **Name**: `Activity Log` (or match your `LIST_ACTIVITY_LOG` constant)
*   **Columns**:
    *   `Activity` (Text): Description of the action (e.g., "Approved Document1.pdf").
    *   `UserId` (Person or Group): The user who performed the action.
    *   `ProcessTitle` (Text): Title of the related process.

### 5. 🗺️ List: `Approval Map`
Mapping configuration for approval routing.
*   **Name**: `Approval Map` (or match your `LIST_APPROVAL_MAP` constant)
*   **Columns**:
    *   `Title` (Text): Role Name or Description (e.g., "Manager", "Director").
    *   `Official` (Person or Group): The primary approver for this role.
    *   `onBehalf` (Person or Group - Allow Multiple): Users authorized to approve on behalf of the Official.

### 6. 🌐 List: `Active Sites`
Registry of active sites where the eSign solution is enabled.
*   **Name**: `Active Sites` (or match your `LIST_ACTIVE_SITES` constant)
*   **Columns**:
    *   `Title` (Text): Server relative URL of the site/subsite (e.g., `/sites/HR`).
    *   `statusSites` (Choice): "Active" or "Inactive". Only "Active" sites are processed.
    *   `allDirectory` (Choice): "True" or "False". If "True", all libraries in the site are enabled.
    *   `directoryContainers` (Text): Semicolon-separated list of document library names (e.g., `Documents; Contracts`) to enable if `allDirectory` is "False".

### 7. 📱 List: `Whitelist WA Notification`
Phone numbers whitelisted for WhatsApp notifications.
*   **Name**: `Whitelist WA Notification` (or match your `LIST_WHITELIST_WA` constant)
*   **Columns**:
    *   `Title` (Text): Standard Title column (Optional/Default).
    *   `Email` (Person or Group): **Mandatory**. The user to be notified. Enforce unique values if possible.
    *   `Phone Number` (Text): **Mandatory**. Internal name `Phone_x0020_Number`. Stores the mobile number.
    *   `Status` (Choice): **Mandatory**. Options: `Active`, `Inactive`. Default: `Inactive`.

## 🔗 Associating the Form Customizer

After deploying the solution (or for testing), you must associate the **Tasks** list forms with the SPFx Form Customizer component. You can do this easily using **PnP PowerShell**.

1.  **Connect to your site**:
    ```powershell
    Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/YourSite" -Interactive
    ```

2.  **Verify your List Content Type**:
    ```powershell
    Get-PnPContentType -List "Tasks" | Select-Object Name, StringId
    # Note the ID for "Item" (usually starts with 0x0100...)
    ```

3.  **Associate the Component**:
    Run the following script to force the New, Edit, and Display forms to use the SPFx component.

    ```powershell
    # 1. Define the Component ID (from manifest)
    $formCustomizerId = "c87d879a-9cb0-467b-81c0-34ce5b484d18"

    # 2. Update the "Item" Content Type on the "Tasks" list
    Set-PnPContentType -List "Tasks" -Identity "Item" `
        -NewFormClientSideComponentId $formCustomizerId `
        -EditFormClientSideComponentId $formCustomizerId `
        -DisplayFormClientSideComponentId $formCustomizerId
    ```

4.  **Verify**:
    If successful, navigating to a Task item (New/Edit/View) should now render using the custom SPFx form.

## 📂 Project Structure

The project follows a modular structure to separate concerns and promote reuse:

```
src/
├── common/                     # Shared code, configuration, and utilities
│   ├── assets/                 # Static assets (e.g., pdf.worker.min.js)
│   ├── models/                 # Shared Interfaces and Types
│   ├── services/               # Business Logic Services (EsignDataService, UserService, etc.)
│   ├── utils/                  # Helper utilities (CanvasUtils, PdfUtils, etc.)
│   ├── constants.ts            # Site URLs, List Names, Settings
│   └── pnpjsConfig.ts          # PnP JS Initialization
│
├── extensions/                 # SPFx Extensions
│   ├── initiateEsign/          # ListView Command Set (Context menu actions)
│   └── initiateEsignForm/      # Form Customizer for the 'Tasks' list
│
└── webparts/                   # SPFx Web Parts
    ├── docIntegrityChecker/    # Web Part to validate document hash/integrity
    ├── documentTrackerEsign/   # Web Part to track document status
    ├── esignTaskList/          # Web Part for managing eSign tasks
    ├── initiateSignature/      # Web Part to start the workflow (upload/select PDF)
    └── specimen/               # Web Part for managing user signatures/initials
```

## 🏃‍♂️ Running the Project

To start the local development server:

```bash
npm run serve
```

For the Form Customizer, you may need to use specific debug query parameters to test against a live list item:
`?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&loadSPFX=true`

## 🚢 Deployment

1.  **Bundle and Package**:
    ```bash
    npm run package
    ```
    *(This runs `heft build --clean --production` and then `heft package-solution --production`)*

2.  **Upload**: Upload the `.sppkg` file from `sharepoint/solution` to your App Catalog.
3.  **Deploy**: Deploy the app and ensure it is installed on the target site.
