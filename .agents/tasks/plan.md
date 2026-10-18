# Implementation Plan — Batch e-Sign (all-or-nothing) for Finance

This plan turns the fixed DECISIONS into concrete, ordered edits. All paths are relative to the
worktree root `d:\E-Sign\E-sign-B-Simple\.worktrees\batch-esign`. Edit files there, not in the repo root.

## Environment / build facts (discovered during exploration)

- SPFx 1.22.2, TypeScript 5.3, React 17, PnPjs 4.14, Node 22. Build system is **Heft** (not Gulp).
- Build command: `npm run build` (= `heft build --clean`). This runs TypeScript + ESLint and is the
  real verification gate. `npm run test` exists but **no Jest/vitest config is present**, so there are
  no runnable unit tests; do not claim test runs — verify with the build.
- **Quirk (must do first):** the worktree has **no `node_modules`**. `tsconfig.json` extends a rig under
  `./node_modules/@microsoft/spfx-web-build-rig/...`, so the build fails until dependencies are installed
  in the worktree. Run `npm install` in the worktree before the first build. (Confirmed: parent repo has
  `node_modules`, worktree does not; building against the parent binary fails on the `extends` path.)
- `getSP(context, webUrl?)` from `src/common/pnpjsConfig.ts` returns an `SPFI`. Cross-site reads use
  `getSP(context, "https://bumiresources.sharepoint.com/sites/DMS")` style base URLs.
- OData single-quote escaping convention used everywhere: `value.replace(/'/g, "''")`. Keep it.

## New SharePoint columns (created MANUALLY by the owner — never provisioned from code)

On list `Approvals Process` (constant `LIST_PROCESS`):
- `BatchId` (Text, internal name `BatchId`): a GUID grouping all documents of one batch. Empty for
  normal single requests.
- `BatchMode` (Choice, internal name `BatchMode`): values `None` and `AllOrNothing`, default `None`.

Code must be tolerant: select these fields defensively and guard reads so a missing/empty column never
throws and never changes the existing single-request flow. Treat empty / missing / `None` as "not a batch".

### What BatchMode is for (owner asked)
`BatchMode` labels how a Process item participates in a batch, so the reject logic and reporting can tell
batch members apart from ordinary single requests:
- `None` (default): a normal standalone request. No sibling coupling. This is every existing item and the
  single-flow default — behavior is unchanged.
- `AllOrNothing`: the item belongs to a batch (identified by `BatchId`). If any sibling is **rejected**,
  the other siblings still in `Processing` are set to `Canceled` (all-or-nothing). Siblings already
  `Approved` are left as `Approved`; already `Rejected`/`Canceled` are left untouched. A batch "succeeds"
  only when every member ends `Approved`.

## Query-param contract for the initiate page

- Single flow (unchanged): `RequestSignatures.aspx?fileRef=<uriEncoded serverRelativeUrl>&source=<uriEncoded webUrl>`.
- Batch flow (new): `RequestSignatures.aspx?fileRefs=<uriEncoded JSON string array>&source=<uriEncoded webUrl>`.
  - `fileRefs` value = `encodeURIComponent(JSON.stringify(arrayOfServerRelativeUrls))`.
  - `InitiateSignature.tsx` reads `fileRefs` first; if present and parses to >=1 entries, it runs batch
    mode. Otherwise it falls back to the existing `fileRef` single-file path with zero behavior change.

## createProcessForDocument helper (shared by single + batch save)

Extract the body of the existing `handleSaveWorkflow` per-document work into one async helper so the batch
loop reuses it. Suggested signature (inside `InitiateSignature.tsx`, or a small local module imported by it):

```ts
interface ICreateProcessParams {
  spDoc: SPFI;            // SPFI for the source site where the PDF lives
  spList: SPFI;           // SPFI for the eSign site (LIST_PROCESS, LIST_ACTIVITY_LOG)
  context: WebPartContext;
  fileRef: string;        // server-relative URL of this document's PDF
  fileBlob: Blob;         // the loaded PDF bytes for this document
  approvers: IApprover[]; // this document's approvers
  placeholders: ISignaturePlaceholder[]; // this document's placeholders
  approvalMap: IApprovalMapItem[];
  reviewerCanDownload: boolean;
  embedQrCode: boolean;
  currentUserId: number;  // resolved once via spList.web.currentUser()
  batch?: { batchId: string; batchMode: "AllOrNothing" }; // omitted/undefined for single flow
}

// Returns the created process Title (REQ-ESIGN-...) so the caller can log / surface progress.
async function createProcessForDocument(params: ICreateProcessParams): Promise<string>
```

The helper contains exactly today's steps: generate `REQ-ESIGN-...` Title, optional QR embed +
`setContentChunked` overwrite, `generatePDFHash`, build `ProcessSignature` JSON, `add` to `LIST_PROCESS`,
`add` to `LIST_ACTIVITY_LOG`. When `params.batch` is set it ALSO writes `BatchId` and `BatchMode` on the
`LIST_PROCESS` item; when it is undefined the add payload is byte-for-byte the current single-flow payload.

### BatchId/BatchMode write points
- Written once per document in `createProcessForDocument` on the `LIST_PROCESS.items.add({...})` call,
  only when `params.batch` is provided: `BatchId: params.batch.batchId, BatchMode: params.batch.batchMode`.

### BatchId/BatchMode read points
- `InitiateEsignForm.tsx` `_submitReject`: after the reject succeeds, read the current process's `BatchId`
  and `BatchMode` from `LIST_PROCESS` (defensive select), then run sibling cancellation when
  `BatchMode === "AllOrNothing"` and `BatchId` is non-empty.

---

## Ordered items

- [ ] 1. Install worktree dependencies so the build can run.
      Run `npm install` inside the worktree. This is required once because the worktree has no
      `node_modules` and `tsconfig.json` extends a rig package resolved under `./node_modules`.
      Files: (none — environment setup)
      Verify: `npm run build` reaches the TypeScript compile stage (no "extends ... cannot be resolved"
      error). A clean baseline build should complete with exit code 0.

- [ ] 2. Add the `LIST_FINANCE_PATH` constant.
      Add `export const LIST_FINANCE_PATH = "/sites/DMS/Finance";` to `src/common/constants.ts`.
      Files: src/common/constants.ts
      Verify: `npm run build` succeeds; constant is importable.

- [ ] 3. Declare `COMMAND_BATCH` in the command-set manifest.
      In the `items` map, add a `COMMAND_BATCH` entry with `"title": { "default": "Request Signature (Batch)" }`,
      a `"type": "command"`, and an `iconImageUrl` (reuse an existing data-URI icon, e.g. COMMAND_1's, to
      avoid a new asset). Do NOT touch COMMAND_1/COMMAND_2/COMMAND_3 entries.
      Files: src/extensions/initiateEsign/InitiateEsignCommandSet.manifest.json
      Verify: `npm run build` succeeds (manifest schema valid).

- [ ] 4. Implement `COMMAND_BATCH` visibility in the command set (Finance-only, mutually exclusive with COMMAND_1).
      In `InitiateEsignCommandSet.ts`: import `LIST_FINANCE_PATH` and `SITE_REDIRECT`. In
      `_updateCommandVisibility()`, after the existing COMMAND_1/2/3 logic (do NOT alter it), compute a
      `COMMAND_BATCH` Command via `tryGetCommand("COMMAND_BATCH")`. Detect Finance library by checking
      `this.context.pageContext.web.serverRelativeUrl.toLowerCase().indexOf(LIST_FINANCE_PATH.toLowerCase()) === 0`
      (and/or current list root folder path), case-insensitive. Show COMMAND_BATCH only when: in Finance
      library AND ((selected PDFs count >= 2) OR (exactly one selected row is a folder)). Detect folders
      via the row (e.g. `getValueByName("FSObjType") === "1"`, with a fallback of a blank/zero-length
      `FileLeafRef` extension); detect PDFs via `FileLeafRef` ending `.pdf`. When exactly 1 PDF is selected
      in Finance, keep COMMAND_BATCH hidden (COMMAND_1 owns that). COMMAND_BATCH must be `visible = false`
      whenever location is not Finance. Call `this.raiseOnChange()` as the existing code does.
      Files: src/extensions/initiateEsign/InitiateEsignCommandSet.ts
      Verify: `npm run build` succeeds. Manual: in a Finance library, selecting >=2 PDFs or 1 folder shows
      the batch command; selecting 1 PDF shows COMMAND_1 only; outside Finance it never appears.

- [ ] 5. Implement `COMMAND_BATCH` onExecute (collect/expand PDFs, redirect with fileRefs).
      Add a `case "COMMAND_BATCH":` in `onExecute`. Build an `SPFI` against the DMS site
      (`${TENANT_DOMAIN}/${SITE_REDIRECT}`). For each selected row: if it is a PDF file
      (`FileLeafRef` ends `.pdf`), collect its `FileRef`; if it is a folder, enumerate PDFs inside it with
      PnPjs (e.g. `sp.web.getFolderByServerRelativePath(folderRef).files()` and/or a filtered items query),
      recursing or one level as needed, and collect those `FileRef`s. Deduplicate the list. If 0 PDFs after
      expansion: `Dialog.alert` explaining none were found and stop. Otherwise (including exactly 1 PDF —
      proceed as a 1-document batch; add a code comment noting this), redirect to
      `${TENANT_DOMAIN}/${SITES_ESIGN}/SitePages/${PAGE_INITIATE_ESIGN}?fileRefs=${encodeURIComponent(JSON.stringify(fileRefs))}&source=${encodeURIComponent(webUrlNow)}`.
      Do not change the COMMAND_1 redirect.
      Files: src/extensions/initiateEsign/InitiateEsignCommandSet.ts
      Verify: `npm run build` succeeds. Manual: executing on a folder/multi-select navigates to the initiate
      page with a `fileRefs` JSON array; empty selection shows the alert.

- [ ] 6. Refactor single-flow save into `createProcessForDocument` (no behavior change to single flow).
      In `InitiateSignature.tsx`, extract the per-document creation logic currently inside
      `handleSaveWorkflow` (QR embed + overwrite, hash, build `ProcessSignature` payload, add to
      `LIST_PROCESS`, add to `LIST_ACTIVITY_LOG`) into the `createProcessForDocument(params)` helper with
      the signature documented above. Make `handleSaveWorkflow` call it with `batch` undefined, passing the
      current component state (approvers, placeholders, fileBlob, fileMetadata.fileRef). Keep the add
      payload identical when `batch` is undefined (Title, HashHex, FileRef0, RequestorId, Status,
      ProcessData, UserNotif, isDownloadable). Resolve `currentUser.Id` once and pass it in.
      Files: src/webparts/initiateSignature/components/InitiateSignature.tsx
      Verify: `npm run build` succeeds. Manual: the existing single-document initiate flow produces exactly
      one `Approvals Process` item with no `BatchId`/`BatchMode`, identical to before.

- [ ] 7. Add batch mode to the initiate page: detect `fileRefs`, per-document state, navigation, validation, batch submit.
      In `InitiateSignature.tsx`: read `fileRefs` from the query string; if present and parses to >=1
      entries, enter batch mode. Maintain per-document state (each document keeps its own approvers +
      placeholders; load each PDF as the user navigates or preload). Add simple navigation between the N
      documents (stepper/tabs/list) reusing the existing `PdfViewer` and `Sidebar` per current document.
      Validation before submit: every document must have >=1 approver and every approver in a document must
      have >=1 placeholder (same rule as single flow); block submit and indicate which document is
      incomplete. On submit: generate one `batchId` (`crypto.randomUUID()` with a UUID v4 fallback), then
      loop documents calling `createProcessForDocument({... batch: { batchId, batchMode: "AllOrNothing" }})`.
      If one document fails mid-loop, surface a clear error and stop; do NOT delete already-created items
      (document this out-of-scope rollback limitation in a comment). When `fileRefs` is absent, the page
      behaves exactly as today.
      Files: src/webparts/initiateSignature/components/InitiateSignature.tsx,
      src/webparts/initiateSignature/components/Sidebar.tsx (only if navigation/validation UI needs props;
      keep single-flow props intact), src/webparts/initiateSignature/components/IInitiateSignatureProps.ts
      (only if a new prop is genuinely required)
      Verify: `npm run build` succeeds. Manual: opening the page with a `fileRefs` array lets the user
      configure each document separately and, on submit, creates N `Approvals Process` items that all share
      one `BatchId` and `BatchMode = AllOrNothing`; an incomplete document blocks submit with a clear message.

- [ ] 8. Add all-or-nothing sibling cancellation on reject in the approval form.
      In `InitiateEsignForm.tsx` `_submitReject`, AFTER the existing reject update (current task
      `Status="Rejected"` and process set to Rejected) and the existing activity log, add a defensive
      batch-cancel block wrapped in try/catch. Read the current process's `BatchId` and `BatchMode` from
      `LIST_PROCESS` via `getById(this.state.processId).select("BatchId","BatchMode")()` (guard for
      missing fields). If `BatchMode === "AllOrNothing"` and `BatchId` is non-empty: query sibling process
      items with `filter(\`BatchId eq '${batchId.replace(/'/g, "''")}' and Status eq 'Processing'\`)`
      excluding the current `processId`. For each sibling: update `Status="Canceled"` and
      `Notes="System Canceled: Batch rejected (sibling document <rejected process Title> was rejected)"`
      (exact `System Canceled:` prefix so `getProcessCancelInfo`/`getCanceledDialogContent` detection keeps
      working), then cancel that sibling's pending Tasks (query `LIST_TASKS` where `ProcessID/Id eq <siblingId>`
      and `Status eq 'Pending'`, set each to `Status="Canceled"` with a `Comments` note) — mirror the
      existing `_cancelPendingTasks` / `_cancelProcess` patterns. Never let a sibling-cancel failure prevent
      the reject from completing or the normal reject-success dialog from showing. Do NOT touch siblings
      that are `Approved`, `Rejected`, or `Canceled`. If `BatchMode` is None/empty/missing, behave exactly
      as today (no sibling lookups). You may add a small private helper
      `_cancelBatchSiblings(sp, batchId, rejectedProcessId, rejectedProcessTitle)` on the class.
      Files: src/extensions/initiateEsignForm/components/InitiateEsignForm.tsx
      Verify: `npm run build` succeeds. Manual: rejecting one document of an `AllOrNothing` batch flips its
      `Processing` siblings to `Canceled` (with the `System Canceled:` note) and cancels their pending
      tasks, while `Approved` siblings stay `Approved`; a non-batch reject is unchanged.

- [ ] 9. Final full build and lint gate.
      Run the production-style build to confirm no type/lint regressions across all edited files.
      Files: (none)
      Verify: `npm run build` exits 0 with no TypeScript or ESLint errors.

## Notes / assumptions / non-goals

- Owner creates `BatchId`/`BatchMode` columns manually; no PnP/CSOM provisioning from code.
- The existing Power Automate flow triggers on **item-created** in `LIST_PROCESS`; the batch loop adds one
  item per document, so per-document initiate emails fire automatically. No email-sending code is added
  here. Batch cancel/summary emails are handled entirely in Power Automate.
- Folder expansion depth: expand folders to collect PDFs; if nested subfolders must be included, recurse.
  Assumption: one-level expansion is sufficient for the Finance library layout; revisit if QA shows nested
  folders are in scope.
- Distributed rollback of already-created batch items is explicitly out of scope; on mid-loop failure the
  code stops and reports, leaving created items intact (documented in code).
