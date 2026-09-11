# SharePoint Publisher File Converter

Tenant-wide `.pub` discovery, PDF conversion and re-upload for Microsoft 365 /
SharePoint Online, driven from one numbered console menu.

Built to the Claude Code brief *SharePoint Publisher File Converter* (Affinity
IT), around Tom's existing `Convert-PubFileToPDF.ps1`, whose COM mechanics are
preserved intact in `modules/Convert.Publisher.ps1`.

---

## What it does

1. **Setup** — creates (or connects to) an Azure AD app registration, generates
   a certificate, uploads the public key, and grants admin consent. No manual
   portal work unless the signed-in account cannot grant consent.
2. **Discovery** — crawls every SharePoint site the app can see (or a scoped
   subset) for `.pub` files and exports a CSV inventory.
3. **Download & convert** — downloads the selected files, mirroring their
   SharePoint structure locally, and converts each to PDF through Microsoft
   Publisher COM automation.
4. **Upload** — puts each PDF back in the SharePoint folder its `.pub` came
   from.

The CSV is the single source of truth between phases. Each phase updates its
`Status` column, so a run can be interrupted and resumed without redoing
completed work.

## Quick start

```powershell
cd SPO-PubConverter
.\Start-Menu.ps1
```

Then, first time through:

| Step | Menu option | What happens |
|------|-------------|--------------|
| 1 | `1` → *Create a new app registration* | Interactive admin sign-in, app + service principal created, permissions consented |
| 2 | `2` | Self-signed certificate created and its public key uploaded to the app |
| 3 | `3` | App-only sign-in tested against Graph **and** SharePoint |
| 4 | `4` | Tenant (or scoped) crawl for `.pub` files, CSV written automatically |
| 5 | `6` | Load that CSV and narrow it down if you want a test batch first |
| 6 | `7` then `8` | Download, then convert to PDF |
| 7 | `9` | Confirm, then upload the PDFs back to SharePoint |

`10` runs 4 → 9 in one go after a single typed confirmation. `0` exits from any
menu level; anything unrecognised simply re-prompts.

## Prerequisites

| Requirement | Needed for | Notes |
|-------------|-----------|-------|
| Windows PowerShell 5.1 **or** PowerShell 7 | the menu and all Graph work | The menu itself runs on either |
| `Microsoft.Graph.Authentication` | everything | Every call goes through `Invoke-MgGraphRequest`; the tool tells you how to install it if it is missing |
| **Windows host with Microsoft Publisher installed** | the conversion step only | Hard prerequisite, checked at startup and before any conversion |
| Global Administrator / Application Administrator | setup options 1, 2 and the Sites.Selected grant | Only for setup — day-to-day runs use the certificate |
| OpenSSL | certificate generation on non-Windows hosts | Fallback path only |

### Where each phase can run

Discovery, download and upload are pure Graph calls and run anywhere PowerShell
does. **Conversion only runs on Windows with Publisher installed.** If your
jump box has no Publisher, split the work:

1. On any host: options `4` (scan), `6`, `7` (download).
2. Copy the working folder and the CSV to the Publisher host.
3. There: options `6` (load the same CSV), `8` (convert), `9` (upload).

The startup banner warns when Publisher is not detected, and options `8` and
`10` refuse to run rather than failing halfway through a batch.

## Why the conversion runs out of process, under PowerShell 5.1

There is no Graph or cloud API that converts `.pub` to PDF — Publisher has no
server-side conversion story at all. The only reliable route is COM automation
against a locally installed Publisher.

Early-bound Interop (`Add-Type -AssemblyName Microsoft.Office.Interop.Publisher`,
resolved from the GAC) does not load reliably under PowerShell 7's .NET Core
runtime. So `modules/Convert.psm1` does the Graph half in whichever PowerShell
you launched, then shells out to **Windows PowerShell 5.1** to run
`modules/Convert.Publisher.ps1` for the COM half.

The process boundary earns its keep twice over: it keeps Tom's working COM code
working unchanged, and a `.pub` that hangs Publisher (a password prompt on a
protected file, say) kills one child process rather than the whole run. Work is
sent in batches of 25; each file's result is appended to a JSONL file as it
finishes, so even a killed batch reports everything that completed, and the
file that hung is marked `Failed` with a reason.

## Microsoft Graph SDK, not PnP PowerShell

The brief asked for one approach with the trade-off documented. This build uses
a **bespoke single-tenant app registration driven by the Microsoft Graph
PowerShell SDK** for everything.

**For:** one principal the customer owns, named for this tool, revocable on its
own; explicit, auditable permissions that can be narrowed to `Sites.Selected`;
no dependency on the Microsoft-operated PnP Management Shell multi-tenant app
that other tooling in the tenant may also rely on — revoking that would break
those too.

**Against:** a little more setup code than
`Register-PnPManagementShellAccess`, and no access to CSOM-only SharePoint
features. Nothing this tool does (enumerate sites, read and write drive items)
needs CSOM.

## Permissions and blast radius

Default (`AllSites`) mode requests these **application** permissions:

| Permission | Purpose |
|------------|---------|
| `Sites.Read.All` | Enumerate site collections and crawl document libraries |
| `Sites.ReadWrite.All` | Upload converted PDFs back to the source library |
| `Files.ReadWrite.All` | Download source `.pub` files and write PDFs via drive items |
| `Directory.Read.All` | Resolve site/user metadata (optional) |

> **`Sites.ReadWrite.All` and `Files.ReadWrite.All` are tenant-wide write
> permissions.** They grant read and write to every SharePoint site and every
> OneDrive in the tenant, with no user context and no per-site restriction.
> Anything holding this app's certificate can read or modify any file in the
> tenant.

Tighter alternative, offered in the setup menu: **`Sites.Selected`**. The app
then has no access at all until an administrator grants it per site (menu
`1` → *Grant this app access to one specific site*). That is one extra grant
per site — the recommended posture for a scoped retirement project. Discovery
needs `read`; the full pipeline needs `write`.

Either way, delete the app registration when the project is done. It is a
project tool, not permanent infrastructure.

## Certificates and secrets

* Windows: `New-SelfSignedCertificate`, 2048-bit RSA/SHA-256, two-year default
  validity, private key left in `Cert:\CurrentUser\My`. Only the `.cer` public
  key is uploaded to the app registration.
* Non-Windows: OpenSSL fallback producing `.key` / `.cer` / `.pfx` under
  `.certs`. The generated `.pfx` password goes to SecretManagement if a vault is
  registered, otherwise to a DPAPI-protected file under `.secrets` — never to
  `config.json`, the CSV or the logs. (`Export-Clixml` only encrypts on
  Windows; register a SecretManagement vault if you use the OpenSSL path in
  anger.)
* `Save-PubConfig` actively refuses to persist any key whose name looks like a
  secret, so a later edit cannot start leaking one by accident.
* Expiry is re-checked **at every startup**, not just during setup, and warns
  inside 30 days. The menu header shows the expiry date at all times.

## The CSV inventory

Written to `working\inventory\PublisherFileInventory_<yyyy-MM-dd_HHmm>.csv`,
and the path is echoed to the console when the scan finishes.

| Column | Description |
|--------|-------------|
| `SiteUrl` | SharePoint site collection URL |
| `LibraryName` | Document library the file lives in |
| `FolderPath` | Library-relative folder path (`/` at the library root) |
| `FileName` | Original file name including `.pub` |
| `FileSizeKB` | Size at time of scan |
| `LastModified` | SharePoint last-modified timestamp |
| `ModifiedBy` | Last-modified-by UPN or display name |
| `UniqueId` | SharePoint file GUID — the join key between phases |
| `Status` | `Pending` / `Downloaded` / `Converted` / `Uploaded` / `Skipped` / `Failed` |
| `Notes` | Free text; error detail on failure |

Plus operational columns the later phases need in order to find their way back
to the file without re-crawling the tenant, and to resume after an
interruption: `SiteId`, `DriveId`, `ItemId`, `ParentItemId`, `LocalPath`,
`PdfPath`, `LastAction`.

Trimming the CSV down in Excel is an expected workflow. Missing optional
columns are repaired on load; a CSV with no `DriveId`/`ItemId` is rejected with
a message telling you to re-run discovery, because those rows cannot be
uploaded back.

## Existing files: skip, overwrite or version

Menu option `13` sets both rules; neither is hardcoded.

| Setting | Applies to | `Skip` | `Overwrite` | `Version` (default) |
|---------|-----------|--------|-------------|---------------------|
| Existing local PDF | conversion | leave it, mark the row `Skipped` | delete and re-convert | write `Name (2).pdf` |
| Name collision in SharePoint | upload | leave it, mark the row `Skipped` | replace the file | SharePoint renames to `Name 1.pdf` |

Re-running a batch after a partial failure is a normal flow, so an existing PDF
is never an error — that is the one behaviour of the original standalone script
that had to change.

## Write actions always need confirmation

* Scans, downloads and conversions run without prompting — they only read from
  SharePoint.
* Upload shows the count and destination and needs a typed `Y`.
* Deleting the source `.pub` after upload is **off by default**, has to be
  armed in the settings menu, and still asks you to type `DELETE` in full on
  every run. Deleted files go to the site recycle bin.
* The unattended pipeline (option `10`) asks you to type `RUN` once up front,
  which covers its upload step.

## Logging

Every run writes `logs\PubConverter_<phase>_<timestamp>.log` plus a PowerShell
transcript, with one summary line per file:

```
2026-09-11 10:31:02 [INFO   ] DOWNLOADED | Newsletter.pub | C:\...\originals\...\Newsletter.pub
2026-09-11 10:31:14 [INFO   ] CONVERTED  | Newsletter.pub | C:\...\converted\...\Newsletter.pdf
2026-09-11 10:33:41 [ERROR  ] FAILED     | Brochure.pub   | Error opening file: ...
```

Each phase ends with a summary — attempted / succeeded / failed / skipped, plus
where the CSV and log were written. Menu option `11` tails the most recent log
without leaving the console.

## Error handling

* Every Graph call goes through one wrapper with retry-and-backoff: HTTP 429
  honours `Retry-After`, transient 5xx backs off 2s → 4s → 8s → 16s, up to five
  attempts. Permanent 4xx failures are not retried.
* Per-file failures are logged, marked `Failed` in the CSV with a reason, and
  never stop the batch.
* A failure inside a menu option is caught, logged, and returns you to the
  menu — the tool does not drop the technician back to a bare prompt.
* Long local paths (250+ characters) are warned about before they bite.

## Layout

```
SPO-PubConverter/
  Start-Menu.ps1              entry point, renders the numbered menu
  config.json                 persisted tenant/app/cert config (git-ignored)
  config.sample.json          the shape of config.json, with dummy values
  modules/
    AppRegistration.psm1      Phase 0: create/detect app reg + certificate
    Discovery.psm1            Phase 1: tenant crawl, CSV export, CSV schema
    Convert.psm1              Phase 2 driver: download + invokes PS 5.1
    Convert.Publisher.ps1     PS 5.1 COM automation, called out-of-process
    Upload.psm1               Phase 3: upload + status update
    Logging.psm1              shared transcript/log helper
    Config.psm1               config.json + secret handling      (addition)
    Graph.psm1                Graph connect + throttling retry   (addition)
  tests/Run-Tests.ps1         offline checks, no tenant needed   (addition)
  working/                    downloaded originals, /converted, /inventory
  logs/
```

`Config.psm1` and `Graph.psm1` are additions to the module layout in the brief:
config handling and the throttling-aware request wrapper are used by every
phase, and putting them in one place each is what stops the retry logic being
copy-pasted five times. `Discovery.psm1` owns the CSV schema, so `Convert.psm1`
and `Upload.psm1` import it rather than defining the format again.

## Tests

```powershell
.\tests\Run-Tests.ps1
```

75 offline checks: every file parses and every module imports, the CSV schema
matches the brief exactly, local paths mirror SharePoint without collisions,
filters and status counts behave, config round-trips without persisting
secrets, certificate expiry warns at the right thresholds, the
skip/overwrite/version rule does what it says, and the preserved parts of Tom's
conversion script (the Interop enum, the COM pattern, `app.Quit()` in
`finally`) are still there. Nothing touches a tenant, so it is safe to run any
time — including on the Linux/macOS host you might be editing from.

## Known limits and future considerations

* **No Publisher, no conversion.** If this ever needs to run unattended on a
  server with no Publisher install, the realistic options are a licensed
  third-party conversion API/library, or a dedicated VM image with Publisher
  installed and this tool scheduled on it. Not built now, by design.
* **Password-protected `.pub` files** cannot be converted — Publisher wants the
  password interactively. They are caught by the batch timeout, marked `Failed`
  and skipped, and the rest of the batch continues.
* **Single tenant per run.** The pattern extends to a TenantDeck-style
  multi-tenant loop, but v1 is scoped to one tenant.
* **v1 converts `.pub` only**, does not edit Publisher content, and has no GUI
  beyond this console menu.

## Troubleshooting

| Symptom | Cause / fix |
|---------|-------------|
| `App-only sign-in failed` right after setup | Azure AD takes a minute to publish a new certificate. Wait, then retry option `3`. |
| Option `3` reads Graph but fails on SharePoint | Admin consent not granted, or `Sites.Selected` mode with no site grant yet. Menu `1` → option `3` or `4`. |
| Everything `Failed` with `HTTP 403` | The app has no access to that site. In `Sites.Selected` mode each site needs its own grant. |
| Conversion says Publisher is not available | Run the conversion phase on the Windows host with Publisher, per *Where each phase can run*. |
| A batch stalls then every file in it fails | One file hung Publisher. The batch timed out and was killed; re-run option `8` and it resumes from the files with no result. Check Task Manager for a stray `MSPUB.EXE`. |
| Crawl is slow on a large tenant | Expected — it is per-site, per-library, per-folder. Scope the first run to a couple of sites using option `4` → *Specific sites*. |
