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
| 1 | `1` → *Create a new app registration* | Interactive admin sign-in, app + service principal created, permissions consented. With PnP installed, the one-step path does the certificate and consent here too — skip step 2 |
| 2 | `2` | Self-signed certificate created and its public key uploaded to the app |
| 3 | `3` | App-only sign-in tested against Graph **and** SharePoint |
| 4 | `4` | Tenant (or scoped) crawl for `.pub` files, CSV written automatically |
| 5 | `6` | Load that CSV and narrow it down if you want a test batch first |
| 6 | `7` then `8` | Download, then convert to PDF |
| 7 | `9` | Confirm, then upload the PDFs back to SharePoint |

`10` runs 4 → 9 in one go after a single typed confirmation. `0` exits from any
menu level; anything unrecognised simply re-prompts.

## The menu

Everything is driven from one numbered menu that redisplays after each action,
so several phases can be run back to back. It shows where the run currently
stands, and tells you which option to pick next:

```
==============================================================================
   SharePoint Publisher File Converter                                  v1.0
==============================================================================
  Tenant         : contoso.onmicrosoft.com
  Tenant access  : ready - certificate expires 2028-09-11
  Site discovery : tenant admin list - finds every site (test with option 3)
  File list      : PublisherFileInventory_2026-09-11_1030.csv (214 files)
                   12 to download  |  30 to convert  |  170 to upload
                   2 failed
  Working on     : every row in the file list
  Settings       : existing PDF: Version  |  name clashes: Version
                   originals kept
------------------------------------------------------------------------------
  NEXT: option 7 - download the 12 file(s) not yet on this machine
------------------------------------------------------------------------------

  SETUP - do these once, in order
    1) Generate or connect Azure AD App Registration          done
    2) Generate & upload authentication certificate           done
    3) Test connection to Microsoft Graph / SharePoint        verify setup

  DISCOVERY - find the Publisher files
    4) Scan tenant for Publisher (.pub) files
    5) Export / re-export scan results to CSV                 214 rows

  CONVERSION - fetch the files and make the PDFs
    6) Load a CSV and select files to process                 change files
    7) Download selected files                                12 to download
    8) Convert downloaded files to PDF                        no Publisher

  PUBLISH - put the PDFs back in SharePoint
    9) Upload converted PDFs to original SharePoint location  170 ready

  UTILITIES
   10) Run full pipeline (4 -> 9) unattended                  no Publisher
   11) View recent log                                        run history
   12) Open working folder                                    files on disk
   13) Change conversion & upload settings                    rules, folders

    0) Exit
------------------------------------------------------------------------------
 Select an option:
```

Reading it:

- **The state block** says what is set up, which file list is loaded, how many
  files are at each stage, and whether the tool is working on everything or a
  subset you narrowed at option 6. Anything dangerous is spelled out here —
  `ORIGINALS DELETED after upload` appears the moment that setting is armed.
- **NEXT** names the single option to pick next, worked out from what is
  actually outstanding. An expired certificate outranks everything else.
- **The second column** is per-option status: `done`, `12 to download`,
  `no Publisher`, `needs setup`. Steps that cannot run yet are dimmed and say
  why rather than failing after you pick them.
- **0 always goes back or exits**, at every level. Typing `q`, `back` or `x`
  is treated the same way, a stray `)` or spaces are forgiven, and anything
  unrecognised re-prompts with the valid choices rather than erroring out.

Sub-menus follow the same convention — numbered, described, `0` last. The
status filter, for example, lists what each stage means with a live count, so
"retry just the failures" is one keystroke rather than a guess.

## Prerequisites

| Requirement | Needed for | Notes |
|-------------|-----------|-------|
| Windows PowerShell 5.1 **or** PowerShell 7 | the menu and all Graph work | The menu itself runs on either |
| `Microsoft.Graph.Authentication` | everything | Every call goes through `Invoke-MgGraphRequest`; the tool tells you how to install it if it is missing |
| `PnP.PowerShell` (PowerShell 7.4.6+) | one-step setup and complete site enumeration | Optional but recommended — `Install-Module PnP.PowerShell -Scope CurrentUser` |
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

## Graph for the pipeline, PnP for what Graph cannot automate

The brief asked for one approach with the trade-off documented. This build uses
a **bespoke single-tenant app registration** — one principal the customer owns,
named for this tool, revocable on its own — and drives it two ways:

| Work | Library | Why |
|------|---------|-----|
| Crawl, download, upload | Microsoft Graph SDK | Drive-item APIs are first-class in Graph, and it needs only `Microsoft.Graph.Authentication` |
| Tenant site enumeration | **PnP PowerShell** | `Get-PnPTenantSite` reads the SharePoint tenant admin list — the only complete one |
| App registration setup | **PnP PowerShell** (optional) | `Register-PnPEntraIDApp` does app + certificate + consent in one step |

An earlier draft of this README justified avoiding PnP on the grounds that it
would mean depending on the shared PnP Management Shell multi-tenant app. That
was wrong: PnP connects perfectly well with your own app registration and
certificate (`Connect-PnPOnline -ClientId <your app> -Tenant <x> -Thumbprint <y>`),
which is exactly what this tool does. No Management Shell app is registered or
used.

PnP stays **optional**. Where it is absent the tool runs end to end on Graph
alone, with the enumeration caveats below. What PnP genuinely costs:

- Current PnP needs **PowerShell 7.4.6+**; Windows PowerShell 5.1 is only
  supported by PnP 1.12.0, which is unmaintained. The menu still runs under
  5.1 without PnP, and the COM conversion step always runs under 5.1 either
  way.
- The tenant admin site list requires **SharePoint `Sites.FullControl.All`** —
  see below.

## Permissions and blast radius

Setup offers three scopes. All permissions are **application** permissions —
there is no user context at run time.

| Scope | Permissions | Site discovery |
|-------|-------------|----------------|
| **`TenantAdmin`** (default) | the `AllSites` set **plus** SharePoint `Sites.FullControl.All` | Fully automatic and complete — reads the tenant admin site list |
| `AllSites` | `Sites.Read.All`, `Sites.ReadWrite.All`, `Files.ReadWrite.All`, `Directory.Read.All` (Graph) | Falls back to the search index, which can miss sites |
| `SitesSelected` | `Sites.Selected`, `Directory.Read.All` (Graph) | Only the sites explicitly granted to the app |

What each Graph permission is for:

| Permission | Purpose |
|------------|---------|
| `Sites.Read.All` | Enumerate site collections and crawl document libraries |
| `Sites.ReadWrite.All` | Upload converted PDFs back to the source library |
| `Files.ReadWrite.All` | Download source `.pub` files and write PDFs via drive items |
| `Directory.Read.All` | Resolve site/user metadata (optional) |

> **`Sites.FullControl.All` (SharePoint API) is the largest grant here** — full
> administrative control of every site collection in the tenant, beyond the
> tenant-wide read/write below. It is what makes site discovery complete and
> automatic, and it is required for that: `Sites.Read.All` is explicitly not
> enough to list tenant sites, and `Sites.Manage.All` is the documented floor.
> It cannot be combined with `Sites.Selected`. Choose it deliberately, and
> delete the app registration when the retirement project is finished.

App role IDs are resolved live from each resource service principal by
permission name rather than from hardcoded GUIDs, so consent does not silently
grant the wrong role if Microsoft ever changes one.

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

## Which sites get found

**Ownership and membership are irrelevant.** Discovery runs app-only: the tool
authenticates as the app registration with its certificate, with no user
context at all. `Sites.Read.All` as an *application* permission covers every
site collection in the tenant — private team sites, sites with broken
inheritance, sites the operator has never been a member of and could not open
in a browser. The operator's own account only matters during setup, to create
the app and grant consent. (In `Sites.Selected` mode the opposite applies: only
sites explicitly granted to the app are visible.)

**Enumeration, not access, is the real limit.** The app can read any site; the
question is whether the tool can *list* a site in order to go there. Three
routes are tried, best first, and the one actually used is logged with the site
count:

| Route | Source | Complete? | Needs |
|-------|--------|-----------|-------|
| **`Get-PnPTenantSite`** | SharePoint tenant admin | **yes** | PnP + `TenantAdmin` scope |
| `v1.0 /sites/getAllSites` | tenant store | yes — not available in every tenant | Graph |
| `beta /sites/getAllSites` | tenant store | yes | Graph |
| `v1.0 /sites?search=*` | **search index** | no — see below | Graph |

With the `TenantAdmin` scope and PnP installed, the first route is used and
nothing is missed — no manual export, no search-index gaps. The PnP route
returns URLs, which are then resolved to Graph sites and crawled by the same
code path as a hand-supplied site list.

If it falls through to site search, the scan prints a warning, because that
route can miss sites excluded from search indexing, sites created too recently
to be indexed, and Teams private-channel sites (which are separate site
collections).

**If you are not using the PnP route**, the guaranteed-complete fallback is
manual: SharePoint admin centre → Active sites →
**Export to CSV**, then menu `4` → *read the URLs from a text or CSV file*. That
list comes from the SharePoint tenant store rather than the search index, so
nothing is missing. The export loads unedited — its column is `URL`, and the
loader also accepts `SiteUrl`, `Site URL`, `Url`, `WebUrl` or a plain text file
of one URL per line with `#` comments.

**OneDrive.** Personal (`-my.sharepoint.com`) sites are excluded by default and
the scan reports how many it skipped. The scan-scope prompt offers to include
them — slower, and it reads every user's personal files, so only where the
Publisher retirement has to cover OneDrive too.

Subsites are pulled in separately for every site found, on all routes, because
site search does not reliably return them.

A site can still be out of reach if a Restricted Access Control policy or a
Graph application access policy blocks the app; those failures are logged
per-site and do not stop the crawl.

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
* The password is stored under a name derived from the `.pfx` file itself
  (`PfxPassword_<pfx file name>`), recorded in config as
  `CertificateSecretName`, and looked up the same way — so the writer and the
  reader can never disagree about the key. Older spellings are still read, so a
  certificate created by an earlier build keeps working.
* The `.certs` and `.secrets` folders belong together: copying the tool to
  another machine without `.secrets` leaves a `.pfx` that cannot be opened.
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

Menu option `13` sets both rules — plus the working folder and the site
enumeration method (`Auto` / `PnP` / `Graph`). Neither collision rule is
hardcoded.

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

## Watching a long crawl

A tenant-wide scan runs for a long time, and most of that time is spent inside
a single site paging through one large library. Progress is reported at three
levels so it is always clear the run is alive:

```
Scanning SharePoint for .pub files
  Site 15 of 45 | 312 .pub found | 6m 04s elapsed | about 12m 30s left
  https://contoso.sharepoint.com/sites/Marketing
    Library 2 of 6: Documents
      184 folder(s) read | 41,900 item(s) seen | 12 .pub found | 37 folder(s) queued
      /Campaigns/2024/Print (page 9)
```

* The **site bar** carries elapsed time, the running find count, and a rough
  estimate of time remaining (from the third site onwards — sites vary enormously
  in size, so treat it as an order of magnitude, not a promise).
* The **library bar** updates on every page of every folder, so a library with
  tens of thousands of items visibly ticks over rather than looking hung.
* **Throttling** — the usual reason a crawl seems to stop dead — shows its own
  countdown bar while it waits, instead of silence.
* Any site that takes over a minute gets a line in the log saying how long it
  took, so the slow ones can be identified after the run.

**Partial results are saved as it goes.** Every 10 sites the inventory so far is
written to `PublisherFileInventory_<timestamp>_partial.csv`, and the path is
logged. If a long crawl is interrupted, load that file with menu option `6`
rather than starting again.

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
    PnP.psm1                  optional PnP: setup + enumeration  (addition)
  tests/Run-Tests.ps1         offline checks, no tenant needed   (addition)
  working/                    downloaded originals, /converted, /inventory
  logs/
```

`Config.psm1`, `Graph.psm1` and `PnP.psm1` are additions to the module layout in
the brief:
config handling and the throttling-aware request wrapper are used by every
phase, and putting them in one place each is what stops the retry logic being
copy-pasted five times. `PnP.psm1` keeps every PnP dependency behind one
boundary, so the tool still runs with PnP absent. `Discovery.psm1` owns the CSV
schema, so `Convert.psm1`
and `Upload.psm1` import it rather than defining the format again.

## Tests

```powershell
.\tests\Run-Tests.ps1
```

189 offline checks: every file parses and every module imports, the CSV schema
matches the brief exactly, local paths mirror SharePoint without collisions,
filters and status counts behave, config round-trips without persisting
secrets, certificate expiry warns at the right thresholds, the
skip/overwrite/version rule does what it says, scope files load (including the
SharePoint admin centre export unedited), the permission scopes stay separated
and app role ids resolve live, Graph paging hands each page to its progress callback,
a full discovery crawl runs against stubbed Graph responses (site resolution, library filtering, folder recursion, extension
matching, row construction), the menu renders correctly in every state
(right options, right order, correct next step, nothing wider than 80
columns), and the preserved parts of Tom's conversion script (the Interop enum, the COM pattern, `app.Quit()` in
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
| `Argument types do not match`, usually just after site resolution | Fixed. Some PowerShell builds miscompile `@( )` around a generic `List`, which is what the crawl used to build its site list. Update to the current version. |
| `The certificate data cannot be read with the provided password` | Fixed — the password was stored under one name and read under another. Update to the current version; your existing certificate and `.pfx` still work, nothing needs re-issuing. |
| Same error after updating | The `.secrets` folder is missing (not copied from the machine that ran setup, or deleted). Re-run setup option `2` for a fresh certificate. |
| Discovery says the tenant admin route needs the TenantAdmin scope | The app was registered tenant-wide without the SharePoint permission. Menu `1` → option `5` adds it to the app you already have, keeping the same App ID and certificate. |
| PnP admin sign-in fails with 403 | The app lacks SharePoint `Sites.FullControl.All`, or the grant has not replicated yet (it can take several minutes). Re-test with option `3`. |
| Setup offers no one-step path | PnP is not installed, or the host is on PowerShell 5.1. Both are fine — the Graph step-by-step path does the same job. |
| Scan found fewer sites than expected | It fell through to the search-index route — the log says which route was used. Re-run scoped to the admin centre's Active sites export. |
| A known site is missing from the CSV | Same cause, or the site is blocked by a Restricted Access Control / application access policy (logged per site). |
| Crawl is slow on a large tenant | Expected — it is per-site, per-library, per-folder. Scope the first run to a couple of sites using option `4` → *Specific sites*. |
