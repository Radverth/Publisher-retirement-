# Publisher retirement

Tooling for retiring Microsoft Publisher files from a Microsoft 365 tenant.

## [SPO-PubConverter](SPO-PubConverter/)

PowerShell tool that finds every `.pub` file across SharePoint Online, converts
the selected ones to PDF, and uploads the PDFs back beside the originals — all
from one numbered console menu.

```powershell
cd SPO-PubConverter
.\Start-Menu.ps1
```

See [SPO-PubConverter/README.md](SPO-PubConverter/README.md) for prerequisites,
the permission model and its blast radius, the CSV schema, and why the
conversion step has to run under Windows PowerShell 5.1 on a host with
Publisher installed.

Offline checks (no tenant required):

```powershell
.\SPO-PubConverter\tests\Run-Tests.ps1
```
