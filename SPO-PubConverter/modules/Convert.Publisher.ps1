<#
.SYNOPSIS
	Converts Microsoft Publisher .pub files to PDF format.

.DESCRIPTION
	The COM/Interop half of the SharePoint Publisher File Converter, adapted
	from Tom's standalone Convert-PubFileToPDF.ps1. It keeps that script's
	proven mechanics unchanged:

	  * Add-Type -AssemblyName Office / Microsoft.Office.Interop.Publisher
	  * New-Object -ComObject Publisher.Application
	  * $doc.ExportAsFixedFormat([Microsoft.Office.Interop.Publisher.PbFixedFormatType]::pbFixedFormatTypePDF, $pdfFilePath)
	  * one try/finally around the whole batch, with $app.Quit() in the finally
	  * per-file try/catch and success/failure counters

	What changed for the pipeline:
	  * A -JobFile mode: the batch is supplied as JSON by Convert.psm1 (source
	    path, destination PDF path, and the SiteUrl/LibraryName/FolderPath and
	    UniqueId context), instead of being discovered from a folder filter.
	  * An existing PDF is no longer an error. -ExistingPdfAction decides:
	    Skip, Overwrite, or Version (write Name (2).pdf and so on).
	  * Results are appended to -ResultFile as one JSON object per line, as
	    each file finishes, so a hung or killed run still reports everything
	    that completed. Convert.psm1 feeds those into the CSV's Status/Notes.
	  * Each document is released with ReleaseComObject after Close so a long
	    batch cannot accumulate orphaned MSPUB processes.

	The original -Filter / -Recurse parameters still work for standalone use.

	THIS SCRIPT MUST RUN UNDER WINDOWS POWERSHELL 5.1 on a host with Microsoft
	Publisher installed. Early-bound Interop does not load reliably under
	PowerShell 7's .NET Core runtime - see the README.

.PARAMETER JobFile
	Path to a JSON array of jobs from Convert.psm1. Each element needs
	SourcePath and PdfPath; UniqueId, SiteUrl, LibraryName, FolderPath and
	FileName are carried through to the result so rows can be matched back to
	the CSV.

.PARAMETER ResultFile
	Path the per-file results are appended to, one JSON object per line.

.PARAMETER Filter
	Standalone mode: a file name or wildcard pattern, e.g. "*.pub".

.PARAMETER Recurse
	Standalone mode: search subdirectories too.

.PARAMETER ExistingPdfAction
	What to do when the destination PDF already exists: Skip, Overwrite or
	Version. Defaults to Version.

.PARAMETER LogFile
	Optional path to append plain-text log lines to.

.PARAMETER TestOnly
	Probe for Microsoft Publisher and exit: 0 if Publisher could be started,
	1 if not. Used by the main menu's prerequisite check.

.EXAMPLE
	Convert.Publisher.ps1 -Filter "*.pub" -Recurse
	Standalone: converts every Publisher file below the current directory.

.EXAMPLE
	powershell.exe -File Convert.Publisher.ps1 -JobFile jobs.json -ResultFile results.jsonl -ExistingPdfAction Overwrite
	Pipeline: converts the batch Convert.psm1 prepared.
#>
[CmdletBinding(DefaultParameterSetName = 'Pipeline')]
param
(
	[Parameter(ParameterSetName = 'Pipeline')]
	[ValidateNotNullOrEmpty()]
	[string]
	$JobFile,

	[Parameter(ParameterSetName = 'Pipeline')]
	[string]
	$ResultFile,

	[Parameter(ParameterSetName = 'Standalone')]
	[ValidateNotNullOrEmpty()]
	[string]
	$Filter,

	[Parameter(ParameterSetName = 'Standalone')]
	[switch]
	$Recurse,

	[ValidateSet('Skip', 'Overwrite', 'Version')]
	[string]
	$ExistingPdfAction = 'Version',

	[string]
	$LogFile,

	[Parameter(ParameterSetName = 'Test')]
	[switch]
	$TestOnly
)

$ErrorActionPreference = 'Continue';

function Write-ConversionLog {
	param
	(
		[string] $Message,
		[ValidateSet('Info', 'Success', 'Warn', 'Error')]
		[string] $Level = 'Info'
	)

	$line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level.ToUpper().PadRight(7), $Message;

	if ($LogFile) {
		try { Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue; } catch { }
	}

	if ($Level -eq 'Error') { Write-Error $Message; }
	else { Write-Output $line; }
}

function Write-ConversionResult {
	<#
	.SYNOPSIS
		Appends one result object to the JSONL result file.
	#>
	param
	(
		[Parameter(Mandatory = $true)] $Job,
		[Parameter(Mandatory = $true)]
		[ValidateSet('Converted', 'Skipped', 'Failed')]
		[string] $Status,
		[string] $PdfPath,
		[string] $Message
	)

	if (-not $ResultFile) { return; }

	$uniqueId = '';
	if ($Job.PSObject.Properties['UniqueId']) { $uniqueId = [string] $Job.UniqueId; }

	$sourcePath = '';
	if ($Job.PSObject.Properties['SourcePath']) { $sourcePath = [string] $Job.SourcePath; }

	$result = [pscustomobject] @{
		UniqueId   = $uniqueId
		SourcePath = $sourcePath
		PdfPath    = $PdfPath
		Status     = $Status
		Message    = ($Message -replace '[\r\n]+', ' ')
		Completed  = (Get-Date).ToString('s')
	};

	try {
		Add-Content -LiteralPath $ResultFile -Value ($result | ConvertTo-Json -Depth 4 -Compress) -Encoding UTF8 -ErrorAction Stop;
	} catch {
		Write-ConversionLog -Level Warn -Message ("Could not append to the result file: {0}" -f $_);
	}
}

function Resolve-PdfDestination {
	<#
	.SYNOPSIS
		Applies the Skip / Overwrite / Version rule to an existing PDF.

	.DESCRIPTION
		Replaces the original script's "PDF file already exists" error, which
		made re-running a partially failed batch impossible.

		Returns a hashtable: Action = Convert | Skip, Path = the PDF to write.
	#>
	param
	(
		[Parameter(Mandatory = $true)] [string] $PdfPath,
		[Parameter(Mandatory = $true)] [string] $Action
	)

	if (-not (Test-Path -LiteralPath $PdfPath)) {
		return @{ Action = 'Convert'; Path = $PdfPath; Message = '' };
	}

	switch ($Action) {
		'Skip' {
			return @{ Action = 'Skip'; Path = $PdfPath; Message = 'PDF already exists - skipped per the Skip setting.' };
		}
		'Overwrite' {
			try {
				Remove-Item -LiteralPath $PdfPath -Force -ErrorAction Stop;
				return @{ Action = 'Convert'; Path = $PdfPath; Message = 'Existing PDF overwritten.' };
			} catch {
				return @{ Action = 'Skip'; Path = $PdfPath; Message = ("Could not overwrite the existing PDF: {0}" -f $_) };
			}
		}
		default {
			$directory = [System.IO.Path]::GetDirectoryName($PdfPath);
			$baseName  = [System.IO.Path]::GetFileNameWithoutExtension($PdfPath);

			for ($version = 2; $version -le 999; $version++) {
				$candidate = Join-Path $directory ("{0} ({1}).pdf" -f $baseName, $version);
				if (-not (Test-Path -LiteralPath $candidate)) {
					return @{ Action = 'Convert'; Path = $candidate; Message = ("Existing PDF kept - writing version {0}." -f $version) };
				}
			}

			return @{ Action = 'Skip'; Path = $PdfPath; Message = 'Too many existing versions of this PDF (999).' };
		}
	}
}

function Get-JobList {
	<#
	.SYNOPSIS
		Builds the batch: from the JSON job file, or from -Filter standalone.
	#>
	param()

	if ($JobFile) {
		if (-not (Test-Path -LiteralPath $JobFile)) {
			Write-ConversionLog -Level Error -Message ("Job file not found: {0}" -f $JobFile);
			return @();
		}

		try {
			$content = Get-Content -LiteralPath $JobFile -Raw -Encoding UTF8;
			if ([string]::IsNullOrWhiteSpace($content)) { return @(); }
			return @($content | ConvertFrom-Json);
		} catch {
			Write-ConversionLog -Level Error -Message ("Could not read the job file: {0}" -f $_);
			return @();
		}
	}

	# --- standalone mode, as in the original script ---
	if (-not ($Filter -like "*.pub")) {
		Write-ConversionLog -Level Error -Message "The filter must specify .pub files (e.g. '*.pub' or 'file.pub').";
		return @();
	}

	$files = Get-ChildItem $Filter -File -Recurse:$Recurse;
	if (-not $files) {
		Write-ConversionLog -Level Error -Message ("No Publisher files found for the filter: {0}" -f $Filter);
		return @();
	}

	$jobs = @();
	foreach ($file in $files) {
		if ($file.Extension -ne '.pub') { continue; }
		$jobs += [pscustomobject] @{
			UniqueId    = $file.FullName
			SourcePath  = $file.FullName
			PdfPath     = [System.IO.Path]::ChangeExtension($file.FullName, '.pdf')
			FileName    = $file.Name
			SiteUrl     = ''
			LibraryName = ''
			FolderPath  = ''
		};
	}

	return $jobs;
}

# --------------------------------------------------------------------------
# Prerequisite: Publisher must be installed on this host.
# --------------------------------------------------------------------------
if ($TestOnly) {
	try {
		Add-Type -AssemblyName Office -ErrorAction Stop;
		Add-Type -AssemblyName Microsoft.Office.Interop.Publisher -ErrorAction Stop;
		$probe = New-Object -ComObject Publisher.Application -ErrorAction Stop;
		try { $probe.Quit(); } catch { }
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject($probe) | Out-Null;
		Write-Output 'Publisher OK';
		exit 0;
	} catch {
		Write-Output ("Publisher not available: {0}" -f $_.Exception.Message);
		exit 1;
	}
}

if ($PSCmdlet.ParameterSetName -eq 'Standalone' -and -not $PSBoundParameters.ContainsKey('Filter')) {
	Write-Error "The -Filter parameter is required in standalone mode.";
	exit 1;
}

if (-not $JobFile -and -not $Filter) {
	Write-Error "Supply either -JobFile (pipeline mode) or -Filter (standalone mode).";
	exit 1;
}

$app = $null;

try {
	$jobs = Get-JobList;
	if (-not $jobs -or @($jobs).Count -eq 0) {
		Write-ConversionLog -Level Warn -Message 'Nothing to convert.';
		exit 1;
	}

	Write-ConversionLog -Message ("Running... {0} file(s) queued, existing-PDF action '{1}'." -f @($jobs).Count, $ExistingPdfAction);

	Add-Type -AssemblyName Office;
	Add-Type -AssemblyName Microsoft.Office.Interop.Publisher;
	try {
		$app = New-Object -ComObject Publisher.Application;
	} catch {
		Write-ConversionLog -Level Error -Message "Microsoft Publisher is not installed or accessible.";
		exit 1;
	}

	$successCount = 0;
	$failCount    = 0;
	$skipCount    = 0;

	foreach ($job in $jobs) {
		$fileFullName = [string] $job.SourcePath;

		if ([string]::IsNullOrWhiteSpace($fileFullName) -or -not (Test-Path -LiteralPath $fileFullName)) {
			$failCount++;
			Write-ConversionLog -Level Error -Message ("Source file not found: {0}" -f $fileFullName);
			Write-ConversionResult -Job $job -Status Failed -Message 'Source file not found.';
			continue;
		}

		$pdfFilePath = [string] $job.PdfPath;
		if ([string]::IsNullOrWhiteSpace($pdfFilePath)) {
			$pdfFilePath = [System.IO.Path]::ChangeExtension($fileFullName, '.pdf');
		}

		$destinationFolder = [System.IO.Path]::GetDirectoryName($pdfFilePath);
		if ($destinationFolder -and -not (Test-Path -LiteralPath $destinationFolder)) {
			New-Item -Path $destinationFolder -ItemType Directory -Force | Out-Null;
		}

		$destination = Resolve-PdfDestination -PdfPath $pdfFilePath -Action $ExistingPdfAction;
		if ($destination.Action -eq 'Skip') {
			$skipCount++;
			Write-ConversionLog -Level Warn -Message ("Skipped {0}: {1}" -f $fileFullName, $destination.Message);
			Write-ConversionResult -Job $job -Status Skipped -PdfPath $destination.Path -Message $destination.Message;
			continue;
		}

		$pdfFilePath = $destination.Path;
		if ($destination.Message) { Write-ConversionLog -Message $destination.Message; }

		# Open the file
		$doc = $null;
		try {
			$doc = $app.Open($fileFullName);
		} catch {
			$failCount++;
			Write-ConversionLog -Level Error -Message ("Error opening file: {0} {1}" -f $fileFullName, $_);
			Write-ConversionResult -Job $job -Status Failed -Message ("Error opening file: {0}" -f $_);
			continue;
		}

		if (-not($doc)) {
			$failCount++;
			Write-ConversionLog -Level Error -Message ("Failed to open file: {0}" -f $fileFullName);
			Write-ConversionResult -Job $job -Status Failed -Message 'Failed to open file (Publisher returned no document).';
			continue;
		}

		try {
			# Export file as PDF
			$doc.ExportAsFixedFormat([Microsoft.Office.Interop.Publisher.PbFixedFormatType]::pbFixedFormatTypePDF, $pdfFilePath);
			if (Test-Path $pdfFilePath) {
				Write-ConversionLog -Level Success -Message ("Exported to {0}." -f $pdfFilePath);
				Write-ConversionResult -Job $job -Status Converted -PdfPath $pdfFilePath -Message '';
				$successCount++;
			} else {
				$failCount++;
				Write-ConversionLog -Level Error -Message ("Failed to export file: {0}" -f $fileFullName);
				Write-ConversionResult -Job $job -Status Failed -Message 'Export reported success but no PDF was written.';
			}
		} catch {
			$failCount++;
			Write-ConversionLog -Level Error -Message ("Error during export: {0}" -f $_);
			Write-ConversionResult -Job $job -Status Failed -Message ("Error during export: {0}" -f $_);
		}

		# Close and release the document so MSPUB does not accumulate handles
		# across a long batch.
		try { $doc.Close(); } catch { Write-ConversionLog -Level Warn -Message ("Error closing document: {0}" -f $_); }
		try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null; } catch { }
		$doc = $null;
	}

	#Log output
	Write-ConversionLog -Message ("Converted {0} files with {1} errors and {2} skipped." -f $successCount, $failCount, $skipCount);

	if ($failCount -gt 0 -and $successCount -eq 0) { exit 2; }
	exit 0;
} catch {
	Write-ConversionLog -Level Error -Message ("{0}" -f $_);
	exit 3;
} finally {
	if ($app) {
		#Quit Publisher
		try { $app.Quit(); } catch { }
		try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null; } catch { }
		[System.GC]::Collect();
		[System.GC]::WaitForPendingFinalizers();
	}
}
