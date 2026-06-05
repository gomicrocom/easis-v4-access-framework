param(
    [Parameter(Mandatory = $true)]
    [string]$SourceRepoPath,

    [Parameter(Mandatory = $true)]
    [string]$TargetAccdbPath,

    [switch]$BackupBeforeImport,
    [switch]$DryRun,
    [switch]$AsJson,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$acModule = 5

function New-SyncEntry {
    param(
        [string]$Type,
        [string]$Name,
        [string]$Status,
        [string]$Message
    )

    [pscustomobject]@{
        type    = $Type
        name    = $Name
        status  = $Status
        message = $Message
    }
}

function Resolve-ExportRoot {
    param(
        [string]$BasePath
    )

    $resolvedBase = (Resolve-Path -LiteralPath $BasePath).Path
    $directModules = Join-Path $resolvedBase "modules"
    $directQueries = Join-Path $resolvedBase "queries"

    if ((Test-Path -LiteralPath $directModules) -and (Test-Path -LiteralPath $directQueries)) {
        return $resolvedBase
    }

    $repoExportRoot = Join-Path $resolvedBase "src\access\exported"
    if ((Test-Path -LiteralPath (Join-Path $repoExportRoot "modules")) -and (Test-Path -LiteralPath (Join-Path $repoExportRoot "queries"))) {
        return $repoExportRoot
    }

    throw "Could not resolve export root below '$BasePath'."
}

function Ensure-TestTargetAllowed {
    param(
        [string]$DatabasePath,
        [bool]$AllowForce,
        [bool]$SimulateOnly
    )

    $extension = [System.IO.Path]::GetExtension($DatabasePath)
    if ($extension -notin @(".accdb", ".mdb")) {
        throw "TargetAccdbPath must point to an .accdb or .mdb file."
    }

    if (-not $SimulateOnly -and -not (Test-Path -LiteralPath $DatabasePath)) {
        throw "TargetAccdbPath does not exist: $DatabasePath"
    }

    $leaf = [System.IO.Path]::GetFileName($DatabasePath)
    if ($leaf -ieq "easis.accdb" -and -not $AllowForce) {
        throw "TargetAccdbPath points to easis.accdb. Use a copied test ACCDB or pass -Force explicitly."
    }
}

function New-BackupCopy {
    param(
        [string]$DatabasePath
    )

    $directory = Split-Path -Parent $DatabasePath
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($DatabasePath)
    $extension = [System.IO.Path]::GetExtension($DatabasePath)
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = Join-Path $directory ($stem + ".backup_" + $timestamp + $extension)

    Copy-Item -LiteralPath $DatabasePath -Destination $backupPath -Force
    return $backupPath
}

function Write-AccessCompatibleModuleTempFile {
    param(
        [string]$ModulePath
    )

    $content = [System.IO.File]::ReadAllText($ModulePath, [System.Text.Encoding]::UTF8)
    $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName() + ".bas")
    $ansi = [System.Text.Encoding]::GetEncoding(1252)
    [System.IO.File]::WriteAllText($tempPath, $content, $ansi)
    return $tempPath
}

function Test-AccessModuleExists {
    param(
        [object]$AccessApplication,
        [string]$ModuleName
    )

    try {
        $null = $AccessApplication.CurrentProject.AllModules.Item($ModuleName)
        return $true
    }
    catch {
        return $false
    }
}

function Import-StandardModules {
    param(
        [object]$AccessApplication,
        [System.IO.FileInfo[]]$ModuleFiles,
        [bool]$SimulateOnly
    )

    $results = New-Object System.Collections.ArrayList
    $tempFiles = New-Object System.Collections.ArrayList

    try {
        foreach ($file in $ModuleFiles) {
            $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

            if ($SimulateOnly) {
                [void]$results.Add((New-SyncEntry "module" $moduleName "planned" "Module would be replaced from repo export."))
                continue
            }

            $tempPath = Write-AccessCompatibleModuleTempFile -ModulePath $file.FullName
            [void]$tempFiles.Add($tempPath)

            try {
                if (Test-AccessModuleExists -AccessApplication $AccessApplication -ModuleName $moduleName) {
                    $AccessApplication.DoCmd.DeleteObject($acModule, $moduleName)
                }

                $AccessApplication.LoadFromText($acModule, $moduleName, $tempPath)
                [void]$results.Add((New-SyncEntry "module" $moduleName "imported" "Module replaced successfully."))
            }
            catch {
                [void]$results.Add((New-SyncEntry "module" $moduleName "failed" $_.Exception.Message))
            }
        }
    }
    finally {
        foreach ($tempPath in $tempFiles) {
            try {
                if (Test-Path -LiteralPath $tempPath) {
                    Remove-Item -LiteralPath $tempPath -Force
                }
            }
            catch {}
        }
    }

    return @($results)
}

function Import-Queries {
    param(
        [object]$AccessApplication,
        [System.IO.FileInfo[]]$QueryFiles,
        [bool]$SimulateOnly
    )

    $results = New-Object System.Collections.ArrayList

    foreach ($file in $QueryFiles) {
        $queryName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

        if ($SimulateOnly) {
            [void]$results.Add((New-SyncEntry "query" $queryName "planned" "Query would be recreated from repo SQL."))
            continue
        }

        try {
            $sqlText = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)
            $db = $AccessApplication.CurrentDb()

            try {
                $db.QueryDefs.Delete($queryName)
            }
            catch {}

            $null = $db.CreateQueryDef($queryName, $sqlText)
            [void]$results.Add((New-SyncEntry "query" $queryName "imported" "Query replaced successfully."))
        }
        catch {
            [void]$results.Add((New-SyncEntry "query" $queryName "failed" $_.Exception.Message))
        }
    }

    return @($results)
}

function Write-HumanReadableSummary {
    param(
        [object]$Report
    )

    Write-Output "=== Access Test Sync ==="
    Write-Output ("Export root: " + $Report.export_root)
    Write-Output ("Target ACCDB: " + $Report.target_accdb)
    Write-Output ("DryRun: " + $Report.dry_run)
    Write-Output ("Backup path: " + $(if ([string]::IsNullOrWhiteSpace($Report.backup_path)) { "(none)" } else { $Report.backup_path }))
    Write-Output ""
    Write-Output ("Modules imported/planned: " + $Report.summary.modules_imported)
    Write-Output ("Queries imported/planned: " + $Report.summary.queries_imported)
    Write-Output ("Skipped: " + $Report.summary.skipped)
    Write-Output ("Failed: " + $Report.summary.failed)
    Write-Output ""

    if ($Report.warnings.Count -gt 0) {
        Write-Output "Warnings:"
        foreach ($warning in $Report.warnings) {
            Write-Output ("  - " + $warning)
        }
        Write-Output ""
    }

    Write-Output "Entries:"
    foreach ($entry in $Report.entries) {
        Write-Output ("  [{0}] {1} {2} | {3}" -f $entry.status, $entry.type, $entry.name, $entry.message)
    }
    Write-Output ""
    Write-Output "Next manual step:"
    Write-Output "  Access oeffnen -> Ziel-ACCDB laden -> Debug > Kompilieren -> Smoke-Test ausfuehren"
}

$resolvedTarget = if (Test-Path -LiteralPath $TargetAccdbPath) {
    (Resolve-Path -LiteralPath $TargetAccdbPath).Path
} else {
    [System.IO.Path]::GetFullPath($TargetAccdbPath)
}

Ensure-TestTargetAllowed -DatabasePath $resolvedTarget -AllowForce $Force.IsPresent -SimulateOnly $DryRun.IsPresent

$exportRoot = Resolve-ExportRoot -BasePath $SourceRepoPath
$moduleDir = Join-Path $exportRoot "modules"
$queryDir = Join-Path $exportRoot "queries"

$moduleFiles = @(Get-ChildItem -LiteralPath $moduleDir -Filter *.bas | Sort-Object Name)
$queryFiles = @(Get-ChildItem -LiteralPath $queryDir -Filter *.sql | Sort-Object Name)

$warnings = New-Object System.Collections.ArrayList
if ($moduleFiles.Count -eq 0) {
    [void]$warnings.Add("No .bas module exports were found.")
}
if ($queryFiles.Count -eq 0) {
    [void]$warnings.Add("No .sql query exports were found.")
}

$backupPath = ""
if ($BackupBeforeImport -and -not $DryRun) {
    $backupPath = New-BackupCopy -DatabasePath $resolvedTarget
}

$entries = New-Object System.Collections.ArrayList
$access = $null

try {
    if (-not $DryRun) {
        $access = New-Object -ComObject Access.Application
        $access.OpenCurrentDatabase($resolvedTarget)
    }

    foreach ($entry in (Import-StandardModules -AccessApplication $access -ModuleFiles $moduleFiles -SimulateOnly $DryRun.IsPresent)) {
        [void]$entries.Add($entry)
    }

    foreach ($entry in (Import-Queries -AccessApplication $access -QueryFiles $queryFiles -SimulateOnly $DryRun.IsPresent)) {
        [void]$entries.Add($entry)
    }
}
finally {
    if ($access -ne $null) {
        try { $access.CloseCurrentDatabase() } catch {}
        try { $access.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($access)
    }
}

$moduleImportedCount = @($entries | Where-Object { $_.type -eq "module" -and $_.status -in @("imported", "planned") }).Count
$queryImportedCount = @($entries | Where-Object { $_.type -eq "query" -and $_.status -in @("imported", "planned") }).Count
$skippedCount = @($entries | Where-Object { $_.status -eq "skipped" }).Count
$failedCount = @($entries | Where-Object { $_.status -eq "failed" }).Count

$report = [pscustomobject]@{
    export_root = $exportRoot
    target_accdb = $resolvedTarget
    dry_run = [bool]$DryRun
    backup_path = $backupPath
    warnings = @($warnings)
    entries = @($entries)
    summary = [pscustomobject]@{
        modules_imported = $moduleImportedCount
        queries_imported = $queryImportedCount
        skipped = $skippedCount
        failed = $failedCount
    }
}

if ($AsJson) {
    $report | ConvertTo-Json -Depth 8
} else {
    Write-HumanReadableSummary -Report $report
}
