param(
    [ValidateSet("ImportOnly", "Stage", "Promote")]
    [string]$Mode = "ImportOnly",

    [string]$ConfigPath = (Join-Path $PSScriptRoot "access-sync.config.json"),

    [string]$SourceRepoPath,
    [string]$TargetAccdbPath,
    [string]$ActiveAccdbPath,
    [string]$StagingAccdbPath,
    [string]$BackupFolder,

    [switch]$BackupBeforeImport,
    [switch]$DryRun,
    [switch]$AsJson,
    [switch]$Force,
    [switch]$DeleteStagingAfterPromote
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

function Resolve-PathIfPossible {
    param(
        [string]$PathText
    )

    if ([string]::IsNullOrWhiteSpace($PathText)) {
        return ""
    }

    if (Test-Path -LiteralPath $PathText) {
        return (Resolve-Path -LiteralPath $PathText).Path
    }

    return [System.IO.Path]::GetFullPath($PathText)
}

function Get-JsonConfigValue {
    param(
        [object]$ConfigObject,
        [string]$PropertyName
    )

    if ($null -eq $ConfigObject) {
        return ""
    }

    $property = $ConfigObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return ""
    }

    return [string]$property.Value
}

function Resolve-SyncConfiguration {
    param(
        [string]$ModeName,
        [string]$ConfigFilePath
    )

    $configObject = $null
    $resolvedConfigPath = ""

    if (-not [string]::IsNullOrWhiteSpace($ConfigFilePath) -and (Test-Path -LiteralPath $ConfigFilePath)) {
        $resolvedConfigPath = (Resolve-Path -LiteralPath $ConfigFilePath).Path
        $configObject = Get-Content -LiteralPath $resolvedConfigPath -Raw | ConvertFrom-Json
    }

    $resolvedSourceRepoPath = if (-not [string]::IsNullOrWhiteSpace($SourceRepoPath)) { $SourceRepoPath } else { Get-JsonConfigValue $configObject "SourceRepoPath" }
    $resolvedTargetAccdbPath = if (-not [string]::IsNullOrWhiteSpace($TargetAccdbPath)) { $TargetAccdbPath } else { Get-JsonConfigValue $configObject "TargetAccdbPath" }
    $resolvedActiveAccdbPath = if (-not [string]::IsNullOrWhiteSpace($ActiveAccdbPath)) { $ActiveAccdbPath } else { Get-JsonConfigValue $configObject "ActiveAccdbPath" }
    $resolvedStagingAccdbPath = if (-not [string]::IsNullOrWhiteSpace($StagingAccdbPath)) { $StagingAccdbPath } else { Get-JsonConfigValue $configObject "StagingAccdbPath" }
    $resolvedBackupFolder = if (-not [string]::IsNullOrWhiteSpace($BackupFolder)) { $BackupFolder } else { Get-JsonConfigValue $configObject "BackupFolder" }

    if ($ModeName -eq "ImportOnly" -and [string]::IsNullOrWhiteSpace($resolvedTargetAccdbPath)) {
        $resolvedTargetAccdbPath = $resolvedStagingAccdbPath
    }

    [pscustomobject]@{
        config_path        = $resolvedConfigPath
        source_repo_path   = Resolve-PathIfPossible $resolvedSourceRepoPath
        target_accdb_path  = Resolve-PathIfPossible $resolvedTargetAccdbPath
        active_accdb_path  = Resolve-PathIfPossible $resolvedActiveAccdbPath
        staging_accdb_path = Resolve-PathIfPossible $resolvedStagingAccdbPath
        backup_folder      = Resolve-PathIfPossible $resolvedBackupFolder
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

function Ensure-AccdbPathLooksValid {
    param(
        [string]$DatabasePath,
        [string]$ParameterName
    )

    $extension = [System.IO.Path]::GetExtension($DatabasePath)
    if ($extension -notin @(".accdb", ".mdb")) {
        throw "$ParameterName must point to an .accdb or .mdb file."
    }
}

function Ensure-TestTargetAllowed {
    param(
        [string]$DatabasePath,
        [bool]$AllowForce,
        [bool]$SimulateOnly
    )

    Ensure-AccdbPathLooksValid -DatabasePath $DatabasePath -ParameterName "TargetAccdbPath"

    if (-not $SimulateOnly -and -not (Test-Path -LiteralPath $DatabasePath)) {
        throw "TargetAccdbPath does not exist: $DatabasePath"
    }

    $leaf = [System.IO.Path]::GetFileName($DatabasePath)
    if ($leaf -ieq "easis.accdb" -and -not $AllowForce) {
        throw "TargetAccdbPath points to easis.accdb. Use a copied test ACCDB or pass -Force explicitly."
    }
}

function Test-DatabaseClosed {
    param(
        [string]$DatabasePath
    )

    if (-not (Test-Path -LiteralPath $DatabasePath)) {
        return $true
    }

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($DatabasePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $stream) {
            $stream.Close()
            $stream.Dispose()
        }
    }
}

function Ensure-DatabaseClosed {
    param(
        [string]$DatabasePath,
        [string]$Label
    )

    if (-not (Test-DatabaseClosed -DatabasePath $DatabasePath)) {
        throw "$Label is currently open or locked: $DatabasePath"
    }
}

function Ensure-BackupFolder {
    param(
        [string]$FolderPath,
        [bool]$SimulateOnly
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        throw "BackupFolder is required for this mode."
    }

    if ($SimulateOnly) {
        return
    }

    if (-not (Test-Path -LiteralPath $FolderPath)) {
        $null = New-Item -ItemType Directory -Path $FolderPath -Force
    }
}

function New-BackupCopy {
    param(
        [string]$DatabasePath,
        [string]$BackupFolderPath
    )

    Ensure-BackupFolder -FolderPath $BackupFolderPath -SimulateOnly:$false

    $stem = [System.IO.Path]::GetFileNameWithoutExtension($DatabasePath)
    $extension = [System.IO.Path]::GetExtension($DatabasePath)
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = Join-Path $BackupFolderPath ($stem + ".backup_" + $timestamp + $extension)

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

function Invoke-AccessImport {
    param(
        [string]$ExportRoot,
        [string]$DatabasePath,
        [bool]$SimulateOnly
    )

    $moduleDir = Join-Path $ExportRoot "modules"
    $queryDir = Join-Path $ExportRoot "queries"

    $moduleFiles = @(Get-ChildItem -LiteralPath $moduleDir -Filter *.bas | Sort-Object Name)
    $queryFiles = @(Get-ChildItem -LiteralPath $queryDir -Filter *.sql | Sort-Object Name)

    $warnings = New-Object System.Collections.ArrayList
    if ($moduleFiles.Count -eq 0) {
        [void]$warnings.Add("No .bas module exports were found.")
    }
    if ($queryFiles.Count -eq 0) {
        [void]$warnings.Add("No .sql query exports were found.")
    }

    $entries = New-Object System.Collections.ArrayList
    $access = $null

    try {
        if (-not $SimulateOnly) {
            $access = New-Object -ComObject Access.Application
            $access.OpenCurrentDatabase($DatabasePath)
        }

        foreach ($entry in (Import-StandardModules -AccessApplication $access -ModuleFiles $moduleFiles -SimulateOnly $SimulateOnly)) {
            [void]$entries.Add($entry)
        }

        foreach ($entry in (Import-Queries -AccessApplication $access -QueryFiles $queryFiles -SimulateOnly $SimulateOnly)) {
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

    return [pscustomobject]@{
        warnings = @($warnings)
        entries   = @($entries)
    }
}

function New-StageNextSteps {
    param(
        [string]$StagingPath
    )

    @(
        "Staging-ACCDB oeffnen: $StagingPath"
        "In Access: Debug > Kompilieren"
        "Smoke-Test im Staging durchfuehren"
        "Nach erfolgreichem Test Access schliessen und dieses Skript mit -Mode Promote ausfuehren"
    )
}

function New-PromoteNextSteps {
    param(
        [string]$ActivePath
    )

    @(
        "Aktive FE wieder oeffnen: $ActivePath"
        "Kurzen Start-/Smoke-Test durchfuehren"
        "Bei Problemen Backup aus dem BackupFolder wiederherstellen"
    )
}

function Write-HumanReadableSummary {
    param(
        [object]$Report
    )

    Write-Output "=== Access Sync ==="
    Write-Output ("Mode: " + $Report.mode)
    if (-not [string]::IsNullOrWhiteSpace($Report.config_path)) {
        Write-Output ("Config: " + $Report.config_path)
    }
    Write-Output ("Export root: " + $Report.export_root)
    Write-Output ("DryRun: " + $Report.dry_run)
    if (-not [string]::IsNullOrWhiteSpace($Report.active_accdb)) {
        Write-Output ("Active ACCDB: " + $Report.active_accdb)
    }
    if (-not [string]::IsNullOrWhiteSpace($Report.staging_accdb)) {
        Write-Output ("Staging ACCDB: " + $Report.staging_accdb)
    }
    if (-not [string]::IsNullOrWhiteSpace($Report.target_accdb)) {
        Write-Output ("Target ACCDB: " + $Report.target_accdb)
    }
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

    if ($Report.entries.Count -gt 0) {
        Write-Output "Entries:"
        foreach ($entry in $Report.entries) {
            Write-Output ("  [{0}] {1} {2} | {3}" -f $entry.status, $entry.type, $entry.name, $entry.message)
        }
        Write-Output ""
    }

    if ($Report.next_steps.Count -gt 0) {
        Write-Output "Next manual step:"
        foreach ($step in $Report.next_steps) {
            Write-Output ("  " + $step)
        }
    }
}

$config = Resolve-SyncConfiguration -ModeName $Mode -ConfigFilePath $ConfigPath

if ([string]::IsNullOrWhiteSpace($config.source_repo_path)) {
    throw "SourceRepoPath is required."
}

$exportRoot = Resolve-ExportRoot -BasePath $config.source_repo_path
$warnings = New-Object System.Collections.ArrayList
$entries = New-Object System.Collections.ArrayList
$backupPath = ""
$nextSteps = @()
$targetPathForSummary = ""

switch ($Mode) {
    "ImportOnly" {
        if ([string]::IsNullOrWhiteSpace($config.target_accdb_path)) {
            throw "TargetAccdbPath is required for -Mode ImportOnly."
        }

        $targetPathForSummary = $config.target_accdb_path
        Ensure-TestTargetAllowed -DatabasePath $config.target_accdb_path -AllowForce $Force.IsPresent -SimulateOnly $DryRun.IsPresent

        if ($BackupBeforeImport -and -not $DryRun) {
            $backupFolderForImportOnly = if ([string]::IsNullOrWhiteSpace($config.backup_folder)) { Split-Path -Parent $config.target_accdb_path } else { $config.backup_folder }
            $backupPath = New-BackupCopy -DatabasePath $config.target_accdb_path -BackupFolderPath $backupFolderForImportOnly
        }

        $importResult = Invoke-AccessImport -ExportRoot $exportRoot -DatabasePath $config.target_accdb_path -SimulateOnly $DryRun.IsPresent
        foreach ($warning in $importResult.warnings) { [void]$warnings.Add($warning) }
        foreach ($entry in $importResult.entries) { [void]$entries.Add($entry) }
        $nextSteps = @("Access oeffnen -> Ziel-ACCDB laden -> Debug > Kompilieren -> Smoke-Test ausfuehren")
    }

    "Stage" {
        if ([string]::IsNullOrWhiteSpace($config.active_accdb_path)) {
            throw "ActiveAccdbPath is required for -Mode Stage."
        }
        if ([string]::IsNullOrWhiteSpace($config.staging_accdb_path)) {
            throw "StagingAccdbPath is required for -Mode Stage."
        }
        if ([string]::IsNullOrWhiteSpace($config.backup_folder)) {
            throw "BackupFolder is required for -Mode Stage."
        }
        if ($config.active_accdb_path -ieq $config.staging_accdb_path) {
            throw "ActiveAccdbPath and StagingAccdbPath must not point to the same file."
        }
        if (-not (Test-Path -LiteralPath $config.active_accdb_path)) {
            throw "ActiveAccdbPath does not exist: $($config.active_accdb_path)"
        }

        Ensure-AccdbPathLooksValid -DatabasePath $config.active_accdb_path -ParameterName "ActiveAccdbPath"
        Ensure-AccdbPathLooksValid -DatabasePath $config.staging_accdb_path -ParameterName "StagingAccdbPath"
        if (-not $DryRun) {
            Ensure-DatabaseClosed -DatabasePath $config.active_accdb_path -Label "Active ACCDB"
            if (Test-Path -LiteralPath $config.staging_accdb_path) {
                Ensure-DatabaseClosed -DatabasePath $config.staging_accdb_path -Label "Staging ACCDB"
            }
        } else {
            if (-not (Test-DatabaseClosed -DatabasePath $config.active_accdb_path)) {
                [void]$warnings.Add("Active ACCDB appears to be open or locked. Real Stage execution would fail until Access is closed.")
            }
            if ((Test-Path -LiteralPath $config.staging_accdb_path) -and -not (Test-DatabaseClosed -DatabasePath $config.staging_accdb_path)) {
                [void]$warnings.Add("Staging ACCDB appears to be open or locked. Real Stage execution would fail until Access is closed.")
            }
        }

        $targetPathForSummary = $config.staging_accdb_path

        if ($DryRun) {
            [void]$entries.Add((New-SyncEntry "backup" ([System.IO.Path]::GetFileName($config.active_accdb_path)) "planned" "Active ACCDB would be backed up before staging."))
            [void]$entries.Add((New-SyncEntry "staging" ([System.IO.Path]::GetFileName($config.staging_accdb_path)) "planned" "Staging ACCDB would be recreated from the active ACCDB."))
        } else {
            $backupPath = New-BackupCopy -DatabasePath $config.active_accdb_path -BackupFolderPath $config.backup_folder
            [void]$entries.Add((New-SyncEntry "backup" ([System.IO.Path]::GetFileName($config.active_accdb_path)) "created" "Backup created successfully."))

            if (Test-Path -LiteralPath $config.staging_accdb_path) {
                Remove-Item -LiteralPath $config.staging_accdb_path -Force
            }
            Copy-Item -LiteralPath $config.active_accdb_path -Destination $config.staging_accdb_path -Force
            [void]$entries.Add((New-SyncEntry "staging" ([System.IO.Path]::GetFileName($config.staging_accdb_path)) "created" "Staging ACCDB created from active ACCDB."))
        }

        $importResult = Invoke-AccessImport -ExportRoot $exportRoot -DatabasePath $config.staging_accdb_path -SimulateOnly $DryRun.IsPresent
        foreach ($warning in $importResult.warnings) { [void]$warnings.Add($warning) }
        foreach ($entry in $importResult.entries) { [void]$entries.Add($entry) }
        $nextSteps = New-StageNextSteps -StagingPath $config.staging_accdb_path
    }

    "Promote" {
        if ([string]::IsNullOrWhiteSpace($config.active_accdb_path)) {
            throw "ActiveAccdbPath is required for -Mode Promote."
        }
        if ([string]::IsNullOrWhiteSpace($config.staging_accdb_path)) {
            throw "StagingAccdbPath is required for -Mode Promote."
        }
        if ([string]::IsNullOrWhiteSpace($config.backup_folder)) {
            throw "BackupFolder is required for -Mode Promote."
        }
        if (-not (Test-Path -LiteralPath $config.active_accdb_path)) {
            throw "ActiveAccdbPath does not exist: $($config.active_accdb_path)"
        }
        if (-not (Test-Path -LiteralPath $config.staging_accdb_path)) {
            throw "StagingAccdbPath does not exist: $($config.staging_accdb_path)"
        }

        if (-not $DryRun) {
            Ensure-DatabaseClosed -DatabasePath $config.active_accdb_path -Label "Active ACCDB"
            Ensure-DatabaseClosed -DatabasePath $config.staging_accdb_path -Label "Staging ACCDB"
        } else {
            if (-not (Test-DatabaseClosed -DatabasePath $config.active_accdb_path)) {
                [void]$warnings.Add("Active ACCDB appears to be open or locked. Real Promote execution would fail until Access is closed.")
            }
            if (-not (Test-DatabaseClosed -DatabasePath $config.staging_accdb_path)) {
                [void]$warnings.Add("Staging ACCDB appears to be open or locked. Real Promote execution would fail until Access is closed.")
            }
        }
        $targetPathForSummary = $config.active_accdb_path

        if ($DryRun) {
            [void]$entries.Add((New-SyncEntry "backup" ([System.IO.Path]::GetFileName($config.active_accdb_path)) "planned" "Active ACCDB would be backed up before promotion."))
            [void]$entries.Add((New-SyncEntry "promote" ([System.IO.Path]::GetFileName($config.staging_accdb_path)) "planned" "Staging ACCDB would replace the active ACCDB."))
            if ($DeleteStagingAfterPromote) {
                [void]$entries.Add((New-SyncEntry "staging" ([System.IO.Path]::GetFileName($config.staging_accdb_path)) "planned" "Staging ACCDB would be deleted after promotion."))
            }
        } else {
            $backupPath = New-BackupCopy -DatabasePath $config.active_accdb_path -BackupFolderPath $config.backup_folder
            [void]$entries.Add((New-SyncEntry "backup" ([System.IO.Path]::GetFileName($config.active_accdb_path)) "created" "Backup created successfully."))

            Copy-Item -LiteralPath $config.staging_accdb_path -Destination $config.active_accdb_path -Force
            [void]$entries.Add((New-SyncEntry "promote" ([System.IO.Path]::GetFileName($config.active_accdb_path)) "replaced" "Active ACCDB replaced from staging successfully."))

            if ($DeleteStagingAfterPromote) {
                Remove-Item -LiteralPath $config.staging_accdb_path -Force
                [void]$entries.Add((New-SyncEntry "staging" ([System.IO.Path]::GetFileName($config.staging_accdb_path)) "deleted" "Staging ACCDB deleted after promotion."))
            }
        }

        $nextSteps = New-PromoteNextSteps -ActivePath $config.active_accdb_path
    }
}

$moduleImportedCount = @($entries | Where-Object { $_.type -eq "module" -and $_.status -in @("imported", "planned") }).Count
$queryImportedCount = @($entries | Where-Object { $_.type -eq "query" -and $_.status -in @("imported", "planned") }).Count
$skippedCount = @($entries | Where-Object { $_.status -eq "skipped" }).Count
$failedCount = @($entries | Where-Object { $_.status -eq "failed" }).Count

$report = [pscustomobject]@{
    mode          = $Mode
    config_path   = $config.config_path
    export_root   = $exportRoot
    dry_run       = [bool]$DryRun
    active_accdb  = $config.active_accdb_path
    staging_accdb = $config.staging_accdb_path
    target_accdb  = $targetPathForSummary
    backup_path   = $backupPath
    warnings      = @($warnings)
    entries       = @($entries)
    next_steps    = @($nextSteps)
    summary       = [pscustomobject]@{
        modules_imported = $moduleImportedCount
        queries_imported = $queryImportedCount
        skipped          = $skippedCount
        failed           = $failedCount
    }
}

if ($AsJson) {
    $report | ConvertTo-Json -Depth 8
} else {
    Write-HumanReadableSummary -Report $report
}
