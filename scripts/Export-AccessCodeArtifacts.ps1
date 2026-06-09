param(
    [Parameter(Mandatory = $true)]
    [string]$SourceAccdbPath,

    [Parameter(Mandatory = $true)]
    [string]$TargetRepoPath,

    [switch]$BackupBeforeExport,
    [switch]$DryRun,
    [switch]$AsJson,
    [bool]$IncludeModules = $true,
    [bool]$IncludeQueries = $true,
    [bool]$IncludeFormCodeBehind = $true,
    [bool]$IncludeReportCodeBehind = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$acModule = 5

function New-ExportEntry {
    param(
        [string]$Type,
        [string]$Name,
        [string]$Status,
        [string]$Message,
        [string]$TargetPath = ""
    )

    [pscustomobject]@{
        type        = $Type
        name        = $Name
        status      = $Status
        message     = $Message
        target_path = $TargetPath
    }
}

function Resolve-RepoExportRoot {
    param(
        [string]$BasePath
    )

    $resolvedBase = (Resolve-Path -LiteralPath $BasePath).Path
    $exportRoot = Join-Path $resolvedBase "src\access\exported"

    foreach ($requiredDir in @("modules", "queries", "forms", "reports")) {
        if (-not (Test-Path -LiteralPath (Join-Path $exportRoot $requiredDir))) {
            throw "TargetRepoPath does not contain src/access/exported/$requiredDir."
        }
    }

    return $exportRoot
}

function Write-Utf8File {
    param(
        [string]$Path,
        [string]$Content
    )

    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8)
}

function Read-AccessExportText {
    param(
        [string]$Path
    )

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -eq 0) {
        return ""
    }

    try {
        $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
        return $utf8Strict.GetString($bytes)
    }
    catch {
        $ansi = [System.Text.Encoding]::GetEncoding(1252)
        return $ansi.GetString($bytes)
    }
}

function Ensure-ModuleHeader {
    param(
        [string]$ModuleName,
        [string]$Content
    )

    if ([string]::IsNullOrWhiteSpace($Content)) {
        return ('Attribute VB_Name = "' + $ModuleName + '"' + "`r`n")
    }

    if ([System.Text.RegularExpressions.Regex]::IsMatch($Content, '^Attribute VB_Name\s*=\s*".+"')) {
        return $Content
    }

    return ('Attribute VB_Name = "' + $ModuleName + '"' + "`r`n" + $Content)
}

function Build-CodeBehindClassExport {
    param(
        [string]$ComponentName,
        [string]$CodeText
    )

    $lines = @(
        "VERSION 1.0 CLASS",
        "BEGIN",
        "  MultiUse = -1  'True",
        "END",
        ('Attribute VB_Name = "' + $ComponentName + '"'),
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = True",
        "Attribute VB_PredeclaredId = True",
        "Attribute VB_Exposed = False",
        "' ExportKind: CODE_BEHIND_ONLY"
    )

    if ([string]::IsNullOrEmpty($CodeText)) {
        return ($lines -join "`r`n") + "`r`n"
    }

    return ($lines -join "`r`n") + "`r`n" + $CodeText
}

function Backup-ExistingFile {
    param(
        [string]$SourcePath,
        [string]$BackupRoot,
        [string]$RelativePath,
        [bool]$SimulateOnly
    )

    if (-not (Test-Path -LiteralPath $SourcePath)) {
        return
    }

    if ($SimulateOnly) {
        return
    }

    $destination = Join-Path $BackupRoot $RelativePath
    $destinationDir = Split-Path -Parent $destination
    New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
    Copy-Item -LiteralPath $SourcePath -Destination $destination -Force
}

function Export-StandardModules {
    param(
        [object]$AccessApplication,
        [string]$TargetModuleDir,
        [string]$BackupRoot,
        [bool]$CreateBackup,
        [bool]$SimulateOnly
    )

    $entries = New-Object System.Collections.ArrayList
    $changedFiles = New-Object System.Collections.ArrayList
    $moduleNames = New-Object System.Collections.ArrayList

    foreach ($moduleItem in $AccessApplication.CurrentProject.AllModules) {
        [void]$moduleNames.Add([string]$moduleItem.Name)
    }

    foreach ($moduleName in ($moduleNames | Sort-Object)) {
        $targetPath = Join-Path $TargetModuleDir ($moduleName + ".bas")

        if ($SimulateOnly) {
            [void]$entries.Add((New-ExportEntry "module" $moduleName "planned" "Module would be exported and normalized to UTF-8." $targetPath))
            [void]$changedFiles.Add($targetPath)
            continue
        }

        $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName() + ".bas")
        try {
            $AccessApplication.SaveAsText($acModule, $moduleName, $tempPath)
            $content = Read-AccessExportText -Path $tempPath
            $content = Ensure-ModuleHeader -ModuleName $moduleName -Content $content

            if ($CreateBackup) {
                Backup-ExistingFile -SourcePath $targetPath -BackupRoot $BackupRoot -RelativePath (Join-Path "src\access\exported\modules" ($moduleName + ".bas")) -SimulateOnly $false
            }

            Write-Utf8File -Path $targetPath -Content $content
            [void]$entries.Add((New-ExportEntry "module" $moduleName "exported" "Module exported and normalized to UTF-8." $targetPath))
            [void]$changedFiles.Add($targetPath)
        }
        catch {
            [void]$entries.Add((New-ExportEntry "module" $moduleName "failed" $_.Exception.Message $targetPath))
        }
        finally {
            try {
                if (Test-Path -LiteralPath $tempPath) {
                    Remove-Item -LiteralPath $tempPath -Force
                }
            }
            catch {}
        }
    }

    return [pscustomobject]@{
        entries = @($entries)
        changed_files = @($changedFiles)
    }
}

function Export-Queries {
    param(
        [object]$AccessApplication,
        [string]$TargetQueryDir,
        [string]$BackupRoot,
        [bool]$CreateBackup,
        [bool]$SimulateOnly
    )

    $entries = New-Object System.Collections.ArrayList
    $changedFiles = New-Object System.Collections.ArrayList
    $db = $AccessApplication.CurrentDb()

    foreach ($queryDef in @($db.QueryDefs | Sort-Object Name)) {
        $queryName = [string]$queryDef.Name

        if ($queryName.StartsWith("~")) {
            [void]$entries.Add((New-ExportEntry "query" $queryName "skipped" "Temporary query was ignored."))
            continue
        }

        if ($queryName.StartsWith("MSys", [System.StringComparison]::OrdinalIgnoreCase)) {
            [void]$entries.Add((New-ExportEntry "query" $queryName "skipped" "System query was ignored."))
            continue
        }

        if ([string]::IsNullOrWhiteSpace($queryDef.SQL)) {
            [void]$entries.Add((New-ExportEntry "query" $queryName "skipped" "Query has no SQL text."))
            continue
        }

        $targetPath = Join-Path $TargetQueryDir ($queryName + ".sql")

        if ($SimulateOnly) {
            [void]$entries.Add((New-ExportEntry "query" $queryName "planned" "Query would be exported as UTF-8 SQL." $targetPath))
            [void]$changedFiles.Add($targetPath)
            continue
        }

        try {
            if ($CreateBackup) {
                Backup-ExistingFile -SourcePath $targetPath -BackupRoot $BackupRoot -RelativePath (Join-Path "src\access\exported\queries" ($queryName + ".sql")) -SimulateOnly $false
            }

            Write-Utf8File -Path $targetPath -Content ([string]$queryDef.SQL)
            [void]$entries.Add((New-ExportEntry "query" $queryName "exported" "Query exported as UTF-8 SQL." $targetPath))
            [void]$changedFiles.Add($targetPath)
        }
        catch {
            [void]$entries.Add((New-ExportEntry "query" $queryName "failed" $_.Exception.Message $targetPath))
        }
    }

    return [pscustomobject]@{
        entries = @($entries)
        changed_files = @($changedFiles)
    }
}

function Export-CodeBehindArtifacts {
    param(
        [object]$AccessApplication,
        [string]$CollectionName,
        [string]$ComponentPrefix,
        [string]$TargetDir,
        [string]$ArtifactPrefix,
        [string]$ArtifactType,
        [string]$BackupRoot,
        [bool]$CreateBackup,
        [bool]$SimulateOnly
    )

    $entries = New-Object System.Collections.ArrayList
    $changedFiles = New-Object System.Collections.ArrayList
    $project = $AccessApplication.VBE.ActiveVBProject
    $objects = $AccessApplication.CurrentProject.$CollectionName

    foreach ($projectItem in @($objects | Sort-Object Name)) {
        $objectName = [string]$projectItem.Name
        $componentName = $ComponentPrefix + $objectName
        $targetPath = Join-Path $TargetDir ($ArtifactPrefix + $objectName + ".cls")

        try {
            $vbComponent = $project.VBComponents.Item($componentName)
        }
        catch {
            [void]$entries.Add((New-ExportEntry $ArtifactType $objectName "skipped" "$($ArtifactType.Substring(0,1).ToUpper() + $ArtifactType.Substring(1)) has no code-behind; skipped code export." $targetPath))
            continue
        }

        if ($SimulateOnly) {
            [void]$entries.Add((New-ExportEntry $ArtifactType $objectName "planned" "Code-behind would be exported as canonical UTF-8 class artifact." $targetPath))
            [void]$changedFiles.Add($targetPath)
            continue
        }

        try {
            $codeModule = $vbComponent.CodeModule
            $lineCount = [int]$codeModule.CountOfLines
            $codeText = if ($lineCount -gt 0) { [string]$codeModule.Lines(1, $lineCount) } else { "" }
            $content = Build-CodeBehindClassExport -ComponentName $componentName -CodeText $codeText

            if ($CreateBackup) {
                $relativePath = Join-Path ("src\access\exported\" + $ArtifactType + "s") ($ArtifactPrefix + $objectName + ".cls")
                Backup-ExistingFile -SourcePath $targetPath -BackupRoot $BackupRoot -RelativePath $relativePath -SimulateOnly $false
            }

            Write-Utf8File -Path $targetPath -Content $content
            [void]$entries.Add((New-ExportEntry $ArtifactType $objectName "exported" "Code-behind exported as canonical UTF-8 class artifact." $targetPath))
            [void]$changedFiles.Add($targetPath)
        }
        catch {
            [void]$entries.Add((New-ExportEntry $ArtifactType $objectName "failed" $_.Exception.Message $targetPath))
        }
    }

    return [pscustomobject]@{
        entries = @($entries)
        changed_files = @($changedFiles)
    }
}

function Write-HumanReadableSummary {
    param(
        [object]$Report
    )

    Write-Output "=== Access Export / Normalize Code Artifacts ==="
    Write-Output ("Source ACCDB: " + $Report.source_accdb)
    Write-Output ("Target export root: " + $Report.export_root)
    Write-Output ("DryRun: " + $Report.dry_run)
    Write-Output ("Backup root: " + $(if ([string]::IsNullOrWhiteSpace($Report.backup_root)) { "(none)" } else { $Report.backup_root }))
    Write-Output ""
    Write-Output ("Exported modules count: " + $Report.summary.modules_exported)
    Write-Output ("Exported queries count: " + $Report.summary.queries_exported)
    Write-Output ("Exported form code-behind count: " + $Report.summary.forms_exported)
    Write-Output ("Exported report code-behind count: " + $Report.summary.reports_exported)
    Write-Output ("Skipped count: " + $Report.summary.skipped)
    Write-Output ("Errors count: " + $Report.summary.failed)
    Write-Output ""
    Write-Output "Changed files:"
    foreach ($path in $Report.changed_files) {
        Write-Output ("  " + $path)
    }
    Write-Output ""
    Write-Output "Entries:"
    foreach ($entry in $Report.entries) {
        Write-Output ("  [{0}] {1} {2} | {3}" -f $entry.status, $entry.type, $entry.name, $entry.message)
    }
    Write-Output ""
    Write-Output "Next step:"
    Write-Output "  git diff pruefen"
    Write-Output "  scripts\\Validate-AccessExportRecovery.ps1 ausfuehren"
}

$resolvedSourceAccdb = (Resolve-Path -LiteralPath $SourceAccdbPath).Path
$exportRoot = Resolve-RepoExportRoot -BasePath $TargetRepoPath
$targetModules = Join-Path $exportRoot "modules"
$targetQueries = Join-Path $exportRoot "queries"
$targetForms = Join-Path $exportRoot "forms"
$targetReports = Join-Path $exportRoot "reports"
$repoRoot = (Resolve-Path -LiteralPath $TargetRepoPath).Path

$backupRoot = ""
if ($BackupBeforeExport) {
    $backupRoot = Join-Path $repoRoot ("backups\access-export_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
    if (-not $DryRun) {
        New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
    }
}

$entries = New-Object System.Collections.ArrayList
$changedFiles = New-Object System.Collections.ArrayList
$access = $null

try {
    $access = New-Object -ComObject Access.Application
    $access.OpenCurrentDatabase($resolvedSourceAccdb)

    if ($IncludeModules) {
        $result = Export-StandardModules -AccessApplication $access -TargetModuleDir $targetModules -BackupRoot $backupRoot -CreateBackup $BackupBeforeExport.IsPresent -SimulateOnly $DryRun.IsPresent
        foreach ($entry in $result.entries) { [void]$entries.Add($entry) }
        foreach ($path in $result.changed_files) { [void]$changedFiles.Add($path) }
    }

    if ($IncludeQueries) {
        $result = Export-Queries -AccessApplication $access -TargetQueryDir $targetQueries -BackupRoot $backupRoot -CreateBackup $BackupBeforeExport.IsPresent -SimulateOnly $DryRun.IsPresent
        foreach ($entry in $result.entries) { [void]$entries.Add($entry) }
        foreach ($path in $result.changed_files) { [void]$changedFiles.Add($path) }
    }

    if ($IncludeFormCodeBehind) {
        $result = Export-CodeBehindArtifacts -AccessApplication $access -CollectionName "AllForms" -ComponentPrefix "Form_" -TargetDir $targetForms -ArtifactPrefix "Form_" -ArtifactType "form" -BackupRoot $backupRoot -CreateBackup $BackupBeforeExport.IsPresent -SimulateOnly $DryRun.IsPresent
        foreach ($entry in $result.entries) { [void]$entries.Add($entry) }
        foreach ($path in $result.changed_files) { [void]$changedFiles.Add($path) }
    }

    if ($IncludeReportCodeBehind) {
        $result = Export-CodeBehindArtifacts -AccessApplication $access -CollectionName "AllReports" -ComponentPrefix "Report_" -TargetDir $targetReports -ArtifactPrefix "Report_" -ArtifactType "report" -BackupRoot $backupRoot -CreateBackup $BackupBeforeExport.IsPresent -SimulateOnly $DryRun.IsPresent
        foreach ($entry in $result.entries) { [void]$entries.Add($entry) }
        foreach ($path in $result.changed_files) { [void]$changedFiles.Add($path) }
    }
}
finally {
    if ($access -ne $null) {
        try { $access.CloseCurrentDatabase() } catch {}
        try { $access.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($access)
    }
}

$report = [pscustomobject]@{
    source_accdb = $resolvedSourceAccdb
    export_root = $exportRoot
    dry_run = [bool]$DryRun
    backup_root = $backupRoot
    entries = @($entries)
    changed_files = @($changedFiles | Select-Object -Unique)
    summary = [pscustomobject]@{
        modules_exported = @($entries | Where-Object { $_.type -eq "module" -and $_.status -in @("exported", "planned") }).Count
        queries_exported = @($entries | Where-Object { $_.type -eq "query" -and $_.status -in @("exported", "planned") }).Count
        forms_exported = @($entries | Where-Object { $_.type -eq "form" -and $_.status -in @("exported", "planned") }).Count
        reports_exported = @($entries | Where-Object { $_.type -eq "report" -and $_.status -in @("exported", "planned") }).Count
        skipped = @($entries | Where-Object { $_.status -eq "skipped" }).Count
        failed = @($entries | Where-Object { $_.status -eq "failed" }).Count
    }
}

if ($AsJson) {
    $report | ConvertTo-Json -Depth 8
} else {
    Write-HumanReadableSummary -Report $report
}
