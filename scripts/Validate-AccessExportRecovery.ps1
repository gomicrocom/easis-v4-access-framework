param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$AccdbPath = "",
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function New-ValidationIssue {
    param(
        [string]$Severity,
        [string]$Category,
        [string]$Component,
        [string]$Message
    )

    [pscustomobject]@{
        severity  = $Severity
        category  = $Category
        component = $Component
        message   = $Message
    }
}

function New-CleanupCandidate {
    param(
        [string]$Component,
        [string]$Action,
        [string]$Reason,
        [string[]]$Paths
    )

    [pscustomobject]@{
        component = $Component
        action    = $Action
        reason    = $Reason
        paths     = @($Paths)
    }
}

function Test-ByteArrayStartsWith {
    param(
        [byte[]]$Bytes,
        [byte[]]$Prefix
    )

    if ($Bytes.Length -lt $Prefix.Length) {
        return $false
    }

    for ($i = 0; $i -lt $Prefix.Length; $i++) {
        if ($Bytes[$i] -ne $Prefix[$i]) {
            return $false
        }
    }

    return $true
}

function Get-FileEncodingKind {
    param(
        [string]$Path
    )

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -eq 0) {
        return "EMPTY"
    }

    if (Test-ByteArrayStartsWith -Bytes $bytes -Prefix ([byte[]](0xEF, 0xBB, 0xBF))) {
        return "UTF8_BOM"
    }

    try {
        $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
        $decoded = $utf8Strict.GetString($bytes)
        $roundtrip = [System.Text.Encoding]::UTF8.GetBytes($decoded)

        if ($roundtrip.Length -eq $bytes.Length) {
            for ($i = 0; $i -lt $bytes.Length; $i++) {
                if ($roundtrip[$i] -ne $bytes[$i]) {
                    return "ANSI_OR_OTHER"
                }
            }
            return "UTF8"
        }
    }
    catch {
        return "ANSI_OR_OTHER"
    }

    return "ANSI_OR_OTHER"
}

function Get-ArtifactFormatKind {
    param(
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return "MISSING"
    }

    $head = @(Get-Content -LiteralPath $Path -TotalCount 6)
    $joined = ($head -join "`n")
    $extension = [System.IO.Path]::GetExtension($Path)

    if ($extension -ieq ".sql") {
        return "SQL"
    }

    if ($joined -match 'VERSION 1\.0 CLASS' -and $joined -match 'Attribute VB_Name') {
        return "FULL_CLASS_EXPORT"
    }

    if ($joined -match '^Attribute VB_Name' -or $joined -match "`nAttribute VB_Name") {
        return "VB_MODULE_EXPORT"
    }

    if ($joined -match '^Option Compare Database' -or $joined -match '^Option Explicit') {
        return "CODE_BEHIND_ONLY"
    }

    return "UNKNOWN"
}

function Get-ExpectedEncodingFileSet {
    param(
        [string]$RootPath
    )

    @(Get-ChildItem -Path $RootPath -Recurse -File -Include *.bas,*.cls,*.frm,*.txt,*.sql,*.md)
}

function Get-RequiredComponentManifest {
    [pscustomobject]@{
        modules = @(
            "modBootstrap",
            "modConfigIni",
            "modConstants",
            "modGlobals",
            "modLoggingHandler",
            "modErrorHandler",
            "modTenantContext",
            "modSessionContext",
            "modTenantRepository",
            "modUserRepository",
            "modTranslationService",
            "modFwTranslationRuntime",
            "modFormRuntime",
            "modModuleManager",
            "modAppShell",
            "modAppWorkspaceService",
            "modAppNavigationService"
        )
        forms = @(
            "frmAppShell",
            "frmAppDashboard",
            "frmAppNavigation"
        )
        reports = @(
            "rpt_document",
            "srpt_document_vat_summary"
        )
        deprecated = @(
            "frmFwTranslations",
            "frmFwTranslationList"
        )
    }
}

function Get-ArtifactIdentity {
    param(
        [System.IO.FileInfo]$File,
        [string]$ExportRoot
    )

    $fullPath = $File.FullName
    $moduleDir = Join-Path $ExportRoot "modules"
    $formDir = Join-Path $ExportRoot "forms"
    $reportDir = Join-Path $ExportRoot "reports"

    if ($fullPath.StartsWith($moduleDir, [System.StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{
            object_type = "MODULE"
            object_name = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
        }
    }

    if ($fullPath.StartsWith($formDir, [System.StringComparison]::OrdinalIgnoreCase)) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
        if ($name.StartsWith("Form_", [System.StringComparison]::OrdinalIgnoreCase)) {
            $name = $name.Substring(5)
        }
        return [pscustomobject]@{
            object_type = "FORM"
            object_name = $name
        }
    }

    if ($fullPath.StartsWith($reportDir, [System.StringComparison]::OrdinalIgnoreCase)) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
        if ($name.StartsWith("Report_", [System.StringComparison]::OrdinalIgnoreCase)) {
            $name = $name.Substring(7)
        }
        return [pscustomobject]@{
            object_type = "REPORT"
            object_name = $name
        }
    }

    return [pscustomobject]@{
        object_type = "OTHER"
        object_name = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    }
}

function Test-RepositoryArtifacts {
    param(
        [string]$RootPath
    )

    $exportRoot = Join-Path $RootPath "src\access\exported"
    $manifest = Get-RequiredComponentManifest
    $issues = New-Object System.Collections.ArrayList
    $cleanupCandidates = New-Object System.Collections.ArrayList

    $moduleDir = Join-Path $exportRoot "modules"
    $formDir = Join-Path $exportRoot "forms"
    $reportDir = Join-Path $exportRoot "reports"
    $queryDir = Join-Path $exportRoot "queries"

    $moduleFiles = @(Get-ChildItem -LiteralPath $moduleDir -Filter *.bas -ErrorAction SilentlyContinue)
    $formFiles = @(Get-ChildItem -LiteralPath $formDir -Filter *.cls -ErrorAction SilentlyContinue)
    $reportFiles = @(Get-ChildItem -LiteralPath $reportDir -Filter *.cls -ErrorAction SilentlyContinue)
    $queryFiles = @(Get-ChildItem -LiteralPath $queryDir -Filter *.sql -ErrorAction SilentlyContinue)

    foreach ($moduleName in $manifest.modules) {
        $path = Join-Path $moduleDir ($moduleName + ".bas")
        if (-not (Test-Path -LiteralPath $path)) {
            [void]$issues.Add((New-ValidationIssue "ERROR" "repo.module.missing" $moduleName "Required module export file is missing."))
            continue
        }

        $item = Get-Item -LiteralPath $path
        if ($item.Length -le 0) {
            [void]$issues.Add((New-ValidationIssue "ERROR" "repo.module.empty" $moduleName "Required module export file is empty."))
        }

        $formatKind = Get-ArtifactFormatKind -Path $path
        if ($formatKind -notin @("VB_MODULE_EXPORT", "FULL_CLASS_EXPORT")) {
            [void]$issues.Add((New-ValidationIssue "WARN" "repo.module.format" $moduleName ("Unexpected module export format: " + $formatKind)))
            [void]$cleanupCandidates.Add((New-CleanupCandidate $moduleName "MIGRATE" "Module should use a consistent VBA export format with Attribute VB_Name." @($path)))
        }
    }

    foreach ($formName in $manifest.forms) {
        $path = Join-Path $formDir ("Form_" + $formName + ".cls")
        if (-not (Test-Path -LiteralPath $path)) {
            [void]$issues.Add((New-ValidationIssue "ERROR" "repo.form.missing" $formName "Required form code export file is missing."))
            continue
        }

        $item = Get-Item -LiteralPath $path
        if ($item.Length -le 0) {
            [void]$issues.Add((New-ValidationIssue "ERROR" "repo.form.empty" $formName "Required form code export file is empty."))
        }

        $formatKind = Get-ArtifactFormatKind -Path $path
        if ($formatKind -eq "CODE_BEHIND_ONLY") {
            [void]$issues.Add((New-ValidationIssue "WARN" "repo.form.partial" $formName "Form export looks like code-behind only, not a full SaveAsText class export."))
            [void]$cleanupCandidates.Add((New-CleanupCandidate $formName "MIGRATE" "Form currently looks like code-behind only; target state is one canonical full form export artifact." @($path)))
        }
    }

    foreach ($reportName in $manifest.reports) {
        $path = Join-Path $reportDir ("Report_" + $reportName + ".cls")
        if (-not (Test-Path -LiteralPath $path)) {
            [void]$issues.Add((New-ValidationIssue "WARN" "repo.report.missing" $reportName "Expected report code export file is missing."))
            continue
        }

        $formatKind = Get-ArtifactFormatKind -Path $path
        if ($formatKind -eq "CODE_BEHIND_ONLY") {
            [void]$issues.Add((New-ValidationIssue "WARN" "repo.report.partial" $reportName "Report export looks like code-behind only, not a full SaveAsText class export."))
            [void]$cleanupCandidates.Add((New-CleanupCandidate $reportName "MIGRATE" "Report currently looks like code-behind only; target state is one canonical full report export artifact." @($path)))
        }
    }

    foreach ($file in @($moduleFiles + $formFiles + $reportFiles)) {
        if ($file.Length -le 0) {
            [void]$issues.Add((New-ValidationIssue "ERROR" "repo.file.empty" $file.Name "Export file is empty."))
        }

        $tail = @((Get-Content -LiteralPath $file.FullName | Where-Object { $_.Trim().Length -gt 0 }) | Select-Object -Last 1)
        if ($tail.Count -eq 0) {
            [void]$issues.Add((New-ValidationIssue "ERROR" "repo.file.blank" $file.Name "Export file has no non-empty content."))
        }
    }

    $encodingFiles = @(Get-ExpectedEncodingFileSet -RootPath $RootPath)
    foreach ($file in $encodingFiles) {
        $encodingKind = Get-FileEncodingKind -Path $file.FullName
        $relativePath = $file.FullName.Substring($RootPath.Length + 1)

        if ($encodingKind -eq "ANSI_OR_OTHER") {
            [void]$issues.Add((New-ValidationIssue "WARN" "repo.encoding.non_utf8" $relativePath "File is not valid UTF-8 and should be normalized."))
            [void]$cleanupCandidates.Add((New-CleanupCandidate $relativePath "MIGRATE" "Normalize file encoding to UTF-8 using the repository export convention." @($relativePath)))
        }
    }

    $legacyLocalizationTxt = Join-Path $moduleDir "Form_frmLocalisation.txt"
    $legacyLocalizationCls = Join-Path $formDir "Form_frmLocalisation.cls"
    if ((Test-Path -LiteralPath $legacyLocalizationTxt) -and (Test-Path -LiteralPath $legacyLocalizationCls)) {
        [void]$issues.Add((New-ValidationIssue "WARN" "repo.duplicate.legacy" "frmLocalisation" "Legacy duplicate export artifacts exist in both forms and modules folders."))
        [void]$cleanupCandidates.Add((New-CleanupCandidate "frmLocalisation" "REVIEW" "Duplicate legacy artifacts exist; keep one canonical artifact after confirming import path." @($legacyLocalizationTxt, $legacyLocalizationCls)))
    }

    $allArtifactFiles = @($moduleFiles + $formFiles + $reportFiles)
    $groups = @{}
    foreach ($file in $allArtifactFiles) {
        $identity = Get-ArtifactIdentity -File $file -ExportRoot $exportRoot
        $key = $identity.object_type + ":" + $identity.object_name.ToUpperInvariant()
        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = New-Object System.Collections.ArrayList
        }
        [void]$groups[$key].Add($file.FullName)
    }

    foreach ($key in $groups.Keys) {
        $paths = @($groups[$key])
        if ($paths.Count -gt 1) {
            [void]$issues.Add((New-ValidationIssue "WARN" "repo.duplicate.artifact" $key "Multiple export artifacts exist for the same object identity."))
            [void]$cleanupCandidates.Add((New-CleanupCandidate $key "REVIEW" "Multiple artifacts map to the same object identity; retain exactly one canonical artifact." $paths))
        }
    }

    foreach ($deprecatedName in $manifest.deprecated) {
        $deprecatedPath = Join-Path $formDir ("Form_" + $deprecatedName + ".cls")
        if (Test-Path -LiteralPath $deprecatedPath) {
            [void]$issues.Add((New-ValidationIssue "WARN" "repo.deprecated.object" $deprecatedName "Deprecated object artifact is still present."))
            [void]$cleanupCandidates.Add((New-CleanupCandidate $deprecatedName "KEEP" "Deprecated artifact is still intentionally retained for transition; do not delete blindly." @($deprecatedPath)))
        }
    }

    [pscustomobject]@{
        export_root              = $exportRoot
        module_count             = $moduleFiles.Count
        form_count               = $formFiles.Count
        report_count             = $reportFiles.Count
        query_count              = $queryFiles.Count
        export_convention        = [pscustomobject]@{
            modules = "src/access/exported/modules/*.bas -> VBA module export with Attribute VB_Name, UTF-8"
            forms   = "src/access/exported/forms/Form_<FormName>.cls -> one canonical form artifact per form, UTF-8"
            reports = "src/access/exported/reports/Report_<ReportName>.cls -> one canonical report artifact per report, UTF-8"
            queries = "src/access/exported/queries/*.sql -> UTF-8"
            texts   = "*.md, *.txt, *.frm -> UTF-8"
        }
        required_manifest        = $manifest
        issues                   = @($issues)
        cleanup_candidates       = @($cleanupCandidates)
    }
}

function Get-AccessProjectObjectNames {
    param(
        [object]$ProjectCollection
    )

    $names = New-Object System.Collections.ArrayList
    foreach ($item in $ProjectCollection) {
        [void]$names.Add([string]$item.Name)
    }
    @($names)
}

function Test-AccessProjectAgainstRepository {
    param(
        [string]$DatabasePath,
        [string]$RootPath
    )

    $issues = New-Object System.Collections.ArrayList
    $manifest = Get-RequiredComponentManifest
    $moduleDir = Join-Path $RootPath "src\access\exported\modules"
    $formDir = Join-Path $RootPath "src\access\exported\forms"
    $reportDir = Join-Path $RootPath "src\access\exported\reports"

    $access = $null
    try {
        $access = New-Object -ComObject Access.Application
        $access.OpenCurrentDatabase($DatabasePath)

        $allModules = @(Get-AccessProjectObjectNames -ProjectCollection $access.CurrentProject.AllModules)
        $allForms = @(Get-AccessProjectObjectNames -ProjectCollection $access.CurrentProject.AllForms)
        $allReports = @(Get-AccessProjectObjectNames -ProjectCollection $access.CurrentProject.AllReports)

        foreach ($moduleName in $manifest.modules) {
            if ($allModules -notcontains $moduleName) {
                [void]$issues.Add((New-ValidationIssue "ERROR" "access.module.missing" $moduleName "Required module is missing in the ACCDB project."))
            }

            $repoPath = Join-Path $moduleDir ($moduleName + ".bas")
            if (-not (Test-Path -LiteralPath $repoPath)) {
                [void]$issues.Add((New-ValidationIssue "ERROR" "repo.module.missing" $moduleName "Required module export file is missing while validating ACCDB."))
            }
        }

        foreach ($formName in $manifest.forms) {
            if ($allForms -notcontains $formName) {
                [void]$issues.Add((New-ValidationIssue "ERROR" "access.form.missing" $formName "Required form is missing in the ACCDB project."))
            }

            $repoPath = Join-Path $formDir ("Form_" + $formName + ".cls")
            if (-not (Test-Path -LiteralPath $repoPath)) {
                [void]$issues.Add((New-ValidationIssue "ERROR" "repo.form.missing" $formName "Required form code export file is missing while validating ACCDB."))
            }
        }

        foreach ($reportName in $manifest.reports) {
            if ($allReports -notcontains $reportName) {
                [void]$issues.Add((New-ValidationIssue "WARN" "access.report.missing" $reportName "Expected report is missing in the ACCDB project."))
            }

            $repoPath = Join-Path $reportDir ("Report_" + $reportName + ".cls")
            if (-not (Test-Path -LiteralPath $repoPath)) {
                [void]$issues.Add((New-ValidationIssue "WARN" "repo.report.missing" $reportName "Expected report export file is missing while validating ACCDB."))
            }
        }

        [pscustomobject]@{
            database_path = $DatabasePath
            access_modules = $allModules
            access_forms = $allForms
            access_reports = $allReports
            issues = @($issues)
        }
    }
    finally {
        if ($access -ne $null) {
            try { $access.CloseCurrentDatabase() } catch {}
            try { $access.Quit() } catch {}
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($access)
        }
    }
}

function Write-HumanReadableReport {
    param(
        [object]$RepoReport,
        [object]$AccessReport
    )

    Write-Output "=== Access Export Recovery Validation ==="
    Write-Output ("Repo root: " + $RepoRoot)
    Write-Output ("Export root: " + $RepoReport.export_root)
    Write-Output ("Counts: modules={0}; forms={1}; reports={2}; queries={3}" -f $RepoReport.module_count, $RepoReport.form_count, $RepoReport.report_count, $RepoReport.query_count)
    Write-Output ""
    Write-Output "Export convention:"
    Write-Output ("  modules: " + $RepoReport.export_convention.modules)
    Write-Output ("  forms:   " + $RepoReport.export_convention.forms)
    Write-Output ("  reports: " + $RepoReport.export_convention.reports)
    Write-Output ("  queries: " + $RepoReport.export_convention.queries)
    Write-Output ("  text:    " + $RepoReport.export_convention.texts)
    Write-Output ""
    Write-Output "Recovery checklist:"
    Write-Output "1. Start from a known-good FE.accdb if form/report layout is not fully SaveAsText-exported."
    Write-Output "2. Import required modules, forms, reports, and queries."
    Write-Output "3. Verify required modules exist in Access and in src/access/exported."
    Write-Output "4. Check Access references."
    Write-Output "5. Run Debug > Compile."
    Write-Output "6. Test bootstrap, frmAppShell, navigation, and workspace switching."
    Write-Output ""
    Write-Output "Repository issues:"
    if ($RepoReport.issues.Count -eq 0) {
        Write-Output "  none"
    } else {
        foreach ($issue in $RepoReport.issues) {
            Write-Output ("  [{0}] {1} | {2} | {3}" -f $issue.severity, $issue.category, $issue.component, $issue.message)
        }
    }

    Write-Output ""
    Write-Output "Cleanup candidates:"
    if ($RepoReport.cleanup_candidates.Count -eq 0) {
        Write-Output "  none"
    } else {
        foreach ($candidate in $RepoReport.cleanup_candidates) {
            Write-Output ("  [{0}] {1} | {2}" -f $candidate.action, $candidate.component, $candidate.reason)
            foreach ($path in $candidate.paths) {
                Write-Output ("      " + $path)
            }
        }
    }

    if ($null -ne $AccessReport) {
        Write-Output ""
        Write-Output ("ACCDB validation: " + $AccessReport.database_path)
        if ($AccessReport.issues.Count -eq 0) {
            Write-Output "  no Access-vs-repo issues detected"
        } else {
            foreach ($issue in $AccessReport.issues) {
                Write-Output ("  [{0}] {1} | {2} | {3}" -f $issue.severity, $issue.category, $issue.component, $issue.message)
            }
        }
    }
}

$repoReport = Test-RepositoryArtifacts -RootPath $RepoRoot
$accessReport = $null

if ([string]::IsNullOrWhiteSpace($AccdbPath) -eq $false) {
    $resolvedAccdbPath = (Resolve-Path -LiteralPath $AccdbPath).Path
    $accessReport = Test-AccessProjectAgainstRepository -DatabasePath $resolvedAccdbPath -RootPath $RepoRoot
}

if ($AsJson) {
    [pscustomobject]@{
        repository = $repoReport
        access = $accessReport
    } | ConvertTo-Json -Depth 8
} else {
    Write-HumanReadableReport -RepoReport $repoReport -AccessReport $accessReport
}
