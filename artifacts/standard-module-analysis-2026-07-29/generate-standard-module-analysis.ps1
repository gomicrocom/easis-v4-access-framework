$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$moduleDir = Join-Path $repoRoot "src\access\exported\modules"
$outDir = $PSScriptRoot

function Remove-InlineComment {
    param([string]$line)

    $inString = $false
    for ($i = 0; $i -lt $line.Length; $i++) {
        $char = $line[$i]
        if ($char -eq '"') {
            if ($inString -and $i + 1 -lt $line.Length -and $line[$i + 1] -eq '"') {
                $i++
                continue
            }

            $inString = -not $inString
            continue
        }

        if (-not $inString -and $char -eq "'") {
            return $line.Substring(0, $i)
        }
    }

    return $line
}

function Get-KeywordSet {
    $keywords = @(
        "And","As","Boolean","ByRef","ByVal","Case","Const","Currency","Date","Debug","Dim","Do","Double",
        "Each","Else","ElseIf","Empty","End","Eqv","Erase","Error","Event","Exit","False","For","Friend",
        "Function","Get","Global","GoSub","GoTo","If","Imp","In","Integer","Is","LBound","Let","Like","Lock",
        "Long","Loop","LSet","Me","Mod","New","Next","Not","Nothing","Null","Object","On","Open","Option","Optional",
        "Or","ParamArray","Preserve","Private","Property","Public","RaiseEvent","Randomize","ReDim","Rem","Resume",
        "Return","RSet","Select","Set","Single","Static","Step","Stop","String","Sub","Then","To","True","Type",
        "UBound","Variant","Wend","While","With","Xor"
    )

    $set = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($keyword in $keywords) {
        [void]$set.Add($keyword)
    }

    return $set
}

function Get-LogicalLines {
    param([string[]]$lines)

    $logicalLines = New-Object System.Collections.Generic.List[object]
    $buffer = ""
    $startLine = 0

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $currentLine = $lines[$i]
        if ([string]::IsNullOrEmpty($buffer)) {
            $buffer = $currentLine
            $startLine = $i + 1
        }
        else {
            $buffer += "`r`n" + $currentLine
        }

        $trimmed = $currentLine.TrimEnd()
        if ($trimmed.EndsWith(" _")) {
            $buffer = $buffer.Substring(0, $buffer.Length - 2)
            continue
        }

        $logicalLines.Add([PSCustomObject]@{
            StartLine = $startLine
            EndLine   = $i + 1
            Text      = $buffer
        }) | Out-Null

        $buffer = ""
        $startLine = 0
    }

    if (-not [string]::IsNullOrEmpty($buffer)) {
        $logicalLines.Add([PSCustomObject]@{
            StartLine = $startLine
            EndLine   = $lines.Count
            Text      = $buffer
        }) | Out-Null
    }

    return $logicalLines
}

function Parse-Parameters {
    param([string]$signatureText)

    $open = $signatureText.IndexOf("(")
    $close = $signatureText.LastIndexOf(")")
    if ($open -lt 0 -or $close -le $open) {
        return @()
    }

    $parameterText = $signatureText.Substring($open + 1, $close - $open - 1).Trim()
    if ([string]::IsNullOrWhiteSpace($parameterText)) {
        return @()
    }

    return ($parameterText -split ",").ForEach({ $_.Trim() })
}

function Get-ReturnType {
    param([string]$signatureText, [string]$kind)

    if ($kind -eq "Sub") {
        return ""
    }

    if ($signatureText -match "\)\s+As\s+(.+)$") {
        return $matches[1].Trim()
    }

    return "Variant"
}

function Get-BodyMetrics {
    param(
        [string[]]$bodyLines,
        [string[]]$modulePrivateVariables
    )

    $commentLines = 0
    $codeLines = 0
    $usesCurrentDb = $false
    $usesExplicitDaoDatabase = $false
    $usesPrivateVariables = New-Object System.Collections.Generic.List[string]
    $hasErrorHandling = $false
    $hasLogging = $false
    $sideEffects = New-Object System.Collections.Generic.List[string]

    $effectPatterns = @{
        "DB_EXECUTE"      = "\.Execute\b|CurrentDb\s*\(\)\s*\.Execute\b|db\.Execute\b"
        "DB_MUTATION"     = "\b(AddNew|Edit|Update|Delete)\b"
        "FORM_NAVIGATION" = "\bDoCmd\.(OpenForm|Close|GoToRecord|Quit)\b"
        "MESSAGE_BOX"     = "\bMsgBox\b"
        "FILE_IO"         = "\b(FileCopy|MkDir|RmDir|Kill |Name |Open |Print #|Write #|Close #)\b"
        "TRANSACTION"     = "\b(BeginTrans|CommitTrans|Rollback)\b"
    }

    foreach ($rawLine in $bodyLines) {
        $trimmed = $rawLine.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }

        if ($trimmed.StartsWith("'")) {
            $commentLines++
            continue
        }

        $codeLines++
        $lineWithoutComment = Remove-InlineComment $rawLine

        if ($lineWithoutComment -match "\bCurrentDb\b") {
            $usesCurrentDb = $true
        }

        if ($lineWithoutComment -match "\bDAO\.Database\b" -or $lineWithoutComment -match "\bAs\s+Database\b") {
            $usesExplicitDaoDatabase = $true
        }

        if ($lineWithoutComment -match "^\s*On\s+Error\b" -or $lineWithoutComment -match "\bResume\b") {
            $hasErrorHandling = $true
        }

        if ($lineWithoutComment -match "\bmodLoggingHandler\." -or $lineWithoutComment -match "\bLog(Info|Warning|Error|Debug)\b") {
            $hasLogging = $true
        }

        foreach ($privateVar in $modulePrivateVariables) {
            if ($lineWithoutComment -match ("(?i)\b" + [regex]::Escape($privateVar) + "\b")) {
                if (-not $usesPrivateVariables.Contains($privateVar)) {
                    $usesPrivateVariables.Add($privateVar) | Out-Null
                }
            }
        }

        foreach ($key in $effectPatterns.Keys) {
            if ($lineWithoutComment -match $effectPatterns[$key]) {
                if (-not $sideEffects.Contains($key)) {
                    $sideEffects.Add($key) | Out-Null
                }
            }
        }
    }

    return [PSCustomObject]@{
        CodeLines                 = $codeLines
        CommentLines              = $commentLines
        UsesCurrentDb             = $usesCurrentDb
        UsesExplicitDaoDatabase   = $usesExplicitDaoDatabase
        PrivateVariables          = ($usesPrivateVariables -join ";")
        HasErrorHandling          = $hasErrorHandling
        HasLogging                = $hasLogging
        SideEffects               = ($sideEffects -join ";")
    }
}

function Get-NormalizedExactBody {
    param([string[]]$bodyLines)

    $normalized = New-Object System.Collections.Generic.List[string]
    foreach ($line in $bodyLines) {
        $lineWithoutComment = (Remove-InlineComment $line).Trim()
        if ([string]::IsNullOrWhiteSpace($lineWithoutComment)) {
            continue
        }

        $normalized.Add(($lineWithoutComment.ToLowerInvariant() -replace "\s+", " ")) | Out-Null
    }

    return ($normalized -join "`n")
}

function Get-NormalizedSkeleton {
    param(
        [string[]]$bodyLines,
        [System.Collections.Generic.HashSet[string]]$keywordSet
    )

    $raw = Get-NormalizedExactBody $bodyLines
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return ""
    }

    $raw = [regex]::Replace($raw, """(?:""""|[^""])*""", '"STR"')
    $raw = [regex]::Replace($raw, "\b\d+(\.\d+)?\b", "NUM")

    $identifierMap = @{}
    $identifierIndex = 1

    return [regex]::Replace($raw, "\b[A-Za-z_][A-Za-z0-9_]*\b", {
        param($match)
        $value = $match.Value
        if ($keywordSet.Contains($value)) {
            return $value.ToLowerInvariant()
        }

        if (-not $identifierMap.ContainsKey($value)) {
            $identifierMap[$value] = "id$identifierIndex"
            $identifierIndex++
        }

        return $identifierMap[$value]
    })
}

function Get-ProcedureCatalog {
    param([string]$modulePath)

    $moduleName = [IO.Path]::GetFileNameWithoutExtension($modulePath)
    $lines = Get-Content -Path $modulePath -Encoding UTF8
    $logicalLines = Get-LogicalLines $lines

    $modulePrivateVariables = @()
    foreach ($line in $logicalLines) {
        if ($line.Text -match "^\s*Private\s+([A-Za-z_][A-Za-z0-9_]*)\s+As\s+") {
            $modulePrivateVariables += $matches[1]
        }
    }

    $keywordSet = Get-KeywordSet
    $procedures = New-Object System.Collections.Generic.List[object]
    $currentProcedure = $null

    foreach ($line in $logicalLines) {
        $text = $line.Text

        if ($null -eq $currentProcedure) {
            if ($text -match "^\s*(Public|Private)\s+(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\b") {
                $currentProcedure = [ordered]@{
                    ModuleName    = $moduleName
                    ModulePath    = $modulePath
                    Visibility    = $matches[1]
                    Kind          = $matches[2]
                    ProcedureName = $matches[3]
                    Signature     = ($text -replace "\r?\n", " ").Trim()
                    StartLine     = $line.StartLine
                    EndLine       = $line.EndLine
                    BodyLines     = New-Object System.Collections.Generic.List[string]
                }
            }

            continue
        }

        $currentProcedure.EndLine = $line.EndLine
        $currentProcedure.BodyLines.Add($text) | Out-Null

        if ($text -match "^\s*End\s+" + [regex]::Escape($currentProcedure.Kind) + "\s*$") {
            $bodyLines = @($currentProcedure.BodyLines)
            $metrics = Get-BodyMetrics -bodyLines $bodyLines -modulePrivateVariables $modulePrivateVariables
            $parameters = Parse-Parameters $currentProcedure.Signature

            $procedures.Add([PSCustomObject]@{
                ModuleName                  = $currentProcedure.ModuleName
                ModulePath                  = $currentProcedure.ModulePath
                ProcedureName               = $currentProcedure.ProcedureName
                Kind                        = $currentProcedure.Kind
                Visibility                  = $currentProcedure.Visibility
                Signature                   = $currentProcedure.Signature
                ReturnType                  = Get-ReturnType -signatureText $currentProcedure.Signature -kind $currentProcedure.Kind
                Parameters                  = ($parameters -join "; ")
                ParameterCount              = $parameters.Count
                StartLine                   = $currentProcedure.StartLine
                EndLine                     = $currentProcedure.EndLine
                CodeLines                   = $metrics.CodeLines
                CommentLines                = $metrics.CommentLines
                UsesCurrentDb               = $metrics.UsesCurrentDb
                UsesExplicitDaoDatabase     = $metrics.UsesExplicitDaoDatabase
                UsesPrivateModuleVariables  = $metrics.PrivateVariables
                HasErrorHandling            = $metrics.HasErrorHandling
                HasLogging                  = $metrics.HasLogging
                SideEffects                 = $metrics.SideEffects
                ExactBodySignature          = Get-NormalizedExactBody $bodyLines
                SkeletonSignature           = Get-NormalizedSkeleton -bodyLines $bodyLines -keywordSet $keywordSet
            }) | Out-Null

            $currentProcedure = $null
        }
    }

    return $procedures
}

function Get-CallSites {
    param([object[]]$procedures)

    $callSites = New-Object System.Collections.Generic.List[object]
    $proceduresByModule = $procedures | Group-Object ModuleName -AsHashTable -AsString
    $allNames = $procedures.ProcedureName | Sort-Object -Unique

    foreach ($moduleName in $proceduresByModule.Keys) {
        $modulePath = $proceduresByModule[$moduleName][0].ModulePath
        $lines = Get-Content -Path $modulePath -Encoding UTF8

        for ($i = 0; $i -lt $lines.Count; $i++) {
            $lineNumber = $i + 1
            $cleanLine = Remove-InlineComment $lines[$i]
            if ($cleanLine -match "^\s*(Public|Private)\s+(Sub|Function)\s+") {
                continue
            }

            foreach ($name in $allNames) {
                if ($cleanLine -match ("(?i)\b" + [regex]::Escape($name) + "\b")) {
                    $callSites.Add([PSCustomObject]@{
                        ProcedureName = $name
                        CallerModule  = $moduleName
                        LineNumber    = $lineNumber
                        LineText      = $cleanLine.Trim()
                    }) | Out-Null
                }
            }
        }
    }

    return $callSites
}

function Get-TargetModuleRecommendation {
    param([string[]]$procedureNames)

    $joined = ($procedureNames -join " ").ToLowerInvariant()

    if ($joined -match "tableexists|fieldexists|schema|indexexists|columnexists") { return "modDb" }
    if ($joined -match "translation|language|resolvetext") { return "modTranslationService" }
    if ($joined -match "path|folder|file") { return "modOutputPathService" }
    if ($joined -match "log|warning|info|error") { return "modLoggingHandler" }
    if ($joined -match "config|ini") { return "modConfigIni" }

    return "MANUAL_TARGET_SELECTION"
}

function Get-RiskLevel {
    param(
        [object[]]$groupProcedures,
        [string]$classification
    )

    $hasCurrentDbMix = (($groupProcedures | Where-Object UsesCurrentDb).Count -gt 0) -and (($groupProcedures | Where-Object UsesExplicitDaoDatabase).Count -gt 0)
    $hasPrivateState = ($groupProcedures | Where-Object { -not [string]::IsNullOrWhiteSpace($_.UsesPrivateModuleVariables) }).Count -gt 0
    $hasSideEffects = ($groupProcedures | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SideEffects) }).Count -gt 0

    if ($hasPrivateState -or $classification -eq "SAME_NAME_DIFFERENT") { return "HIGH" }
    if ($hasCurrentDbMix -or $hasSideEffects -or $classification -eq "SIMILAR") { return "MEDIUM" }
    return "LOW"
}

function Add-GroupRow {
    param(
        [System.Collections.Generic.List[object]]$groupRows,
        [string]$groupId,
        [string]$classification,
        [object[]]$procedures,
        [object[]]$callSites
    )

    if ($procedures.Count -lt 2) {
        return
    }

    $procedureNames = $procedures.ProcedureName | Sort-Object -Unique
    $sameName = $procedureNames.Count -eq 1
    $allPrivate = ($procedures | Where-Object Visibility -eq "Private").Count -eq $procedures.Count
    $callCount = 0
    $callerModules = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($procedure in $procedures) {
        $matchingCalls = @($callSites | Where-Object { $_.ProcedureName -eq $procedure.ProcedureName })
        $callCount += $matchingCalls.Count
        foreach ($call in $matchingCalls) {
            [void]$callerModules.Add($call.CallerModule)
        }
    }

    $functionalCommonalities = @()
    if (($procedures | Where-Object UsesCurrentDb).Count -gt 0) { $functionalCommonalities += "CurrentDb" }
    if (($procedures | Where-Object UsesExplicitDaoDatabase).Count -gt 0) { $functionalCommonalities += "DAO.Database" }
    if (($procedures | Where-Object HasErrorHandling).Count -gt 0) { $functionalCommonalities += "ErrorHandling" }
    if (($procedures | Where-Object HasLogging).Count -gt 0) { $functionalCommonalities += "Logging" }

    $differences = @()
    if (($procedures.Visibility | Sort-Object -Unique).Count -gt 1) { $differences += "Visibility" }
    if (($procedures.Signature | Sort-Object -Unique).Count -gt 1) { $differences += "Signature" }
    if (($procedures.UsesCurrentDb | Sort-Object -Unique).Count -gt 1) { $differences += "CurrentDb-vs-DAO" }
    if (($procedures.HasErrorHandling | Sort-Object -Unique).Count -gt 1) { $differences += "ErrorHandling" }
    if (($procedures.HasLogging | Sort-Object -Unique).Count -gt 1) { $differences += "Logging" }
    if (($procedures.SideEffects | Sort-Object -Unique).Count -gt 1) { $differences += "SideEffects" }

    $recommendation = switch ($classification) {
        "EXACT" {
            if ($allPrivate -and $callerModules.Count -le 1) { "KEEP_PRIVATE" }
            else { "MERGE_EXACT" }
        }
        "SIMILAR" {
            if ($allPrivate -and $callerModules.Count -le 1) { "KEEP_PRIVATE" }
            else { "MERGE_WITH_SIGNATURE_UNIFICATION" }
        }
        "SAME_NAME_DIFFERENT" { "MANUAL_REVIEW" }
        default { "KEEP_SEPARATE" }
    }

    $groupRows.Add([PSCustomObject]@{
        GroupId                    = $groupId
        Classification             = $classification
        ProcedureNames             = ($procedureNames -join "; ")
        Modules                    = (($procedures.ModuleName | Sort-Object -Unique) -join "; ")
        DefinitionCount            = $procedures.Count
        Visibility                 = (($procedures.Visibility | Sort-Object -Unique) -join "; ")
        Signatures                 = (($procedures.Signature | Sort-Object -Unique) -join " || ")
        FunctionalCommonalities    = ($functionalCommonalities -join "; ")
        RelevantDifferences        = ($differences -join "; ")
        CallSiteCount              = $callCount
        CallingModules             = ((@($callerModules) | ForEach-Object { $_ }) | Sort-Object -Unique) -join "; "
        ProposedCanonicalSignature = $procedures[0].Signature
        ProposedTargetModule       = Get-TargetModuleRecommendation $procedureNames
        RecommendedAction          = $recommendation
        RiskLevel                  = Get-RiskLevel -groupProcedures $procedures -classification $classification
    }) | Out-Null
}

$modules = Get-ChildItem -Path $moduleDir -Filter "*.bas" | Sort-Object Name
$allProcedures = New-Object System.Collections.Generic.List[object]

foreach ($module in $modules) {
    foreach ($procedure in (Get-ProcedureCatalog -modulePath $module.FullName)) {
        $allProcedures.Add($procedure) | Out-Null
    }
}

$callSites = Get-CallSites -procedures $allProcedures
$callSiteSummary = New-Object System.Collections.Generic.List[object]

foreach ($procedure in $allProcedures) {
    $calls = @($callSites | Where-Object { $_.ProcedureName -eq $procedure.ProcedureName })
    $callSiteSummary.Add([PSCustomObject]@{
        ModuleName     = $procedure.ModuleName
        ProcedureName  = $procedure.ProcedureName
        Signature      = $procedure.Signature
        CallSiteCount  = $calls.Count
        CallingModules = (($calls.CallerModule | Sort-Object -Unique) -join "; ")
    }) | Out-Null
}

$duplicateGroups = New-Object System.Collections.Generic.List[object]
$groupIndex = 1

$exactGroups = $allProcedures | Group-Object ExactBodySignature | Where-Object {
    -not [string]::IsNullOrWhiteSpace($_.Name) -and $_.Count -gt 1
}
foreach ($group in $exactGroups) {
    Add-GroupRow -groupRows $duplicateGroups -groupId ("DG-EXACT-" + $groupIndex.ToString("000")) -classification "EXACT" -procedures @($group.Group) -callSites $callSites
    $groupIndex++
}

$similarGroups = $allProcedures | Group-Object SkeletonSignature | Where-Object {
    -not [string]::IsNullOrWhiteSpace($_.Name) -and $_.Count -gt 1 -and
    (@($_.Group | Group-Object ExactBodySignature).Count -gt 1)
}
foreach ($group in $similarGroups) {
    Add-GroupRow -groupRows $duplicateGroups -groupId ("DG-SIM-" + $groupIndex.ToString("000")) -classification "SIMILAR" -procedures @($group.Group) -callSites $callSites
    $groupIndex++
}

$sameNameDifferentGroups = $allProcedures | Group-Object ProcedureName | Where-Object {
    $_.Count -gt 1 -and (@($_.Group | Group-Object ExactBodySignature).Count -gt 1)
}
foreach ($group in $sameNameDifferentGroups) {
    Add-GroupRow -groupRows $duplicateGroups -groupId ("DG-NAME-" + $groupIndex.ToString("000")) -classification "SAME_NAME_DIFFERENT" -procedures @($group.Group) -callSites $callSites
    $groupIndex++
}

$inventory = foreach ($procedure in $allProcedures) {
    $summary = $callSiteSummary | Where-Object {
        $_.ModuleName -eq $procedure.ModuleName -and $_.ProcedureName -eq $procedure.ProcedureName -and $_.Signature -eq $procedure.Signature
    } | Select-Object -First 1

    [PSCustomObject]@{
        ModuleName                 = $procedure.ModuleName
        ProcedureName              = $procedure.ProcedureName
        Kind                       = $procedure.Kind
        Visibility                 = $procedure.Visibility
        Signature                  = $procedure.Signature
        ReturnType                 = $procedure.ReturnType
        Parameters                 = $procedure.Parameters
        ParameterCount             = $procedure.ParameterCount
        StartLine                  = $procedure.StartLine
        EndLine                    = $procedure.EndLine
        CodeLines                  = $procedure.CodeLines
        CommentLines               = $procedure.CommentLines
        CallSiteCount              = if ($null -ne $summary) { $summary.CallSiteCount } else { 0 }
        CallingModules             = if ($null -ne $summary) { $summary.CallingModules } else { "" }
        UsesCurrentDb              = $procedure.UsesCurrentDb
        UsesExplicitDaoDatabase    = $procedure.UsesExplicitDaoDatabase
        UsesPrivateModuleVariables = $procedure.UsesPrivateModuleVariables
        HasErrorHandling           = $procedure.HasErrorHandling
        HasLogging                 = $procedure.HasLogging
        SideEffects                = $procedure.SideEffects
    }
}

$consolidationCandidates = foreach ($group in $duplicateGroups) {
    [PSCustomObject]@{
        GroupId                    = $group.GroupId
        Classification             = $group.Classification
        ProcedureNames             = $group.ProcedureNames
        Modules                    = $group.Modules
        DefinitionCount            = $group.DefinitionCount
        ProposedCanonicalSignature = $group.ProposedCanonicalSignature
        ProposedTargetModule       = $group.ProposedTargetModule
        RecommendedAction          = $group.RecommendedAction
        RiskLevel                  = $group.RiskLevel
    }
}

$analyzedModulesCount = $modules.Count
$analyzedProceduresCount = $allProcedures.Count
$privateCount = @($allProcedures | Where-Object Visibility -eq "Private").Count
$publicCount = @($allProcedures | Where-Object Visibility -eq "Public").Count
$exactCount = @($duplicateGroups | Where-Object Classification -eq "EXACT").Count
$similarCount = @($duplicateGroups | Where-Object Classification -eq "SIMILAR").Count
$sameNameDiffCount = @($duplicateGroups | Where-Object Classification -eq "SAME_NAME_DIFFERENT").Count
$centralPublicCandidates = @($duplicateGroups | Where-Object { $_.RecommendedAction -in @("MERGE_EXACT","MERGE_WITH_SIGNATURE_UNIFICATION") }).Count
$keepPrivateCount = @($duplicateGroups | Where-Object RecommendedAction -eq "KEEP_PRIVATE").Count
$highRiskCount = @($duplicateGroups | Where-Object RiskLevel -eq "HIGH").Count

$recommendedOrder = @(
    "1. Exact helper duplicates mit niedrigerem Risiko und ohne private Modulvariablen konsolidieren.",
    "2. Speziell TableExists/FieldExists in Richtung einer kanonischen Datenbank-API vorbereiten.",
    "3. Aehnliche Duplikate mit CurrentDb-vs-DAO-Unterschieden signaturseitig vereinheitlichen.",
    "4. Gleichnamige, aber unterschiedliche Prozeduren separat manuell pruefen, bevor Sichtbarkeiten veraendert werden.",
    "5. Kandidaten mit Modulzustand oder starken Seiteneffekten erst zuletzt anfassen."
)

$readmeLines = @(
    "# Standard Module Procedure Analysis",
    "",
    "## Summary",
    "- Analysed standard modules: $analyzedModulesCount",
    "- Analysed procedures: $analyzedProceduresCount",
    "- Private procedures: $privateCount",
    "- Public procedures: $publicCount",
    "- Exact duplicate groups: $exactCount",
    "- Similar duplicate groups: $similarCount",
    "- Same-name but different groups: $sameNameDiffCount",
    "- Candidates for a central Public procedure: $centralPublicCandidates",
    "- Procedures/groups that should remain Private: $keepPrivateCount",
    "- High-risk consolidation candidates: $highRiskCount",
    "",
    "## Notes",
    "- Scope includes only exported VBA standard modules under `src/access/exported/modules`.",
    "- Detection is heuristic and normalises comments, whitespace, literals and identifier names for similarity analysis.",
    "- No source modules were modified in this step.",
    "",
    "## Special focus",
    "- `TableExists` and `FieldExists` are explicitly included in the duplicate-group outputs and call-site inventory.",
    "",
    "## Recommended consolidation order"
)
$readmeLines += ($recommendedOrder | ForEach-Object { "- $_" })
$readme = $readmeLines -join "`r`n"

$inventory | Export-Csv -Path (Join-Path $outDir "module-procedure-inventory.csv") -NoTypeInformation -Encoding UTF8
$duplicateGroups | Export-Csv -Path (Join-Path $outDir "module-duplicate-procedure-groups.csv") -NoTypeInformation -Encoding UTF8
$callSites | Export-Csv -Path (Join-Path $outDir "module-procedure-call-sites.csv") -NoTypeInformation -Encoding UTF8
$consolidationCandidates | Export-Csv -Path (Join-Path $outDir "module-consolidation-candidates.csv") -NoTypeInformation -Encoding UTF8
Set-Content -Path (Join-Path $outDir "README.md") -Value $readme -Encoding UTF8

Write-Output ("MODULES=" + $analyzedModulesCount)
Write-Output ("PROCEDURES=" + $analyzedProceduresCount)
Write-Output ("EXACT_GROUPS=" + $exactCount)
Write-Output ("SIMILAR_GROUPS=" + $similarCount)
Write-Output ("SAME_NAME_DIFFERENT_GROUPS=" + $sameNameDiffCount)
Write-Output ("OUTDIR=" + $outDir)
