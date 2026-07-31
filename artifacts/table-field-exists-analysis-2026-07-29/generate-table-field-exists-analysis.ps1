$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$moduleDir = Join-Path $repoRoot "src\access\exported\modules"
$outDir = $PSScriptRoot
$targetNames = @("TableExists", "FieldExists")

function Remove-InlineComment {
    param([string]$line)

    $inString = $false
    $i = 0
    while ($i -lt $line.Length) {
        $ch = $line[$i]
        if ($ch -eq '"') {
            if ($inString -and $i + 1 -lt $line.Length -and $line[$i + 1] -eq '"') {
                $i += 2
                continue
            }
            $inString = -not $inString
            $i++
            continue
        }
        if (-not $inString -and $ch -eq "'") {
            return $line.Substring(0, $i)
        }
        $i++
    }
    return $line
}

function Normalize-Body {
    param([string]$body)

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($rawLine in ($body -split "`r?`n")) {
        $clean = (Remove-InlineComment $rawLine).Trim()
        if (-not [string]::IsNullOrWhiteSpace($clean)) {
            $lines.Add(([regex]::Replace($clean.ToLowerInvariant(), "\s+", " "))) | Out-Null
        }
    }
    return ($lines -join "`n")
}

function Get-BodyHash {
    param([string]$body)
    $sha = [System.Security.Cryptography.SHA1]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes((Normalize-Body $body))
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-ReturnBehavior {
    param([string]$body, [string]$procedureName)
    $clean = Normalize-Body $body
    if ($clean -match ("\b" + [regex]::Escape($procedureName.ToLowerInvariant()) + "\s*=\s*true\b")) {
        if ($clean -match ("\b" + [regex]::Escape($procedureName.ToLowerInvariant()) + "\s*=\s*false\b")) {
            return "Assigns True/False explicitly"
        }
        return "Assigns True explicitly"
    }
    if ($clean -match ("\b" + [regex]::Escape($procedureName.ToLowerInvariant()) + "\s*=\s*false\b")) {
        return "Assigns False explicitly"
    }
    if ($clean -match "\bexit function\b") {
        return "Implicit default False with early Exit Function"
    }
    return "Implicit default return"
}

function Get-EmptyNameHandling {
    param([string]$body)
    $clean = Normalize-Body $body
    if ($clean -match "lenb?\s*\(\s*trim\$\(" -or $clean -match "\blen\s*\(") {
        if ($clean -match "\bexit function\b") {
            return "Checks empty/trimmed input and exits early"
        }
        return "Checks empty/trimmed input"
    }
    return "No explicit empty-name guard"
}

function Get-MissingBehavior {
    param([string]$body)
    $clean = Normalize-Body $body
    if ($clean -match "for each .* in db\.tabledefs") {
        return "Iterates DAO TableDefs; returns False when not found"
    }
    if ($clean -match "tabledefs" -and $clean -match "fields") {
        return "Traverses TableDefs/Fields; returns False when not found"
    }
    return "Manual review"
}

function Parse-Definitions {
    param([string]$modulePath)

    $moduleName = [IO.Path]::GetFileNameWithoutExtension($modulePath)
    $lines = Get-Content -Path $modulePath -Encoding UTF8
    $results = New-Object System.Collections.Generic.List[object]

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line -match "^\s*(Public|Private)\s+Function\s+(TableExists|FieldExists)\b") {
            $visibility = $matches[1]
            $procedureName = $matches[2]
            $startLine = $i + 1
            $bodyLines = New-Object System.Collections.Generic.List[string]
            $bodyLines.Add($line) | Out-Null
            $j = $i + 1
            while ($j -lt $lines.Count) {
                $bodyLines.Add($lines[$j]) | Out-Null
                if ($lines[$j] -match "^\s*End\s+Function\s*$") {
                    break
                }
                $j++
            }
            $endLine = $j + 1
            $body = $bodyLines -join "`r`n"
            $signature = ([regex]::Replace($line.Trim(), "\s+", " "))
            $normalizedSignature = $signature.ToLowerInvariant()
            $bodyHash = Get-BodyHash $body
            $results.Add([PSCustomObject]@{
                procedure_name = $procedureName
                module_name = $moduleName
                module_path = $modulePath
                start_line = $startLine
                end_line = $endLine
                visibility = $visibility
                signature = $signature
                body = $body
                uses_currentdb = ($body -match "\bCurrentDb\b")
                uses_explicit_dao_database = ($signature -match "\bByVal\s+db\s+As\s+DAO\.Database\b")
                has_error_handling = ($body -match "^\s*On\s+Error\b" -or $body -match "\bResume\b")
                return_behavior = Get-ReturnBehavior -body $body -procedureName $procedureName
                empty_name_handling = Get-EmptyNameHandling -body $body
                missing_behavior = Get-MissingBehavior -body $body
                body_hash = $bodyHash
                signature_key = $normalizedSignature
                exact_group_key = ($procedureName.ToLowerInvariant() + "|" + $normalizedSignature + "|" + $bodyHash)
            }) | Out-Null
            $i = $j
        }
    }

    return $results
}

function Is-DefinitionLine {
    param([string]$line)
    return ($line -match "^\s*(Public|Private)\s+Function\s+(TableExists|FieldExists)\b")
}

function Is-DeclarationLike {
    param([string]$line)
    return ($line -match "^\s*(Dim|Const|Private|Public|Static|Function|Sub|Property|Type|Enum)\b")
}

function Is-AssignmentToFunction {
    param([string]$line, [string]$name)
    return ($line -match ("^\s*" + [regex]::Escape($name) + "\s*="))
}

function Extract-Arguments {
    param(
        [string]$line,
        [int]$startIndex
    )

    $openIndex = $line.IndexOf("(", $startIndex)
    if ($openIndex -lt 0) { return "" }

    $depth = 0
    $inString = $false
    for ($i = $openIndex; $i -lt $line.Length; $i++) {
        $ch = $line[$i]
        if ($ch -eq '"') {
            if ($inString -and $i + 1 -lt $line.Length -and $line[$i + 1] -eq '"') {
                $i++
                continue
            }
            $inString = -not $inString
            continue
        }
        if (-not $inString) {
            if ($ch -eq "(") { $depth++ }
            elseif ($ch -eq ")") {
                $depth--
                if ($depth -eq 0) {
                    return $line.Substring($openIndex + 1, $i - $openIndex - 1).Trim()
                }
            }
        }
    }
    return ""
}

function Find-CallSites {
    param([object[]]$definitions)

    $defsByModule = @{}
    foreach ($definition in $definitions) {
        if (-not $defsByModule.ContainsKey($definition.module_name)) {
            $defsByModule[$definition.module_name] = @{}
        }
        $defsByModule[$definition.module_name][$definition.procedure_name.ToLowerInvariant()] = $definition
    }

    $callSites = New-Object System.Collections.Generic.List[object]

    Get-ChildItem -Path $moduleDir -Filter "*.bas" | Sort-Object Name | ForEach-Object {
        $modulePath = $_.FullName
        $moduleName = $_.BaseName
        $lines = Get-Content -Path $modulePath -Encoding UTF8
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $lineNo = $i + 1
            $cleanLine = Remove-InlineComment $lines[$i]
            if ([string]::IsNullOrWhiteSpace($cleanLine)) { continue }
            if (Is-DefinitionLine $cleanLine) { continue }
            if (Is-DeclarationLike $cleanLine) { continue }

            foreach ($procName in $targetNames) {
                if (Is-AssignmentToFunction -line $cleanLine -name $procName) { continue }

                $qualifiedPattern = "\b([A-Za-z_][A-Za-z0-9_]*)\." + [regex]::Escape($procName) + "\s*\("
                $qualifiedMatch = [regex]::Match($cleanLine, $qualifiedPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($qualifiedMatch.Success) {
                    $targetModule = $qualifiedMatch.Groups[1].Value
                    if ($defsByModule.ContainsKey($targetModule) -and $defsByModule[$targetModule].ContainsKey($procName.ToLowerInvariant())) {
                        $targetDef = $defsByModule[$targetModule][$procName.ToLowerInvariant()]
                        $args = Extract-Arguments -line $cleanLine -startIndex $qualifiedMatch.Index
                        $callSites.Add([PSCustomObject]@{
                            procedure_name = $procName
                            target_module = $targetDef.module_name
                            target_signature = $targetDef.signature
                            caller_module = $moduleName
                            line_number = $lineNo
                            line_text = $cleanLine.Trim()
                            call_style = "qualified"
                            arguments = $args
                            needs_explicit_db = $targetDef.uses_explicit_dao_database
                            uses_currentdb_context = ($args -match "\bCurrentDb\b")
                        }) | Out-Null
                    }
                    continue
                }

                $unqualifiedPattern = "\b" + [regex]::Escape($procName) + "\s*\("
                $unqualifiedMatch = [regex]::Match($cleanLine, $unqualifiedPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($unqualifiedMatch.Success) {
                    if ($defsByModule.ContainsKey($moduleName) -and $defsByModule[$moduleName].ContainsKey($procName.ToLowerInvariant())) {
                        $targetDef = $defsByModule[$moduleName][$procName.ToLowerInvariant()]
                        $args = Extract-Arguments -line $cleanLine -startIndex $unqualifiedMatch.Index
                        $callSites.Add([PSCustomObject]@{
                            procedure_name = $procName
                            target_module = $targetDef.module_name
                            target_signature = $targetDef.signature
                            caller_module = $moduleName
                            line_number = $lineNo
                            line_text = $cleanLine.Trim()
                            call_style = "unqualified-local"
                            arguments = $args
                            needs_explicit_db = $targetDef.uses_explicit_dao_database
                            uses_currentdb_context = ($args -match "\bCurrentDb\b")
                        }) | Out-Null
                    }
                }
            }
        }
    }

    return $callSites
}

$definitions = New-Object System.Collections.Generic.List[object]
Get-ChildItem -Path $moduleDir -Filter "*.bas" | Sort-Object Name | ForEach-Object {
    foreach ($definition in (Parse-Definitions -modulePath $_.FullName)) {
        $definitions.Add($definition) | Out-Null
    }
}

$callSites = Find-CallSites -definitions $definitions

$definitionRows = New-Object System.Collections.Generic.List[object]
foreach ($definition in $definitions) {
    $callsForDefinition = @($callSites | Where-Object {
        $_.target_module -eq $definition.module_name -and
        $_.target_signature -eq $definition.signature -and
        $_.procedure_name -eq $definition.procedure_name
    })

    $definitionRows.Add([PSCustomObject]@{
        procedure_name = $definition.procedure_name
        module_name = $definition.module_name
        start_line = $definition.start_line
        end_line = $definition.end_line
        visibility = $definition.visibility
        signature = $definition.signature
        uses_currentdb = $definition.uses_currentdb
        uses_explicit_dao_database = $definition.uses_explicit_dao_database
        has_error_handling = $definition.has_error_handling
        return_behavior = $definition.return_behavior
        empty_name_handling = $definition.empty_name_handling
        missing_behavior = $definition.missing_behavior
        actual_call_count = $callsForDefinition.Count
        body = $definition.body
    }) | Out-Null
}

$differenceRows = New-Object System.Collections.Generic.List[object]
foreach ($procedureName in $targetNames) {
    $procDefs = @($definitions | Where-Object procedure_name -eq $procedureName)
    $exactGroups = $procDefs | Group-Object exact_group_key
    $groupIndex = 1
    foreach ($group in $exactGroups) {
        if ($group.Count -gt 1) {
            $differenceRows.Add([PSCustomObject]@{
                group_id = "$procedureName-EXACT-{0:d3}" -f $groupIndex
                procedure_name = $procedureName
                group_type = "EXACT"
                definition_count = $group.Count
                modules = (($group.Group.module_name | Sort-Object -Unique) -join "; ")
                signatures = (($group.Group.signature | Sort-Object -Unique) -join " || ")
                uses_currentdb = (($group.Group.uses_currentdb | Sort-Object -Unique) -join "; ")
                uses_explicit_dao_database = (($group.Group.uses_explicit_dao_database | Sort-Object -Unique) -join "; ")
                has_error_handling = (($group.Group.has_error_handling | Sort-Object -Unique) -join "; ")
                return_behavior = (($group.Group.return_behavior | Sort-Object -Unique) -join " || ")
                empty_name_handling = (($group.Group.empty_name_handling | Sort-Object -Unique) -join " || ")
                missing_behavior = (($group.Group.missing_behavior | Sort-Object -Unique) -join " || ")
            }) | Out-Null
            $groupIndex++
        }
    }

    $signatureGroups = $procDefs | Group-Object signature_key
    $signatureIndex = 1
    foreach ($group in $signatureGroups) {
        $hashes = @($group.Group.body_hash | Sort-Object -Unique)
        if ($hashes.Count -gt 1) {
            $differenceRows.Add([PSCustomObject]@{
                group_id = "$procedureName-SAME-SIGNATURE-DIFF-{0:d3}" -f $signatureIndex
                procedure_name = $procedureName
                group_type = "SAME_SIGNATURE_DIFFERENT_BODY"
                definition_count = $group.Count
                modules = (($group.Group.module_name | Sort-Object -Unique) -join "; ")
                signatures = $group.Group[0].signature
                uses_currentdb = (($group.Group.uses_currentdb | Sort-Object -Unique) -join "; ")
                uses_explicit_dao_database = (($group.Group.uses_explicit_dao_database | Sort-Object -Unique) -join "; ")
                has_error_handling = (($group.Group.has_error_handling | Sort-Object -Unique) -join "; ")
                return_behavior = (($group.Group.return_behavior | Sort-Object -Unique) -join " || ")
                empty_name_handling = (($group.Group.empty_name_handling | Sort-Object -Unique) -join " || ")
                missing_behavior = (($group.Group.missing_behavior | Sort-Object -Unique) -join " || ")
            }) | Out-Null
            $signatureIndex++
        }
    }
}

$tableDefs = @($definitions | Where-Object procedure_name -eq "TableExists")
$fieldDefs = @($definitions | Where-Object procedure_name -eq "FieldExists")

$callSummaryLines = foreach ($definition in $definitions) {
    $callsForDefinition = @($callSites | Where-Object {
        $_.target_module -eq $definition.module_name -and
        $_.target_signature -eq $definition.signature -and
        $_.procedure_name -eq $definition.procedure_name
    })
    "- $($definition.procedure_name) | $($definition.module_name):$($definition.start_line)-$($definition.end_line) | $($callsForDefinition.Count) call(s)"
}

$groupLines = foreach ($row in $differenceRows) {
    "- $($row.group_id): $($row.procedure_name) | $($row.group_type) | $($row.definition_count) definition(s) | $($row.modules)"
}

$removalCandidates = foreach ($definition in $definitions) {
    "- $($definition.module_name):$($definition.start_line)-$($definition.end_line) $($definition.signature)"
}

$changeCandidates = foreach ($definition in $definitions) {
    if ($definition.uses_explicit_dao_database) {
        "- Module $($definition.module_name): local calls to $($definition.procedure_name) should later be redirected to a shared schema helper with explicit DAO.Database."
    }
    else {
        "- Module $($definition.module_name): local calls to $($definition.procedure_name) should later be redirected either through explicit CurrentDb resolution or an agreed wrapper policy."
    }
}

$planLines = @(
    "# TableExists / FieldExists Refactoring Plan",
    "",
    "## Counts",
    "- Actual TableExists definitions: $($tableDefs.Count)",
    "- Actual FieldExists definitions: $($fieldDefs.Count)",
    "",
    "## Actual call counts per definition"
) + @($callSummaryLines) + @(
    "",
    "## Global identical implementation groups"
) + @($groupLines) + @(
    "",
    "## Functional differences",
    "- Two main API families exist for each helper: CurrentDb-based and explicit DAO.Database-based.",
    "- The explicit DAO.Database variants are the better base for a later shared schema API because they avoid hidden ambient database state.",
    "- The CurrentDb variants are simpler for repository callers but couple the helper to the ambient frontend context.",
    "- Parameter naming differences (tableName vs table_name, fieldName vs field_name) are cosmetic and should not drive separate long-term APIs.",
    "",
    "## Recommended public signatures",
    "- Preferred: Public Function TableExists(ByVal tableName As String, ByVal db As DAO.Database) As Boolean",
    "- Preferred: Public Function FieldExists(ByVal tableName As String, ByVal fieldName As String, ByVal db As DAO.Database) As Boolean",
    "- Assessment: an Optional DAO.Database parameter is technically possible in VBA only via an object/variant pattern, but it is less explicit and less readable than a required DAO.Database argument.",
    "- Recommended CurrentDb caller pattern: caller resolves CurrentDb explicitly and passes db into the canonical helper.",
    "",
    "## Recommended target module",
    "- Preferred target module: modDbSchema",
    "- Reason: both helpers inspect schema metadata and belong beside other field/index/table inspection responsibilities rather than in a broad generic db-access module.",
    "",
    "## Later removal candidates"
) + @($removalCandidates) + @(
    "",
    "## Later call-site changes"
) + @(@($changeCandidates | Sort-Object -Unique)) + @(
    "",
    "## Risks",
    "- Private local helpers currently own unqualified calls inside their own modules; future migration must retarget those calls deliberately.",
    "- CurrentDb-based and explicit-db-based call sites should not be merged blindly without checking backend-routing and transaction expectations.",
    "- Any shared helper must preserve current False-on-missing behavior and existing early-exit/error-handling assumptions.",
    "- Schema-sensitive modules that work against backend databases should keep explicit DAO.Database flow to avoid regressions."
) 

$definitionRows | Export-Csv -Path (Join-Path $outDir "table-field-exists-definitions.csv") -NoTypeInformation -Encoding UTF8
$callSites | Export-Csv -Path (Join-Path $outDir "table-field-exists-call-sites.csv") -NoTypeInformation -Encoding UTF8
$differenceRows | Export-Csv -Path (Join-Path $outDir "table-field-exists-implementation-differences.csv") -NoTypeInformation -Encoding UTF8
Set-Content -Path (Join-Path $outDir "table-field-exists-refactoring-plan.md") -Value ($planLines -join "`r`n") -Encoding UTF8

Write-Output ("TABLE_DEFINITIONS=" + $tableDefs.Count)
Write-Output ("FIELD_DEFINITIONS=" + $fieldDefs.Count)
Write-Output ("CALLS=" + $callSites.Count)
Write-Output ("OUTDIR=" + $outDir)
