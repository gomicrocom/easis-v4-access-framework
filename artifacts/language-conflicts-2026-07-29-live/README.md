# Language-code conflict analysis

## Result
- Real fw_translation conflicts with different target row: 0
- Logged pseudo-conflicts / self-matches: 1092
- Auto-cleanable identical cases: 1092
- Different-text conflicts: 0

## Explanation
- Access evaluates the DCount criteria case-insensitively.
- A row like EN-US already matches the search for en-US and is therefore counted as its own conflict.
- The logged conflict_count=1092 exactly matches the number of rows whose language_code would normalize to a different canonical spelling.

## Files
- fw_translation_language_code_self_matches.csv
- fw_translation_conflicts_identical.csv (real duplicates; currently empty)
- fw_translation_conflicts_different.csv (real text conflicts; currently empty)
