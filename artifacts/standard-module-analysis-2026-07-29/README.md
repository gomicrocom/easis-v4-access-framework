# Standard Module Procedure Analysis

## Summary
- Analysed standard modules: 59
- Analysed procedures: 1000
- Private procedures: 684
- Public procedures: 316
- Exact duplicate groups: 28
- Similar duplicate groups: 81
- Same-name but different groups: 45
- Candidates for a central Public procedure: 78
- Procedures/groups that should remain Private: 31
- High-risk consolidation candidates: 48

## Notes
- Scope includes only exported VBA standard modules under src/access/exported/modules.
- Detection is heuristic and normalises comments, whitespace, literals and identifier names for similarity analysis.
- No source modules were modified in this step.

## Special focus
- TableExists and FieldExists are explicitly included in the duplicate-group outputs and call-site inventory.

## Recommended consolidation order
- 1. Exact helper duplicates mit niedrigerem Risiko und ohne private Modulvariablen konsolidieren.
- 2. Speziell TableExists/FieldExists in Richtung einer kanonischen Datenbank-API vorbereiten.
- 3. Aehnliche Duplikate mit CurrentDb-vs-DAO-Unterschieden signaturseitig vereinheitlichen.
- 4. Gleichnamige, aber unterschiedliche Prozeduren separat manuell pruefen, bevor Sichtbarkeiten veraendert werden.
- 5. Kandidaten mit Modulzustand oder starken Seiteneffekten erst zuletzt anfassen.
