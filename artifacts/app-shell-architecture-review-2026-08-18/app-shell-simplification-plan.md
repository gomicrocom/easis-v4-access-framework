# Vereinfachungsplan

## Empfehlung

Schrittweise vereinfachen, kein Big Bang.

## Phase 1 – Verantwortungen schärfen
1. `CanLeaveWorkspace` auf reine Entscheidungslogik reduzieren.
2. Cleanup-/Delete-Flags aus generischen Leave-Pfaden entkoppeln.
3. Shell als reinen Dispatcher dokumentieren und absichern.

## Phase 2 – Kontexttransport vereinfachen
1. Für gehostete Formen genau einen Payload-Pfad festlegen.
2. Formularinterne Mehrfach-Fallbacks entfernen.
3. `frmOrderDetailNext` und ähnliche Formulare auf diesen einen Pfad harmonisieren.

## Phase 3 – Binding wieder nativer machen
1. Runtime-RecordSource-Manipulation kritisch prüfen.
2. Runtime-Subform-SourceObject-Umschaltung auf wenige Ausnahmefälle begrenzen.
3. Parent/Subform-Linking wo möglich rein nativ halten.

## Phase 4 – Lifecycle vereinfachen
1. `Form_Load` wieder zum eigentlichen Initialisierungspunkt machen, soweit Hosted-Kontext es zulässt.
2. `Form_Current` von Mischverantwortungen befreien.
3. unnötige `RefreshShellStatus`-Kaskaden reduzieren.

## Phase 5 – Problemformulare nachziehen
1. `frmOrderDetailNext`
2. `frmOrderDetail`
3. weitere Workspace-Detailformulare

## Keep / Simplify / Remove Zusammenfassung

- `Workspace SourceObject`: `KEEP`
- `Workspace History`: `KEEP`
- `Pending OpenArgs`: `KEEP_AND_SIMPLIFY`
- `ApplyWorkspaceOpenArgs`: `KEEP`
- `CanLeaveWorkspace`: `KEEP_AND_SIMPLIFY`
- `CommandBar Dispatch`: `KEEP`
- `Form Localization`: `KEEP`
- `Runtime RecordSource Manipulation`: `USE_ACCESS_NATIVE`
- `Runtime Event Wiring`: `REMOVE`
- `Form Mode Flags`: `KEEP_AND_SIMPLIFY`
- `Subform SourceObject Manipulation`: `KEEP_AND_SIMPLIFY`
- `Requery-Kaskaden`: `MANUAL_REVIEW`
- `Cleanup in Leave-Pfaden`: `REMOVE`

## Auswirkungen

Auf `frmOrderDetailNext`:
- weniger defensive Spezialpfade
- klarerer Session-Kontext
- geringere Gefahr von falschem Record/Alt-Datensatz

Auf bestehende Formulare:
- Listen-/einfache Detailformulare weitgehend stabil
- Translation, Logging, Repository/Service, TenantContext bleiben unberührt

Auf Shell/Navigation:
- keine funktionale Abwertung
- eher klarere Trennung und weniger Seiteneffekte
