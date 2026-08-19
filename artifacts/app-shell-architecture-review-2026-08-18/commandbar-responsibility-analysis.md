# CommandBar Responsibility Analysis

## Ist-Zustand

Die Shell besitzt eine globale CommandBar und delegiert Slots an das aktive Workspace-Formular.

Aktueller Ablauf:
- `frmAppShell.cmdL*/cmdR*`
- `modAppShell.ExecuteWorkspaceCommandBarSlot`
- `ResolveSlotCommandKey`
- `ExecuteWorkspaceCommandByKey`
- aktives Formular erhält `ExecuteWorkspaceCommand`

## Positiver Befund

Die Shell speichert in der Regel keine fachlichen Daten selbst.
Sie kennt hauptsächlich:
- Slot
- Command Key
- aktives Workspace-Formular

Das ist architektonisch sinnvoll.

## Problematische Tendenzen

Die Shell kennt inzwischen viele fachlich wirkende Standardkommandos:
- `DETAIL_APPLY`
- `DETAIL_SAVE`
- `DETAIL_CANCEL`
- `LIST_SEARCH`
- `LIST_EDIT`
- `ADDRESS_COCKPIT`

Das ist noch akzeptabel, solange:
- die Shell nur dispatcht
- das Formular selbst entscheidet, ob der Befehl erlaubt ist
- das Formular selbst Dirty-/Datensatz-/Persistierungszustand kennt

## Zielbild

Die Shell kennt:
- Navigation
- Sichtbarkeit / Beschriftung der Slots
- Delegation

Das Formular kennt:
- Datensatz
- Neu-/Edit-Zustand
- Apply/Save/Cancel-Semantik
- Validierung
- Cleanup

## Bewertung pro Mechanismus

- zentrale CommandBar: `KEEP`
- Slot->Command-Key-Mapping: `KEEP`
- `CanExecuteWorkspaceCommand` je Formular: `KEEP`
- Shell rekonstruiert Formularzustand: `REMOVE`
- Shell triggert fachliches Cleanup: `REMOVE`
