# Workspace Runtime Flow

## Aktueller Ablauf bei Workspace-Formwechsel

```text
User Action
-> Shell / Navigation button
-> modAppShell.ExecuteWorkspaceCommandBarSlot or modAppNavigationService.HandleNavigationClick
-> modAppWorkspaceService.OpenWorkspaceForm
-> CanReplaceWorkspaceContent
-> optional CaptureCurrentWorkspaceHistory
-> SetPendingWorkspaceOpenArgs(form_name, open_args)
-> workspaceHost.SourceObject = ""
-> workspaceHost.SourceObject = "Form.<Target>"
-> Access loads hosted form
   -> Form_Open (if implemented)
   -> Form_Load
   -> Form_Current
-> modAppWorkspaceService.ApplyWorkspaceFormState
-> modAppWorkspaceService.ApplyWorkspaceOpenArgs
-> host refresh / command bar refresh
```

## Kritischer Architekturpunkt

Bei gehosteten Formularen kommt der fachliche Kontext nicht zwingend in `Form_Load` an:
- echtes Access-`OpenArgs` ist beim `SourceObject`-Hosting nicht verlässlich
- deshalb existiert `m_pendingWorkspaceOpenArgs`
- das Formular lädt zunächst ohne finalen Kontext
- danach wird `ApplyWorkspaceOpenArgs` manuell aufgerufen

Folgen:
- Formulare mit RecordSource-/Session-Abhängigkeit müssen defensiv leer starten
- zusätzliche Initialisierungsflags werden nötig
- `Form_Load` kann nicht mehr vollständig als nativer Initialisierungspunkt dienen

## Vergleich: Standalone OpenForm

```text
DoCmd.OpenForm form, ..., where_condition, ..., open_args
-> Access erstellt Form
-> OpenArgs ist direkt im Formular verfügbar
-> WhereCondition wirkt beim Öffnen
-> Form_Load kann finalen Kontext verwenden
```

## History Restore

`GoBack`:
- liest serialisierten History-Eintrag
- ruft erneut `OpenWorkspaceForm`
- danach `RestoreWorkspaceHistoryState`

Wiederhergestellte Aspekte:
- Formname
- Filter
- OrderBy
- OpenArgs
- DataMode
- optional benutzerdefinierter `GetWorkspaceState` / `RestoreWorkspaceState`

Bewertung:
- funktional wertvoll
- aber technisch eng an generische Form-APIs gekoppelt
- sollte Navigation wiederherstellen, nicht fachliche Objektzustände rekonstruieren
