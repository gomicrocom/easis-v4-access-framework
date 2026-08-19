# CanLeaveWorkspace Analyse

## Beobachtung

`modAppWorkspaceService.CanReplaceWorkspaceContent` ruft auf gehosteten Formularen direkt:

```text
CallByName(formObject, "CanLeaveWorkspace", VbMethod)
```

Damit ist `CanLeaveWorkspace` das zentrale Gate vor jedem Workspace-Wechsel.

## Positives Ziel

Die Funktion sollte nur beantworten:

```text
Darf das Formular jetzt verlassen werden?
```

## Problem im aktuellen Stand

Beispiel `frmOrderDetailNext`:
- prüft Dirty-Zustand
- zeigt Bestätigungsdialog
- setzt `m_deleteSessionOnClose = True`

Damit beantwortet die Funktion nicht nur eine Frage, sondern bereitet destruktives Verhalten vor.

Risiko:
- Navigation kann anschliessend dennoch scheitern
- der Cleanup-Wunsch bleibt am Formular hängen
- Leave-Entscheidung und Lebenszyklus-Cleanup sind gekoppelt

## Architekturbewertung

Gut:
- unsaved changes abfragen
- `False` zurückgeben, wenn Verlassen nicht erlaubt ist

Nicht gut:
- Löschen vorbereiten
- Persistierungsmodus ändern
- Lifecycle-State künstlich weiterdrehen

## Empfehlung

Kurzfristig:
- `CanLeaveWorkspace` nur lesen/prüfen lassen
- keine Daten löschen
- keine Session zur Löschung vormerken

Separate Lifecycle-Stufe:
- `BeforeWorkspaceLeaveConfirmed`
- oder Cleanup nur in expliziten Formularaktionen wie `DetailCancel`
- oder Cleanup beim tatsächlichen `Form_Unload`, aber nur wenn ein klar gesetzter, fachlicher Cancel-Pfad erfolgreich war

## Entscheidung

- `CanLeaveWorkspace`: `KEEP_AND_SIMPLIFY`
- destruktive Seiteneffekte darin: `REMOVE`
