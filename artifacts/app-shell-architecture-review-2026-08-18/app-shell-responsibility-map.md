# AppShell Responsibility Map

## frmAppShell

Aktuelle Verantwortungen:
- globaler visueller Host
- Subform-Container für Navigation
- Subform-Container für Workspace
- Shell-Statusanzeige
- zentrale CommandBar-Buttons
- globale QuickSearch-Weiterleitung
- Start der Shell-Initialisierung
- Unload-Cleanup für Workspace

Bewertung:
- Gehört sinnvoll in die Shell:
  - Navigation hosten
  - Workspace hosten
  - globale Statusanzeige
  - CommandBar-UI
  - Delegation von Shell-Buttons
- Gehört besser ins eingebettete Formular:
  - Datensatzmodus
  - RecordSource
  - Dirty-/Undo-Entscheidung
  - Zeilen- oder Headerberechnung
  - fachliches Speichern / Verwerfen
- Gehört in Services:
  - Navigation-History
  - Datenbank-/Tenant-Kontext
  - fachliche Persistierung
- Wird bereits nativ von Access gelöst:
  - Formular-Lifecycle
  - Record navigation
  - `OpenArgs` bei normalem `DoCmd.OpenForm`
  - Parent/Subform-Linking

## modAppShell

Sinnvolle Verantwortungen:
- Shell initialisieren
- globale Statuswerte aufbereiten
- zentrale CommandBar-Konfiguration rendern
- Workspace-Form ermitteln
- Kommandos an aktive Form delegieren

Zu tief in UI/Form intern:
- teilweise implizite Annahmen über Workspace-Form-API
- zyklisches `RefreshShellStatus`
- Shell-Logik kennt eine breite Menge formularspezifischer Command-Konventionen

Bewertung:
- `KEEP_AND_SIMPLIFY`

## modAppWorkspaceService

Sinnvolle Verantwortungen:
- Host-Control für Workspace ansprechen
- Formwechsel im Workspace orchestrieren
- History verwalten
- Fehler-Recovery im Host

Problematische Verantwortungen:
- eigener Kontexttransport zusätzlich zu Access
- nachträgliche `ApplyWorkspaceOpenArgs`
- Runtime-Filter-/RecordSource-Manipulation
- `DoCmd.GoToRecord acNewRec` aus Host-Service
- Restore von Formularzustand über generische History-Serialisierung

Bewertung:
- `KEEP_AND_SIMPLIFY`

## modAppNavigationService

Sinnvolle Verantwortungen:
- Navigationstabellen sicherstellen
- Seed der Shell-Navigation
- Navigation-Klicks in Form-/Report-Wechsel übersetzen

Potentielle Übergriffe:
- Navigation-Seed beim Shell-Start
- direkte Tabellen-DDL/Seed-Aktionen während UI-Initialisierung

Bewertung:
- `KEEP`, aber Setup/Bootstrap zeitlich klarer entkoppeln

## Order-/Detailformulare

Sinnvoll im Formular:
- gebundene RecordSource
- Current/Dirty/NewRecord
- Validierung
- Parent/Subform-Verhalten
- fachliche Buttonsemantik

Aktuelle Spannungen:
- Workspace-Kontext kommt teils erst nach `Form_Load`
- Formulare führen leere RecordSource + nachträgliche Session-Bindung ein
- Mode-/Lifecycle-Flags parallel zu nativen Zuständen

Bewertung:
- Formulare selbst bleiben sinnvoll
- aktuelle Workspace-Anpassungen rund um `frmOrderDetail` / `frmOrderDetailNext` sind überkompensiert
