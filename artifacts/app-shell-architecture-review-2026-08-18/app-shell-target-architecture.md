# Zielarchitektur

## frmAppShell

Verantwortlich für:
- globale Navigation hosten
- Workspace-Container hosten
- globale CommandBar darstellen
- globalen Status / Tenant / User anzeigen
- übergreifende Navigation (Home/Back)

Nicht verantwortlich für:
- Datensatzselektion eines Detailformulars
- RecordSource eines Business-Forms
- Dirty-/Undo-/CurrentRecord-Zustand
- Zeilen-/Headerberechnung
- fachliches Save/Cancel

## Workspace-Service

Verantwortlich für:
- Workspace-Host-Control bedienen
- Formwechsel innerhalb des Host-Containers
- History
- technische Kontextweitergabe an gehostete Formen
- Fehler-Recovery des Hostings

Nicht verantwortlich für:
- fachliche Moduslogik
- Berechnung
- Cleanup von Fachdaten
- Record-Navigation innerhalb des Formulars

## Formular

Verantwortlich für:
- gebundene RecordSource
- Filter / Current Record
- NewRecord / Dirty / Undo
- Validierung
- Parent/Subform-Bindung
- lokale Buttonsemantik
- Delegation fachlicher Aktionen an Services/Repositories

## Service

Verantwortlich für:
- Geschäftslogik
- Berechnungen
- mehrstufige Persistierung
- tenant-/datenbankübergreifende Operationen

## Repository

Verantwortlich für:
- expliziten Datenzugriff
- kein UI-Zustand
- keine Formularnavigation

## OpenArgs / Kontexttransport

Bevorzugtes Ziel:
- Standalone: natives `OpenArgs` / `WhereCondition`
- Hosted: genau ein technischer Payload-Mechanismus
- Formulare nicht gleichzeitig aus:
  - `Me.OpenArgs`
  - Pending Payload
  - aktuellem Datensatz
  - History
  ableiten lassen

## Parent / Subform

Bevorzugtes Ziel:
- `LinkMasterFields` / `LinkChildFields`
- gebundene Child-RecordSource
- keine dynamische Umschaltung der Child-Quelle ausser bei bewusstem Spezialfall

## Cleanup

Bevorzugtes Ziel:
- `CanLeaveWorkspace` prüft nur
- Formularaktion `Cancel` entscheidet fachlich über Verwerfen
- tatsächliches Cleanup erst in klarer bestätigter Lifecycle-Stufe
