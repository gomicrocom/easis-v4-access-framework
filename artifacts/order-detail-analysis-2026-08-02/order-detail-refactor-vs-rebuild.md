# Refactor vs Rebuild

## Entscheidung

Empfehlung: `REBUILD_CONTROLLED`

## 1. Bewertung des bestehenden Formulars

### Positiv wiederverwendbar

- Tabellen und Feldwelt `tmp_order`, `tmp_order_line`, `ord_order`, `ord_order_line`
- VAT-Defaultlogik in `modOrderRepository.GetDefaultVatContextForOrder()`
- Persistierungslogik in `modOrderRepository.PersistTemporaryOrder()`
- Totals-Service in `modOrderCalculationService`
- Workspace-Framework grundsaetzlich

### Problematisch im Formular

- mehrfacher Eintrittspfad in denselben Business-Kontext
- `RecordSource`-Wechsel zur Laufzeit
- Child-Form dynamisch zwischen zwei Tabellenwelten
- doppelte Zeilenberechnung
- harte `Me.Requery`-Spruenge
- temp cleanup im Leave-Pfad
- UI-Defaults in `Form_Current`

## 2. Refactoring des bestehenden Formulars

### Vorteile

- bestehendes Layout und Control-Namen koennen bleiben
- weniger initiale Neuverdrahtung

### Nachteile

- viele hoch gekoppelte Prozeduren muessen gleichzeitig umgebaut werden
- schwer sicher zu entflechten, weil Event-Reihenfolgen Access-spezifisch sind
- hohe Gefahr weiterer Regressionen in:
  - OpenArgs
  - History/GoBack
  - Child-Subform
  - temp cleanup
  - Totals

### Geschaetzter Eingriff

Gross:
- `frmOrderDetail` fast vollflaechig
- `sfrmOrderLines` substanziell
- wahrscheinlich auch `modAppWorkspaceService` Touchpoints

## 3. Kontrollierter Neuaufbau

### Vorteile

- explizite Zielarchitektur kann sauber implementiert werden
- alter Patchcode bleibt als Referenz bestehen
- echte Vergleichbarkeit im Test
- Zustandsmodell kann von Beginn an klar sein
- schrittweise Freigabe moeglich

### Nachteile

- zusaetzliches paralleles Entwicklungsobjekt fuer eine Uebergangszeit
- Formularlayout muss einmal gezielt nachgebaut oder abgeleitet werden

### Geschaetzter Eingriff

Mittel bis gross, aber planbarer:
- neues Entwicklungsformular fuer Detail
- ggf. neues Child-Form fuer temp lines
- bestehende Services weitgehend wiederverwendbar

## 4. Konkrete Empfehlung

### Warum nicht weiter patchen

Die Fehler liegen nicht in einer einzelnen Funktion, sondern im Zusammenspiel aus:
- Access Eventmodell
- Workspace Navigation
- temp/persisted Mischbetrieb
- lokaler vs. servicebasierter Berechnung

Weitere Punktfixes wuerden sehr wahrscheinlich:
- neue Randfehler erzeugen
- die Eventlogik noch undurchsichtiger machen
- die Wartbarkeit weiter verschlechtern

### Warum Rebuild sinnvoller ist

Weil grosse Teile des Werts bereits ausserhalb des Formulars existieren:
- Repository
- Calculation
- VAT-Defaults
- Workspace

Neu gebaut werden muss vor allem:
- der Form-Controller
- die Event-Semantik
- die Trennung der Modi

## 5. Migrationsplan fuer REBUILD_CONTROLLED

1. Bestehendes `frmOrderDetail` unveraendert als Referenz behalten
2. neues Entwicklungsformular anlegen, z.B. `frmOrderDetail_vNext`
3. neuen Controllerfluss definieren:
   - `InitializeOrderContext`
   - `LoadTemporaryOrder`
   - `LoadPersistedOrder`
4. Child-Strategie festlegen:
   - eigenes temp line subform oder stabiler gemeinsamer Child-Controller
5. zuerst nur Header laden und speichern
6. danach Zeilen anhaengen
7. danach lokale Zeilenberechnung und sofortige Totals
8. danach temp -> persisted commit
9. danach CommandBar und Navigation
10. erst nach Abnahme altes Formular ersetzen
11. erst danach alten Patchcode entfernen

## 6. Risiken

- Access-Eventreihenfolge bleibt auch im Neuaufbau kritisch
- alte Layout-Abhaengigkeiten koennen versteckte Control-Erwartungen enthalten
- Persistierung und History muessen sauber voneinander getrennt werden

Diese Risiken sind bei einem kontrollierten Neuaufbau aber deutlich besser beherrschbar als bei einem weiteren In-Place-Patching.

## 7. Geschaetzter Umsetzungsumfang

Technischer Umfang:
- mittel bis hoch

Praktisch:
- 1 Schritt fuer Header-Flow
- 1 Schritt fuer Child-Flow
- 1 Schritt fuer Totals/Persistierung
- 1 Schritt fuer Navigation/CommandBar

Das ist groesser als ein lokaler Fix, aber kleiner und sicherer als fortgesetztes Patchen an der jetzigen Struktur.
