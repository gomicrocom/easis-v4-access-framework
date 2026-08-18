# Zielarchitektur

## Empfehlung

Empfohlenes Sollmodell:
- ein klarer Order-Detail-Controller
- explizite Formularzustaende
- eine kanonische Berechnungslogik
- getrennte Verantwortung fuer:
  - Temp-Session
  - Persistierung
  - Zeilenberechnung
  - Header-Totale
  - Navigation

## 1. Explizite Formularzustaende

Nur diese drei fachlichen Hauptzustaende:

- `NEW_TEMPORARY`
- `EDIT_TEMPORARY`
- `EDIT_PERSISTED`

Ergaenzende technische Uebergangszustaende:

- `LOADING`
- `PERSISTING`
- `CANCELLING`
- `CLOSED`

## 2. Eine einzige Eintrittsstelle

Das Ziel-Formular soll genau eine Kontextinitialisierung besitzen:

- `InitializeOrderContext(contextPayload)`

Der Payload muss typisiert sein, zum Beispiel:

- `mode = TEMP`
- `tmp_order_id = 30`

oder

- `mode = PERSISTED`
- `order_id = 47`

Keine parallelen Pfade mehr ueber:
- `Me.OpenArgs`
- pending OpenArgs
- spaetere direkte `ApplyWorkspaceOpenArgs`
- neutrale Filter-Workarounds

## 3. Verantwortlichkeiten

### Formular

Nur:
- Controls binden
- sichtbaren Modus darstellen
- Benutzeraktionen entgegennehmen
- gezielt Serviceaufrufe delegieren

Nicht:
- doppelte Business-Berechnung
- temp/persisted Tabellenwechsel an vielen Stellen
- Navigation und Datenloeschung gleichzeitig entscheiden

### Temporary Order Service

Verantwortlich fuer:
- temp header anlegen
- temp lines anlegen
- temp defaults setzen
- temp cleanup
- temp -> persisted commit

### Calculation Service

Genau eine kanonische Zeilenberechnung.

Variante A:
- service rechnet direkt auf Datensatzebene

Variante B:
- service bekommt ein Line-DTO / Dictionary und gibt Betragsfelder zurueck

Wichtig:
- dieselbe Formel darf nicht gleichzeitig in Formular und Service gepflegt werden

### Totals Service

Genau eine kanonische Totals-Berechnung.

Regel:
- Totals immer auf derselben Datenbasis wie die aktuelle Zeile

## 4. Ereignisregeln

### Header laden

- bei `LOADING` nur Kontext binden
- erst danach Controls konfigurieren
- erst danach Subform aktivieren

### Neue Zeile

- nur moeglich, wenn der Header-Kontext stabil ist
- `line_no`, `vat_code`, `vat_rate` werden an genau einer Stelle gesetzt

### Zeilenfeld-Aenderung

Regel:
- Zeile neu berechnen
- aktuelle Zeile speichern oder kontrolliert puffern
- Header-Totale sofort mit derselben Datenbasis neu berechnen

### Apply / Refresh

Regel:
- nur Daten persistieren/rechnen
- keine Modusuebersetzung

### Save

Temp:
- commit temp -> persisted
- danach entweder:
  - auf persisted order offen bleiben
  - oder kontrolliert zurueck

Persisted:
- speichern
- definierte Navigation

### Cancel

Temp:
- temp cleanup erst nach bestaetigter, erfolgreicher Formularbeendigung

Persisted:
- alle ungespeicherten Aenderungen kontrolliert verwerfen
- auch Child-Changes muessen klare Semantik haben

## 5. Datenfluss Soll

### Neue Bestellung

Address -> temp_order session -> edit -> recalc -> persist commit -> persisted order

### Bestehende Bestellung

Persisted order -> edit -> recalc -> save

## 6. Temp-gegen-Persisted Strategie

Empfehlung:
- bestehende persisted orders weiterhin direkt auf `ord_order` bearbeiten
- neue Orders ueber temp session

Aber:
- diese beiden Modi sollen nicht dieselbe initiale Access-Eventstrecke mit wechselnder `RecordSource` durchlaufen

Praktisch bessere Varianten:

### Variante 1

Neues Entwicklungsformular fuer temp workflow
- `frmOrderDetailTemp`

und separates persisted:
- `frmOrderDetail`

### Variante 2

Ein neues gemeinsames Formular, aber mit:
- sauberem Controller
- erst nach Kontextwahl gebundener Datenquelle
- keinen spaeten Umschaltaktionen mehr

## 7. Child-Form Strategie

Empfehlung:
- kein dynamischer Kontextwechsel in demselben Child-Form mehr, wenn vermeidbar

Sauberere Optionen:
- `sfrmOrderLinesTmp`
- `sfrmOrderLines`

oder
- ein gemeinsames Child mit explizitem Init vor sichtbarer Aktivierung, aber ohne Default-Load in falscher Tabelle

## 8. Command-Semantik Soll

Einheitlich:

- `R3 = Uebernehmen`
- `R2 = Abbrechen`
- `R1 = Speichern`

Falls im Temp-Modus "Aktualisieren" benoetigt wird, dann nicht denselben Command-Key wie spaeteres Apply missbrauchen.

## 9. Logging Soll

Pro wesentlichem Schritt nur ein klarer Logpunkt:
- context initialized
- line recalculated
- totals recalculated
- temp persisted
- temp discarded
- navigation completed

Keine spaehten UI-/Translation-Nachlaeufer sollen fachliche Flows zerstoeren.
