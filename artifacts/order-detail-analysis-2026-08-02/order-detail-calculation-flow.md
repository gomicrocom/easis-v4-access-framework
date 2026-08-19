# Berechnungsfluss

## 1. Zeilenberechnung im Ist-Zustand

### Ausloeser

In `sfrmOrderLines` loesen folgende Events eine lokale Berechnung aus:
- `txtQuantity_AfterUpdate`
- `txtUnitPrice_AfterUpdate`
- `txtDiscountValue_AfterUpdate`
- `cboDiscountType_AfterUpdate`
- `txtSurchargeValue_AfterUpdate`
- `cboSurchargeType_AfterUpdate`
- `cboVatCode_AfterUpdate`

Spezialfall:
- `cboVatCode_AfterUpdate` setzt vorher `vat_rate` aus Combo-Spalte 3

### Verwendete Methode

- `Form_sfrmOrderLines.CalcLineTotals()` (Zeilen 308-396)

### Was gelesen wird

Direkt aus aktuellen Controls:
- `txtQuantity`
- `txtUnitPrice`
- `txtVatRate`
- `cboDiscountType`
- `txtDiscountValue`
- `cboSurchargeType`
- `txtSurchargeValue`

Zusatz:
- `vat_mode` ueber Parent `frmOrderDetail`

### Was geschrieben wird

Direkt in Zeilenfelder bzw. gebundene Controls:
- `line_base_amount`
- `line_discount_amount`
- `line_surcharge_amount`
- `line_net_amount`
- `line_vat_amount`
- `line_gross_amount`

### Was bewusst nicht passiert

- kein `SaveRecord`
- kein `Me.Requery`
- kein `Parent.Requery`
- keine Header-Summenberechnung
- keine Repository- oder Schema-Aufrufe

## 2. Servicebasierte Zeilenberechnung

Parallel existiert in `modOrderCalculationService` dieselbe fachliche Verantwortung:

- `CalculateOrderLineAmounts`
- `CalculateTemporaryOrderLineAmounts`
- `CalculateOrderLineAmountsByOrderAndLineNo`
- `CalculateTemporaryOrderLineAmountsByOrderAndLineNo`
- intern `CalculateLineAmountsForContext()`

### Unterschied zur Formularlogik

Service:
- liest aus Recordsets
- schreibt in Datenbankfelder
- benoetigt bereits gespeicherte Zeile mit stabiler ID

Formular:
- liest direkt aus Controls
- arbeitet auch mit noch dirtyem Datensatz
- schreibt lokal in gebundene Felder

## 3. Warum Zeilenberechnung nicht automatisch zu aktuellen Totals fuehrt

Das ist im aktuellen Code kein Zufall, sondern direkte Folge der Architektur:

1. `CalcLineTotals()` rechnet nur lokal in der aktuellen Zeile
2. `Form_AfterUpdate()` im Subform rechnet Totale ausdruecklich nicht
3. `Form_AfterDelConfirm()` rechnet Totale ausdruecklich nicht
4. die Totals-Services lesen aus Tabellen/Recordsets, nicht aus unsaved Control-Werten
5. bei einer noch unsaved Zeile sind die berechneten Werte zwar im Formular sichtbar, aber noch nicht zwangslaeufig in der DB-Basis enthalten, die `CalculateTotalsForContext()` summiert

## 4. Header-Summenberechnung im Ist-Zustand

### Servicepfad

- `modOrderCalculationService.CalculateOrderTotals()`
- `modOrderCalculationService.CalculateTemporaryOrderTotals()`
- intern `CalculateTotalsForContext()` (Zeilen 299-383)

### Was passiert

1. Header-Recordset oeffnen
2. alle Zeilen fuer Header-ID lesen
3. Summen bilden aus:
   - `line_net_amount`
   - `line_vat_amount`
   - `line_gross_amount`
4. Header schreiben:
   - `subtotal_net_amount`
   - `header_discount_amount = 0`
   - `header_surcharge_amount = 0`
   - `net_amount`
   - `vat_amount`
   - `gross_amount`

### Was nicht passiert

- keine Ruecksicht auf noch nicht gespeicherte aktuelle Zeile
- keine Anwendung von Header-Discount/Surcharge, nur Warnung
- keine UI-Aktualisierung von selbst

## 5. Gesamtrecalc

### Temp

- `RecalculateTemporaryOrder(tmp_order_id)`
  - alle tmp-Zeilen neu rechnen
  - danach tmp-Header summieren

### Persisted

- `RecalculateOrder(order_id)`
  - alle persisted Zeilen neu rechnen
  - danach persisted Header summieren

## 6. Wann Totale aktuell neu gerechnet werden

### Ja

- `frmOrderDetail.ApplyCurrentRecord()`
- `frmOrderDetail.DetailSave()` indirekt ueber `ApplyCurrentRecord()` oder Persistierung
- `frmOrderDetail.cboVatMode_AfterUpdate()`
- `frmOrderDetail.ApplyDeliveryAddressVatDefaults()`
- `modOrderRepository.PersistTemporaryOrder()` am Ende fuer persisted order

### Nein

- normale Zeilenfeld-Aenderung in `sfrmOrderLines`
- Zeilenloeschung in `sfrmOrderLines`
- `txtVatRate_AfterUpdate()` im Header
- Sprachwechsel

## 7. Inkonsistenzrisiko vor Persistierung

Vor `PersistTemporaryOrder()` koennen folgende Situationen auftreten:

- aktuelle Zeile lokal sichtbar, aber noch nicht gespeichert
- tmp-Header-Totale alt
- `tmp_order_line` teilweise aktuell, teilweise nicht

`PersistTemporaryOrder()` fuehrt selbst keine Zwangs-Neuberechnung des tmp-Headers vor dem Kopieren aus.
Es vertraut darauf, dass die tmp-Daten bereits korrekt gerechnet wurden.

Das ist ein Risiko:
- fachlich richtige Werte koennen im Formular sichtbar sein
- aber noch nicht konsistent in `tmp_order`/`tmp_order_line` vorliegen

## 8. Hauptursache der Summenprobleme

Die Summenprobleme entstehen aus drei Ebenen gleichzeitig:

1. lokale Zeilenberechnung ist UI-seitig
2. Header-Totale sind DB-seitig
3. die Bruecke zwischen beidem wird nur ueber spaetere Save/Apply/Recalc-Punkte geschlagen

Kurz:
- aktuelle Zeile und Header summieren nicht im selben Moment auf derselben Datenbasis

## 9. Architekturfolgerung

Es braucht kuenftig genau eine Regel:

- entweder
  - aktuelle Zeile sofort speichern, dann Header neu summieren
- oder
  - komplettes Edit-Modell im Speicher halten und Header aus demselben In-Memory-Modell berechnen

Der aktuelle Hybrid aus:
- lokaler UI-Zeilenberechnung
- spaeterer DB-Gesamtrechnung

ist die zentrale technische Ursache der wiederkehrenden Summenabweichungen.
