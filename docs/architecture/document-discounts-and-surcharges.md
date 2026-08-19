# Document Discounts And Surcharges

## Zweck

Dieses Konzept erweitert das Dokumentmodell um deklarative Rabatte und Aufschlaege auf:

- Positionsebene
- spaeter auch Dokumentkopfebene

Wichtig:

- Berechnung erfolgt in Services und Persistenzschicht
- Reports zeigen nur gespeicherte Werte
- keine Kalkulationslogik im Report

## Ist-Zustand

Aktuell verwendet das Framework fuer Dokumente diese Kernfelder:

- `doc_document.total_net`
- `doc_document.total_vat`
- `doc_document.total_gross`
- `doc_document_position.line_total_net`
- `doc_document_position.line_total_vat`
- `doc_document_position.line_total_gross`

Die aktuelle Berechnung sitzt in:

- `modDocumentService`
- `modDocumentRepository.CreateDocumentPosition`
- `modDocumentRepository.UpdateDocumentTotals`

Rabatt- und Aufschlagslogik existiert aktuell noch nicht.

## Zielmodell

### Erlaubte Typen

Fuer Rabatt- und Aufschlagstypen werden nur diese Werte zugelassen:

- `NONE`
- `PERCENT`
- `AMOUNT`

Empfehlung:

- Speicherung als `Short Text`
- Defaultwert `NONE`
- Validierung in VBA und nach Moeglichkeit per Tabellenregel

### Erweiterung `doc_document_position`

Empfohlene neue Felder:

- `discount_type` Short Text, Default `NONE`
- `discount_value` Currency, Default `0`
- `surcharge_type` Short Text, Default `NONE`
- `surcharge_value` Currency, Default `0`
- `line_base_amount` Currency
- `line_discount_amount` Currency
- `line_surcharge_amount` Currency
- `net_amount` Currency
- `vat_amount` Currency
- `gross_amount` Currency

Hinweis zur Migration:

- Die bestehenden Felder `line_total_net`, `line_total_vat`, `line_total_gross` werden aktuell im Code verwendet.
- Fuer einen sicheren ersten Schritt sollten sie vorerst bestehen bleiben.
- Empfohlene Uebergangsregel:
  - `net_amount`, `vat_amount`, `gross_amount` sind die neue fachliche Zielstruktur.
  - `line_total_net`, `line_total_vat`, `line_total_gross` bleiben temporaer fuer Kompatibilitaet bestehen oder werden in Queries auf die neuen Felder gemappt.

### Erweiterung `doc_document`

Empfohlene neue Felder:

- `header_discount_type` Short Text, Default `NONE`
- `header_discount_value` Currency, Default `0`
- `header_surcharge_type` Short Text, Default `NONE`
- `header_surcharge_value` Currency, Default `0`
- `subtotal_net_amount` Currency
- `header_discount_amount` Currency
- `header_surcharge_amount` Currency
- `net_amount` Currency
- `vat_amount` Currency
- `gross_amount` Currency

Hinweis zur Migration:

- Die bestehenden Felder `total_net`, `total_vat`, `total_gross` werden aktuell im Repository verwendet.
- Fuer einen risikoarmen ersten Schritt sollten sie ebenfalls vorerst bestehen bleiben.
- Empfohlene Uebergangsregel:
  - `subtotal_net_amount`, `net_amount`, `vat_amount`, `gross_amount` sind die neue Zielstruktur.
  - `total_net`, `total_vat`, `total_gross` bleiben temporaer kompatibel oder werden spaeter ersetzt.

## Berechnungsregeln

### Positionsebene

Reihenfolge:

1. `line_base_amount = Round(Quantity * UnitPrice, 2)`
2. Rabattbetrag aus `discount_type` und `discount_value`
3. Bemessungsgrundlage nach Rabatt darf nicht unter `0` fallen
4. Aufschlagbetrag aus `surcharge_type` und `surcharge_value` auf Basis der rabattierten Grundlage
5. `net_amount = line_base_amount - line_discount_amount + line_surcharge_amount`
6. `vat_amount` auf Basis von `net_amount`
7. `gross_amount = net_amount + vat_amount` bei exklusive MwSt. bzw. gemaess bestehender VAT-Logik

Regeln:

- `discount_value < 0` ist ungueltig
- `surcharge_value < 0` ist ungueltig
- `PERCENT` bedeutet `value / 100`
- `AMOUNT` bedeutet fixer Betrag
- Rabatt wird auf maximal die aktuelle Bemessungsgrundlage gekappt
- Aufschlag wird nicht negativ
- Rabatt wird immer zuerst berechnet
- Aufschlag wird immer auf die bereits rabattierte Basis berechnet

### Dokumentkopfebene

Empfohlene zweite Phase:

1. `subtotal_net_amount = Sum(position.net_amount)`
2. Header-Rabatt auf `subtotal_net_amount`
3. Header-Aufschlag auf die reduzierte Grundlage nach Rabatt
4. finales Dokument-Netto
5. VAT/Gross auf Dokumentebene neu berechnen oder konsistent aggregieren

Auch auf Dokumentkopfebene gilt damit dieselbe Reihenfolge:

- zuerst Rabatt
- danach Aufschlag auf die rabattierte Basis

## Empfohlene VBA-Service-Funktionen

### In `modDocumentService`

Empfohlene neue Funktionen:

- `Public Function NormalizeAdjustmentType(ByVal AdjustmentType As String) As String`
- `Public Function IsAdjustmentTypeValid(ByVal AdjustmentType As String) As Boolean`
- `Public Function CalculateAdjustmentAmount(ByVal BaseAmount As Currency, ByVal AdjustmentType As String, ByVal AdjustmentValue As Currency) As Currency`
- `Public Function CalculateDocumentLineBaseAmount(ByVal Quantity As Double, ByVal UnitPrice As Currency) As Currency`
- `Public Function CalculateDocumentLineDiscountAmount(ByVal BaseAmount As Currency, ByVal DiscountType As String, ByVal DiscountValue As Currency) As Currency`
- `Public Function CalculateDocumentLineSurchargeAmount(ByVal BaseAmountAfterDiscount As Currency, ByVal SurchargeType As String, ByVal SurchargeValue As Currency) As Currency`
- `Public Function CalculateDocumentLineNetAmount(ByVal Quantity As Double, ByVal UnitPrice As Currency, ByVal DiscountType As String, ByVal DiscountValue As Currency, ByVal SurchargeType As String, ByVal SurchargeValue As Currency) As Currency`
- `Public Function ValidateAdjustmentValue(ByVal AdjustmentType As String, ByVal AdjustmentValue As Currency) As Boolean`

### In `modDocumentRepository`

Empfohlene Erweiterungen:

- `CreateDocumentPosition` um Rabatt-/Aufschlagsparameter erweitern
- `UpdateDocumentTotals` auf neue Summenfelder vorbereiten
- spaeter eigene Funktion:
  - `RecalculateDocumentAmounts(ByVal DocumentId As Long) As Boolean`

### Optional spaeter

- `modDocumentPricingService` als dedizierter Preis-/Rabattservice, wenn die Logik groesser wird

Fuer den ersten Schritt reicht eine Erweiterung von `modDocumentService`.

## Access DDL / Migration

### `doc_document_position`

```sql
ALTER TABLE doc_document_position ADD COLUMN discount_type TEXT(16);
ALTER TABLE doc_document_position ADD COLUMN discount_value CURRENCY;
ALTER TABLE doc_document_position ADD COLUMN surcharge_type TEXT(16);
ALTER TABLE doc_document_position ADD COLUMN surcharge_value CURRENCY;
ALTER TABLE doc_document_position ADD COLUMN line_base_amount CURRENCY;
ALTER TABLE doc_document_position ADD COLUMN line_discount_amount CURRENCY;
ALTER TABLE doc_document_position ADD COLUMN line_surcharge_amount CURRENCY;
ALTER TABLE doc_document_position ADD COLUMN net_amount CURRENCY;
ALTER TABLE doc_document_position ADD COLUMN vat_amount CURRENCY;
ALTER TABLE doc_document_position ADD COLUMN gross_amount CURRENCY;
```

### `doc_document`

```sql
ALTER TABLE doc_document ADD COLUMN header_discount_type TEXT(16);
ALTER TABLE doc_document ADD COLUMN header_discount_value CURRENCY;
ALTER TABLE doc_document ADD COLUMN header_surcharge_type TEXT(16);
ALTER TABLE doc_document ADD COLUMN header_surcharge_value CURRENCY;
ALTER TABLE doc_document ADD COLUMN subtotal_net_amount CURRENCY;
ALTER TABLE doc_document ADD COLUMN header_discount_amount CURRENCY;
ALTER TABLE doc_document ADD COLUMN header_surcharge_amount CURRENCY;
ALTER TABLE doc_document ADD COLUMN net_amount CURRENCY;
ALTER TABLE doc_document ADD COLUMN vat_amount CURRENCY;
ALTER TABLE doc_document ADD COLUMN gross_amount CURRENCY;
```

### Initialisierung bestehender Daten

Empfohlene Initialwerte:

```sql
UPDATE doc_document_position
SET
    discount_type = 'NONE',
    discount_value = 0,
    surcharge_type = 'NONE',
    surcharge_value = 0
WHERE discount_type IS NULL
   OR surcharge_type IS NULL;
```

```sql
UPDATE doc_document
SET
    header_discount_type = 'NONE',
    header_discount_value = 0,
    header_surcharge_type = 'NONE',
    header_surcharge_value = 0
WHERE header_discount_type IS NULL
   OR header_surcharge_type IS NULL;
```

## Betroffene Queries

Sobald die Felder verwendet werden, muessen alle Dokument-Queries geprueft werden, die derzeit direkt auf alte Summenfelder zugreifen.

Typisch betroffen:

- Report-RecordSource fuer `rpt_document`
- Positions-Query fuer Positionslisten / Subreports
- eventuelle Dokumentlisten mit Netto-/MwSt-/Bruttospalten

Empfehlung fuer die erste Umstellung:

- Queries liefern einfache Aliasnamen fuer Reports
- Queries mappen kompatibel:
  - `net_amount AS line_total_net_compat` oder umgekehrt
- keine Rechenformeln im Report selbst

## Auswirkungen auf `rpt_document`

Der Report selbst soll spaeter nur noch lesen:

- auf Positionsebene:
  - `line_base_amount`
  - `line_discount_amount`
  - `line_surcharge_amount`
  - `net_amount`
  - `vat_amount`
  - `gross_amount`

- auf Dokumentebene:
  - `subtotal_net_amount`
  - `header_discount_amount`
  - `header_surcharge_amount`
  - `net_amount`
  - `vat_amount`
  - `gross_amount`

Moegliche spaetere UI-Ausgaben:

- Rabatt in Prozent oder Betrag je Position
- Aufschlag in Prozent oder Betrag je Position
- Zwischensumme vor Kopf-Rabatt
- Kopf-Rabatt / Kopf-Aufschlag
- finale Summen

Wichtig:

- keine `=IIf(...)`-Kalkulationen fuer Preislogik im Report
- nur Formatierung und Darstellung

## Empfohlene Testfaelle

### Positionsrabatt

- `NONE / 0` ergibt unveraenderte Basis
- `PERCENT / 10` auf `100.00` ergibt `10.00`
- `AMOUNT / 15.00` auf `100.00` ergibt `15.00`
- Rabatt `AMOUNT / 150.00` auf `100.00` wird auf `100.00` gekappt

### Positionsaufschlag

- `PERCENT / 10` auf Basis `100.00` ergibt `10.00`
- `AMOUNT / 15.00` ergibt `15.00`
- negativer Wert ist ungueltig

### Kombination

- Basis `100.00`, Rabatt `10%`, Aufschlag `5%`
- erwartetes Netto: `94.50`, wenn der Aufschlag auf die rabattierte Grundlage angewendet wird

### Leer-/Fehlerfaelle

- unbekannter Typ -> ungueltig
- negativer Wert -> ungueltig
- `PERCENT` groesser als `100` ist fachlich erlaubt, Rabattbetrag wird aber auf die Grundlage gekappt

### Dokumentkopf

- zwei Positionen werden korrekt zu `subtotal_net_amount` aufsummiert
- Kopf-Rabatt reduziert die Zwischensumme
- Kopf-Aufschlag erhoeht danach die reduzierte Grundlage
- VAT/Gross stimmen mit bestehender VAT-Logik ueberein

## Empfohlener Einfuehrungsplan

1. Felder per Migration ergaenzen
2. `modDocumentService` um Preis-/Rabattfunktionen erweitern
3. `modDocumentRepository.CreateDocumentPosition` auf neue Felder umstellen
4. `UpdateDocumentTotals` zunaechst positionsbasiert kompatibel erweitern
5. Queries auf neue gespeicherte Felder umstellen
6. erst danach `rpt_document` auf neue Felder umbauen
