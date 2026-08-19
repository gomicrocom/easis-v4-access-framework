# Document Calculation Migration

## Zweck

Diese Migration erweitert `doc_document` und `doc_document_position` um Felder fuer:

- Rabatt
- Aufschlag
- gespeicherte Zwischen- und Endbetraege

Die Berechnung erfolgt anschliessend in `modDocumentCalculationService`.

## Access DDL

### doc_document_position

```sql
ALTER TABLE doc_document_position ADD COLUMN discount_type TEXT(20);
ALTER TABLE doc_document_position ADD COLUMN discount_value CURRENCY;
ALTER TABLE doc_document_position ADD COLUMN surcharge_type TEXT(20);
ALTER TABLE doc_document_position ADD COLUMN surcharge_value CURRENCY;
ALTER TABLE doc_document_position ADD COLUMN line_base_amount CURRENCY;
ALTER TABLE doc_document_position ADD COLUMN line_discount_amount CURRENCY;
ALTER TABLE doc_document_position ADD COLUMN line_surcharge_amount CURRENCY;
```

### doc_document

```sql
ALTER TABLE doc_document ADD COLUMN header_discount_type TEXT(20);
ALTER TABLE doc_document ADD COLUMN header_discount_value CURRENCY;
ALTER TABLE doc_document ADD COLUMN header_surcharge_type TEXT(20);
ALTER TABLE doc_document ADD COLUMN header_surcharge_value CURRENCY;
ALTER TABLE doc_document ADD COLUMN subtotal_net_amount CURRENCY;
ALTER TABLE doc_document ADD COLUMN header_discount_amount CURRENCY;
ALTER TABLE doc_document ADD COLUMN header_surcharge_amount CURRENCY;
```

## Initialwerte

```sql
UPDATE doc_document_position
SET
    discount_type = 'NONE',
    discount_value = 0,
    surcharge_type = 'NONE',
    surcharge_value = 0,
    line_base_amount = 0,
    line_discount_amount = 0,
    line_surcharge_amount = 0
WHERE
    discount_type IS NULL
    OR discount_value IS NULL
    OR surcharge_type IS NULL
    OR surcharge_value IS NULL
    OR line_base_amount IS NULL
    OR line_discount_amount IS NULL
    OR line_surcharge_amount IS NULL;
```

```sql
UPDATE doc_document
SET
    header_discount_type = 'NONE',
    header_discount_value = 0,
    header_surcharge_type = 'NONE',
    header_surcharge_value = 0,
    subtotal_net_amount = 0,
    header_discount_amount = 0,
    header_surcharge_amount = 0
WHERE
    header_discount_type IS NULL
    OR header_discount_value IS NULL
    OR header_surcharge_type IS NULL
    OR header_surcharge_value IS NULL
    OR subtotal_net_amount IS NULL
    OR header_discount_amount IS NULL
    OR header_surcharge_amount IS NULL;
```

## Hinweis

Im aktuellen Implementierungsschritt fuehrt `modDocumentCalculationService.EnsureDocumentCalculationSchema` diese Felder auch defensiv per Code nach, falls sie noch fehlen. Die DDL bleibt die empfohlene explizite Migration fuer kontrollierte Rollouts.
