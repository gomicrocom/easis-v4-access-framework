# Laufzeitpfade

## 1. Neue Bestellung

### 1.1 AddressList -> AddressCockpit

1. `Form_frmAddressList.OpenAddressCockpit()` (Zeilen 375-392)
   - liest `address_id` aus der aktuellen Liste
   - ruft `OpenTargetForm FORM_ADDRESS_COCKPIT, ..., CStr(currentAddressId)` auf
   - OpenArgs: `address_id` als String
   - Workspace: ja

2. `modAppWorkspaceService.OpenWorkspaceForm()` (Zeilen 31-176)
   - setzt `subWorkspaceHost.SourceObject`
   - haengt pending OpenArgs an `m_pendingWorkspaceFormName` / `m_pendingWorkspaceOpenArgs`
   - ruft optional `ApplyWorkspaceOpenArgs` auf dem geladenen Formular
   - sichert optional History des vorherigen Workspace-Formulars

3. `Form_frmAddressCockpit.ApplyWorkspaceOpenArgs()` (Zeilen 193-206)
   - parst `address_id`
   - positioniert Formular auf gewaehlte Adresse via `RestoreCurrentAddressSelection`

### 1.2 AddressCockpit -> Neue Bestellung

1. `Form_frmAddressCockpit.CreateNewSalesOrder()` (Zeilen 232-261)
   - liest `address_id`
   - ruft `modAddressCockpitService.CreateTemporarySalesOrderWorkspaceArgs(addressId)`

2. `modAddressCockpitService.CreateTemporarySalesOrderWorkspaceArgs()` (Zeilen 22-48)
   - validiert Adresse
   - ruft `modOrderRepository.CreateTemporarySalesOrderForAddress(addressId)`
   - erzeugt Rueckgabeformat:
     - `TMP_ORDER;<tmp_order_id>`

3. `modOrderRepository.CreateTemporarySalesOrderForAddress()` (Zeilen 534-596)
   - DB: `modDb.GetCurrentTenantDatabase()`
   - Tabelle: `tmp_order`
   - erzeugt sofort einen physischen Datensatz in `tmp_order`
   - schreibt Defaults:
     - `customer_address_id`
     - `invoice_address_id`
     - `delivery_address_id`
     - `customer_name`
     - `order_type_code=SO`
     - `order_status_code=DRAFT`
     - `order_date=Date`
     - `currency_code`
     - `payment_term_code`
     - `language_code`
     - `vat_mode`, `vat_code`, `vat_rate`
   - noch keine `ord_order`
   - noch keine `order_no`

4. `modAppWorkspaceService.OpenWorkspaceForm(..., "frmOrderDetail", ..., open_args: "TMP_ORDER;<id>")`
   - setzt pending OpenArgs
   - laedt `frmOrderDetail` im Workspace

### 1.3 frmOrderDetail fuer temporaere Bestellung

1. `Form_frmOrderDetail.Form_Open()` (Zeilen 55-63)
   - ruft `PrimePaymentTermComboEarly()`
   - fachlich nur UI-Vorbereitung

2. `Form_frmOrderDetail.Form_Load()` (Zeilen 65-111)
   - liest `Me.OpenArgs`
   - wenn leer: `modAppWorkspaceService.ConsumePendingWorkspaceOpenArgs(Me.Name)`
   - `InitializeRecordSource()`
   - bei OpenArgs vorhanden:
     - `ApplyWorkspaceOpenArgs effectiveOpenArgs`
   - sonst optional neutraler Wait-Zustand ueber `PrepareDeferredWorkspaceContext()`
   - danach:
     - `modFwTranslationRuntime.ApplyTranslations Me`
     - `UpdateLocalCommandButtons`
     - `UpdateHeaderCaptionUi`

3. `Form_frmOrderDetail.ApplyWorkspaceOpenArgs("TMP_ORDER;<id>")` (Zeilen 409-469)
   - erkennt Temp-Modus ueber Prefix `TMP_ORDER`
   - ruft `ActivateTemporaryOrderContext(tmpOrderId)`

4. `Form_frmOrderDetail.ActivateTemporaryOrderContext()` (Zeilen 1003-1036)
   - setzt:
     - `m_isTemporaryOrder = True`
     - `m_temporaryOrderId = tmpOrderId`
     - `m_temporaryOrderDisposed = False`
   - schaltet `Me.RecordSource = "tmp_order"`
   - `Me.Requery`
   - `RestoreCurrentTemporaryOrderSelection tmpOrderId`
   - `ConfigureOrderLinesSubform`
   - `ConfigureHeaderControls`
   - `UpdateOrderLinesAvailability`
   - `RequeryOrderLinesSubform`
   - `UpdateHeaderCaptionUi`

5. `Form_frmOrderDetail.ConfigureOrderLinesSubform()` (Zeilen 566-594)
   - `SourceObject = Form.sfrmOrderLines`
   - Temp-Modus:
     - `LinkMasterFields = tmp_order_id`
     - `LinkChildFields = tmp_order_id`
   - ruft danach `ConfigureOrderLinesSubformInstance()`

6. `Form_frmOrderDetail.ConfigureOrderLinesSubformInstance()` (Zeilen 1059-1077)
   - Temp-Modus:
     - `sfrmOrderLines.ConfigureForOrderContext "tmp_order_line", "tmp_order_id", "tmp_order_line_id"`

7. `Form_sfrmOrderLines.Form_Load()` (Zeilen 46-64)
   - setzt Default-RecordSource zunaechst auf `ord_order_line`, falls Konfiguration noch nicht erfolgt ist
   - konfiguriert Anpassungscombos und VAT-Combo
   - ruft `modFwTranslationRuntime.ApplyTranslations Me`
   - spaeter wird per `ConfigureForOrderContext` auf `tmp_order_line` umgeschaltet

### 1.4 Zeilenanlage im Temp-Modus

1. `sfrmOrderLines.Form_BeforeInsert()` -> `EnsureNewLineDefaults()` (Zeilen 165-287)
   - fragt `GetParentOrderId()` ab
   - Parent ruft `frmOrderDetail.GetCurrentOrderHeaderId()`
   - im Temp-Modus kommt `tmp_order_id`
   - setzt:
     - Parent-Key (`tmp_order_id`)
     - `line_no = DMax(...) + 10`
     - VAT-Defaults aus Parent

2. `sfrmOrderLines` Feld-Aenderungen
   - `txtQuantity_AfterUpdate`
   - `txtUnitPrice_AfterUpdate`
   - `txtDiscountValue_AfterUpdate`
   - `cboDiscountType_AfterUpdate`
   - `txtSurchargeValue_AfterUpdate`
   - `cboSurchargeType_AfterUpdate`
   - `cboVatCode_AfterUpdate`
   - alle rufen `CalcLineTotals()`

3. `CalcLineTotals()` (Zeilen 308-396)
   - liest Werte direkt aus Controls
   - berechnet nur die aktuelle Zeile lokal
   - schreibt nur Zeilenfelder
   - speichert nicht
   - rechnet Header-Totale nicht

### 1.5 Temp -> Persisted

1. `frmOrderDetail.DetailApply()`
   - ruft nur `ApplyCurrentRecord()`
   - Temp-Modus bleibt offen

2. `frmOrderDetail.ApplyCurrentRecord()` (Zeilen 785-818)
   - speichert Header-Datensatz, falls `Me.Dirty`
   - ruft im Temp-Modus `modOrderCalculationService.RecalculateTemporaryOrder(m_temporaryOrderId)`
   - danach `RefreshCalculatedTotals True`
   - keine Persistierung nach `ord_order`

3. `frmOrderDetail.DetailSave()` (Zeilen 479-485)
   - Temp-Modus:
     - `PersistTemporaryOrderAndReload()`

4. `frmOrderDetail.PersistTemporaryOrderAndReload()` (Zeilen 1079-1105)
   - speichert Header, falls `Me.Dirty`
   - ruft `modOrderRepository.PersistTemporaryOrder(m_temporaryOrderId)`
   - setzt lokalen Temp-Status auf erledigt
   - navigiert danach erneut zu `frmOrderDetail` mit OpenArgs = echte `order_id`

5. `modOrderRepository.PersistTemporaryOrder()` (Zeilen 650-797)
   - liest `tmp_order`
   - erzeugt `order_no`
   - legt neuen Datensatz in `ord_order` an
   - kopiert Headerfelder
   - kopiert `tmp_order_line` nach `ord_order_line`
   - ruft `modOrderCalculationService.RecalculateOrder(OrderId)`
   - loescht danach `tmp_order` und `tmp_order_line`

## 2. Bestehende Bestellung

Ein eigenstaendiges `frmOrderList` oder ein dediziertes persistiertes Order-Einstiegsformular wurde im analysierten Exportbestand nicht gefunden.

Tatsaechlich nachweisbare bestaende-Pfade:

1. `PersistTemporaryOrderAndReload()` -> `OpenWorkspaceForm(..., CStr(orderId))`
2. generischer externer Aufrufer koennte `OpenWorkspaceForm(..., "frmOrderDetail", ..., CStr(orderId))` verwenden

### 2.1 Persisted Load

1. `OpenWorkspaceForm(..., open_args = "<order_id>")`
2. `frmOrderDetail.Form_Load()`
3. `ApplyWorkspaceOpenArgs("<order_id>")`
   - setzt:
     - `m_isTemporaryOrder = False`
     - `m_temporaryOrderId = 0`
     - `Me.RecordSource = "ord_order"`
   - validiert `modOrderRepository.OrderExists(orderId)`
   - `RestoreCurrentOrderSelection`
   - `ConfigureOrderLinesSubform`
   - `ConfigureHeaderControls`
   - `RequeryOrderLinesSubform`
   - `UpdateOrderLinesAvailability`
   - `UpdateHeaderCaptionUi`

### 2.2 Persisted Edit

1. Feld-Aenderungen im Header
   - bleiben in gebundenem Formular bis Save/Apply/Cancel
2. Feld-Aenderungen in `sfrmOrderLines`
   - rechnen lokal nur die aktuelle Zeile
   - Header bleibt bis spaeteren Recalc unveraendert
3. `DetailApply()`
   - `ApplyCurrentRecord()`
   - `modOrderCalculationService.RecalculateOrder(order_id)`
   - `RefreshCalculatedTotals True`
   - Formular bleibt offen
4. `DetailSave()`
   - `ApplyCurrentRecord()`
   - danach `ReturnToPreviousWorkspace`

## 3. Zentrale Beobachtungen zum Laufzeitpfad

- `frmOrderDetail` ist gleichzeitig:
  - Workspace-Host-Form
  - Temp-Editor
  - Persisted-Editor
  - Navigationsziel nach Persistierung
- die Reihenfolge `Form_Load -> ApplyWorkspaceOpenArgs -> Requery -> Subform reconfigure -> Translation/UI refresh` ist empfindlich gegen:
  - Back-Navigation waehrend des Ladens
  - History-Restore
  - spaete Requeries
  - erneute `ApplyWorkspaceOpenArgs`
- `sfrmOrderLines` kann vor finaler Kontextumschaltung bereits mit altem oder Default-RecordSource geladen werden
- das Formular arbeitet stark zustandsgetrieben, ohne einen expliziten einzigen Modus-Controller
