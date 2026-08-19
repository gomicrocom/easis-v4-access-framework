# frmOrderDetail Analyse

Datum: 2026-08-02

Analysierter Umfang:
- Formulare: `frmOrderDetail`, `sfrmOrderLines`, `frmAddressCockpit`, `frmAddressList`
- Module: `modAppWorkspaceService`, `modAddressCockpitService`, `modOrderRepository`, `modOrderCalculationService`
- Queries: `qry_order_detail_header`, `qry_order_lines`, `qry_order_list`, `qry_order_vat_summary`

Ziel dieser Analyse:
- den echten Laufzeitpfad fuer neue und bestehende Bestellungen sichtbar machen
- implizite Formularzustaende rekonstruieren
- Konflikte zwischen Workspace, OpenArgs, temp/persisted Daten und Summenberechnung belegen
- eine belastbare Entscheidung zwischen gezieltem Refactoring und kontrolliertem Neuaufbau ermoeglichen

Wichtigste Ergebnisse:
- `frmOrderDetail` vereint aktuell drei Modi in einem gebundenen Access-Formular:
  - `EDIT_PERSISTED`
  - `EDIT_TEMPORARY`
  - `DEFERRED_UNBOUND_WORKSPACE_WAIT`
- die Moduserkennung ist mehrfach und teilweise widerspruechlich implementiert:
  - `Me.OpenArgs`
  - pending OpenArgs aus `modAppWorkspaceService`
  - `m_isTemporaryOrder`
  - `m_workspaceOpenArgsApplied`
  - `GetCurrentOrderId()`
  - `GetCurrentTemporaryOrderId()`
  - Workspace-State-Restore (`tmp_order_id` oder `order_id`)
- die Summen werden absichtlich nicht bei Zeilen-Aenderungen aktualisiert:
  - `sfrmOrderLines.CalcLineTotals()` rechnet nur lokal in der aktuellen Zeile
  - `Form_AfterUpdate` und `Form_AfterDelConfirm` rechnen Totale ausdruecklich nicht neu
  - Header-Totale werden erst ueber `DetailApply`, `DetailSave`, `cboVatMode_AfterUpdate` oder `ApplyDeliveryAddressVatDefaults()` neu gerechnet
- dieselbe Zeilenberechnung existiert doppelt:
  - lokal im Formular `sfrmOrderLines.CalcLineTotals()`
  - servicebasiert in `modOrderCalculationService.CalculateLineAmountsForContext()`
- `CanLeaveWorkspace()` loescht temporaere Daten bereits vor erfolgreichem Abschluss des Workspace-Wechsels
- `frmOrderDetail` schaltet seine `RecordSource` zur Laufzeit zwischen `ord_order` und `tmp_order` um; `sfrmOrderLines` schaltet parallel zwischen `ord_order_line` und `tmp_order_line`

Empfehlung:
- `REBUILD_CONTROLLED`

Begruendung in Kurzform:
- zu viele ueberschneidende Eventpfade
- temp- und persisted-Kontext teilen sich dasselbe gebundene Formularobjekt
- lokale Berechnung, Requery, History-Restore und Navigation greifen ineinander
- weitere punktuelle Fixes werden mit hoher Wahrscheinlichkeit nur neue Randfehler erzeugen

Artefakte:
- `order-detail-runtime-flow.md`
- `order-detail-event-map.csv`
- `order-detail-state-model.md`
- `order-detail-data-mapping.csv`
- `order-detail-calculation-flow.md`
- `order-detail-command-map.csv`
- `order-detail-issues.csv`
- `order-detail-target-architecture.md`
- `order-detail-refactor-vs-rebuild.md`
