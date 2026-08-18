# Zustandsmodell

## 1. Explizit erkennbare Zustandsmerkmale

Im aktuellen Code existiert kein einziges zentrales State-Objekt. Der Modus wird aus mehreren Signalen rekonstruiert:

- `m_isTemporaryOrder`
- `m_temporaryOrderId`
- `m_temporaryOrderDisposed`
- `m_workspaceOpenArgsApplied`
- `m_lastAppliedWorkspaceOpenArgs`
- `Me.RecordSource`
- `Me.Filter` / `Me.FilterOn`
- `GetCurrentOrderId()`
- `GetCurrentTemporaryOrderId()`
- `GetCurrentOrderHeaderId()`
- `Me.NewRecord`
- `modAppWorkspaceService` pending OpenArgs
- Workspace-History-State (`tmp_order_id` oder `order_id`)

## 2. Rekonstruierte Ist-Zustaende

### STATE_A: DEFERRED_UNBOUND_WORKSPACE_WAIT

Erkennung:
- `ResolveHostingShellForm() <> Nothing`
- keine direkten OpenArgs im `Form_Load`
- `PrepareDeferredWorkspaceContext()` setzt:
  - `m_workspaceOpenArgsApplied = False`
  - `m_isTemporaryOrder = False`
  - `m_temporaryOrderId = 0`
  - `Filter = [order_id]=0`
  - `FilterOn = True`

Bedeutung:
- Formular ist geladen, kennt aber noch keinen finalen Business-Kontext

Risiko:
- Subform und UI koennen bereits anlaufen, bevor der eigentliche Modus feststeht

### STATE_B: EDIT_TEMPORARY

Erkennung:
- `m_isTemporaryOrder = True`
- `m_temporaryOrderId > 0`
- `Me.RecordSource = "tmp_order"`

Technische Folgen:
- `subOrderLines` wird auf `tmp_order_line` konfiguriert
- `GetCurrentOrderHeaderId()` liefert `tmp_order_id`
- `DetailSave()` persistiert
- `DetailApply()` rechnet nur neu und bleibt offen
- `CanLeaveWorkspace()` und `CancelCurrentEdit()` loeschen tmp-Daten

### STATE_C: EDIT_PERSISTED

Erkennung:
- `m_isTemporaryOrder = False`
- `GetCurrentOrderId() > 0`
- `Me.RecordSource = "ord_order"`

Technische Folgen:
- `subOrderLines` wird auf `ord_order_line` konfiguriert
- `DetailApply()` speichert und rechnet
- `DetailSave()` speichert und navigiert zurueck
- `CancelCurrentEdit()` macht nur `Me.Undo`

### STATE_D: TEMP_DISPOSED_BUT_FORM_STILL_RUNNING

Erkennung:
- `m_temporaryOrderDisposed = True`
- Formularobjekt ist noch aktiv oder im Schliess-/Navigationspfad

Bedeutung:
- tmp-Daten wurden geloescht
- Form/History/Translation koennen aber noch laufen

Risiko:
- spaete Routinen treffen auf bereits geloeschte oder nicht mehr verfuegbare Daten

### STATE_E: PERSISTING_TEMP_TO_ORDER

Erkennung:
- `PersistTemporaryOrderAndReload()` laeuft
- `m_isTemporaryOrder` wird lokal erst nach Repository-Persistierung umgelegt

Risiko:
- Temp- und Persisted-Form koennen kurz hintereinander bzw. ueberlappend initialisieren

### STATE_F: CHILD_NEW_LINE_PENDING

Erkennung:
- `sfrmOrderLines.Form_BeforeInsert`
- Parent muss bereits einen Header-Key liefern

Risiko:
- im Temp-Modus ist der "Header-Key" nicht `order_id`, sondern `tmp_order_id`
- dieselbe Child-Logik muss beide Welten tragen

## 3. Mehrdeutigkeiten

### 3.1 `GetCurrentOrderHeaderId()`

Problem:
- liefert im Temp-Modus `tmp_order_id`
- liefert im Persisted-Modus `order_id`

Folge:
- der Name suggeriert eine einheitliche `order_id`, tatsaechlich ist es eine kontextabhaengige Header-ID

### 3.2 `GetCurrentWorkspaceRecordId()`

Problem:
- fuer Workspace-History ist die Rueckgabe im Temp-Modus eine `tmp_order_id`
- im Persisted-Modus eine `order_id`

Folge:
- derselbe History-Kanal transportiert zwei unterschiedliche ID-Domaenen

### 3.3 `Me.RecordSource`

Problem:
- dasselbe Formularobjekt wechselt zwischen `ord_order` und `tmp_order`

Folge:
- Control-/Field-Pruefungen, Requery, Filter und Subform-Linking haengen am aktuellen Timing

### 3.4 OpenArgs-Verarbeitung

Es gibt drei konkurrierende Pfade:
- `Me.OpenArgs`
- pending OpenArgs aus `modAppWorkspaceService.ConsumePendingWorkspaceOpenArgs()`
- spaeterer direkter Aufruf `modAppWorkspaceService.ApplyWorkspaceOpenArgs()`

Zusatzsignal:
- `m_lastAppliedWorkspaceOpenArgs` unterdrueckt Duplikate, loest aber die grundsaetzliche Mehrfachinitialisierung nicht sauber auf

## 4. Widerspruechliche Regeln

### 4.1 Temp-Zurueck vs Temp-Weiterbearbeitung

- `CanLeaveWorkspace()` loescht tmp-Daten bereits vor erfolgreichem Verlassen des Forms
- gleichzeitig kann `OpenWorkspaceForm()` spaeter noch superseded oder recoverable werden

Das ist ein Zustandssprung ohne Transaktionsgrenze.

### 4.2 Header-ID fuer Child-Freigabe

- `UpdateOrderLinesAvailability()` erlaubt Zeilenbearbeitung sobald `GetCurrentOrderHeaderId() > 0`
- im Temp-Modus bedeutet das: Zeilen sind erlaubt, obwohl noch keine `ord_order` existiert
- im Persisted-Modus bedeutet es: echte `order_id`

Die gleiche UI-Regel deckt zwei fachlich verschiedene Bedeutungen ab.

### 4.3 Neue Bestellung vs bestehende Bestellung

- neue Bestellung ist bereits ein physischer temp-Datensatz
- aber in UI und Workflow soll sie wie "noch nicht gespeichert" wirken

Das fuehrt zu gemischter Semantik:
- fachlich noch nicht persistiert
- technisch bereits persistiert

## 5. Fehlende explizite Zustaende

Es fehlen klar modellierte States fuer:
- `LOADING_TEMP`
- `LOADING_PERSISTED`
- `NAVIGATING_AWAY`
- `APPLYING_TOTALS`
- `PERSIST_COMMIT_IN_PROGRESS`
- `CANCEL_PENDING_CONFIRMATION`

Diese Zustaende existieren faktisch im Verhalten, aber nicht als sauber pruefbare Zustandsmaschine.

## 6. Schlussfolgerung

Das aktuelle Formular ist nicht nur "komplex", sondern zustandslogisch unscharf:
- mehrere ID-Welten
- mehrere Initialisierungspfade
- mehrere Speichersementiken
- mehrere Zeitpunkte fuer dieselbe Business-Entscheidung

Das ist der Kern der wiederkehrenden Fehlerbilder.
