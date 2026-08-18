# Workspace OpenArgs / Payload Analyse

## Gefundene Mechanismen

1. Echtes Access `OpenArgs`
2. `where_condition` Parameter in `OpenWorkspaceForm`
3. Pending Payload:
   - `m_pendingWorkspaceFormName`
   - `m_pendingWorkspaceOpenArgs`
4. `ConsumePendingWorkspaceOpenArgs(form_name)`
5. Formularmethode `ApplyWorkspaceOpenArgs(openArgs)`
6. History-Payload:
   - `HISTORY_KEY_OPEN_ARGS`
7. Formularspezifische Payloads:
   - z.B. `ORDER_EDIT;<tmp_order_id>`

## Warum das existiert

Das Workspace-Hosting via `subWorkspaceHost.SourceObject = "Form.<Name>"` transportiert `OpenArgs` nicht so wie `DoCmd.OpenForm`.

Deshalb wurde eine technische Brücke gebaut:
- Shell setzt Pending Payload vor `SourceObject`-Wechsel
- Formular liest im `Form_Load` zuerst `Me.OpenArgs`
- falls leer, liest es `ConsumePendingWorkspaceOpenArgs`
- zusätzlich ruft der Workspace-Service später noch `ApplyWorkspaceOpenArgs` explizit auf

## Probleme

1. Doppelter Kontexttransport
- `Me.OpenArgs`
- Pending Payload
- `ApplyWorkspaceOpenArgs`

2. Verspätete Übergabe
- `Form_Load` kann ohne finalen Kontext laufen

3. Konkurrierende Quellen
- `Me.OpenArgs`
- Pending Store
- History Restore
- Filter/WhereCondition

4. Formulare müssen defensiv leer starten
- exemplarisch `frmOrderDetailNext`

## Empfehlung

Kurzfristig:
- Pending Payload als technische Host-Brücke beibehalten
- aber genau einen effektiven Kontextpfad pro Formular definieren
- keine zusätzlichen Fallback-Ebenen pro Formular mehr einführen

Zielbild:
- hosted forms: genau ein Framework-Payload-Mechanismus
- standalone forms: natives `OpenArgs`
- keine Mischung aus mehreren gleichwertigen Quellen im selben Formular

## Konkrete Bewertung

- `Pending OpenArgs`: `KEEP_AND_SIMPLIFY`
- `ConsumePendingWorkspaceOpenArgs`: `KEEP_AND_SIMPLIFY`
- `ApplyWorkspaceOpenArgs`: `KEEP`
- gemischte Formular-Fallbacks: `REMOVE`
- zusätzliches Ableiten des Kontexts aus aktuellem Datensatz: `REMOVE`
