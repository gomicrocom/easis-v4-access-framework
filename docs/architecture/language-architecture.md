# Spracharchitektur

Stand: 2026-07-29

## Zielformat

Easis v4 speichert Sprachcodes im BCP-47-Format:

- Sprache klein
- Bindestrich
- Region gross

Beispiele:

- `de-CH`
- `fr-CH`
- `it-CH`
- `en-US`

## System Default

Der technische System-Default ist `en-US`.

## UI-Sprache

Die UI-Sprache wird in folgender Reihenfolge aufgeloest:

1. `usr_user.language_code`
2. `ten_parameter.DEFAULT_LANGUAGE`
3. `en-US`

Hinweis:
`usr_user.language_code` ist architektonisch vorbereitet, aber noch nicht in jedem Anwendungspfad produktiv befuellt.

## Dokumentensprache

Die Dokumentensprache ist von der UI-Sprache getrennt und wird in folgender Reihenfolge aufgeloest:

1. `adr_address.language_code`
2. `ten_parameter.DEFAULT_LANGUAGE`
3. `en-US`

Die UI-Sprache darf die Dokumentensprache nicht implizit ueberschreiben.

## Bekannte Sprachen

`ref_language` ist der zentrale Sprachkatalog aller im System bekannten Sprachen.

Aktuell vorgesehen:

- `de-CH`
- `fr-CH`
- `en-US`
- `it-CH`
- optional weitere bekannte Sprachen wie `de-DE`

Bekannte Sprachen werden nicht geloescht, nur weil sie noch nicht vollstaendig uebersetzt sind.

## Vollstaendig unterstuetzte Uebersetzungssprachen

Aktuell gelten fuer Audit, Formularuebersetzungen und Standard-UI als vollstaendig unterstuetzt:

- `de-CH`
- `fr-CH`
- `en-US`

Fehlende Uebersetzungen fuer vorbereitete Sprachen wie `it-CH` sind derzeit kein Audit-Fehler.

## Migration alter Sprachcodes

Folgende Altcodes werden auf das Zielformat normalisiert:

- `DE-CH` -> `de-CH`
- `de_CH` -> `de-CH`
- `FR-FR` -> `fr-CH`
- `fr_FR` -> `fr-CH`
- `fr_CH` -> `fr-CH`
- `EN-US` -> `en-US`
- `en_US` -> `en-US`
- `IT-CH` -> `it-CH`
- `it_CH` -> `it-CH`

Die Datenmigration laeuft konfliktarm:

- sichere Normalisierungen werden direkt aktualisiert
- Zielkonflikte werden protokolliert und nicht still zusammengefuehrt

VBA-Einstieg:

- `modFwSetup.NormalizeLanguageCodeData`

## Referenzdaten

Sichtbare Referenztexte werden weiterhin ueber das bestehende Uebersetzungssystem aufgeloest.

Nicht vorgesehen:

- sprachabhaengige Duplikate in Referenztabellen
- alternative Uebersetzungssysteme
- neue Sprachspalten fuer sichtbare Texte in Referenztabellen

## Resthinweis

Einige Runtime-Normalisierungen akzeptieren temporaer noch Legacy-Eingaben, damit bestehende Daten vor der Migration nicht sofort brechen. Nach erfolgreicher Datenmigration und Access-Compile koennen diese Alias-Pfade weiter reduziert werden.
