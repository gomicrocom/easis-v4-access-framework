# PDF Output Service

## Zweck

Der PDF Output Service stellt eine zentrale, konfigurierbare Ablage fuer Dokument-PDFs bereit. Die UI- oder Handler-Schicht soll keine festen Ausgabepfade mehr kennen, sondern den Export ueber die Framework-Services ausfuehren.

Single source of truth fuer den fachlichen Dokument-PDF-Export ist `modPdfExportService`.

## Konfiguration

Der Document Root Path wird wie folgt ermittelt:

1. Wenn in der Backend-Tabelle `ten_parameter` der Parameter `BASIC_DOC_PATH` gesetzt ist, wird dieser Wert verwendet.
2. Wenn `BASIC_DOC_PATH` fehlt oder leer ist, wird der Default-Pfad aus dem Frontend-Verzeichnis aufgebaut:

```text
<CurrentProject.Path>\Docs\<TenantName>
```

`TenantName` kommt bevorzugt aus dem Tenant-Kontext und faellt bei Bedarf auf den Tenant-Parameter `TENANT_NAME` zurueck. Der Wert wird vor der Verwendung als Pfadsegment bereinigt.

Verwendete Tenant-Parameter:

```text
BASIC_DOC_PATH
TENANT_NAME
```

Beispiel:

```text
BASIC_DOC_PATH = C:\Easis\TestDocs
```

## Zielstruktur

PDFs werden nach folgendem Muster abgelegt:

```text
<DocumentDirectory>\<CustomerName>\<DocumentNo>.pdf
```

Ohne `BASIC_DOC_PATH` ergibt sich damit:

```text
<CurrentProject.Path>\Docs\<TenantName>\<CustomerName>\<DocumentNo>.pdf
```

`TenantName`, `CustomerName` und `DocumentNo` werden vor der Verwendung als Pfadsegmente bereinigt.

## Exportverantwortung

- `modPdfExportService.ExportDocumentToPdf` ist die einzige primaere PDF-Exportfunktion.
- `rpt_document` ist der einzige Dokumentreport fuer die PDF-Erzeugung.
- `modOutputPathService` liefert nur Zielpfade und Verzeichnis-Helpers.
- Es gibt keine doppelte fachliche Exportlogik in mehreren Modulen.

## Beteiligte Module

- `modOutputPathService`: liest den tenant-spezifischen Dokumentpfad, bereinigt Pfadsegmente und baut Zielpfade.
- `modPdfExportService`: validiert Dokument, Report und Zielordner und fuehrt den PDF-Export aus.
- `modDocumentExportHandler`: optionaler Legacy-Wrapper fuer bestehende Aufrufer und delegiert an den PDF-Service.
