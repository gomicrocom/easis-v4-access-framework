# AppShell / Workspace Architekturprüfung

Datum: 2026-08-18

Analysierter Fokus:
- `frmAppShell`
- `modAppShell`
- `modAppWorkspaceService`
- `modAppNavigationService`
- `frmAddressList`
- `frmAddressDetail`
- `frmAddressCockpit`
- `frmOrderDetail`
- `frmOrderDetailNext`
- `sfrmOrderLinesNext`
- `modOrderEditSessionService`

Leitfrage:
- Welche Framework-Mechanismen ergänzen Access sinnvoll?
- Welche Mechanismen duplizieren, ersetzen oder stören Access-native Formularlogik?

Kurzfazit:
- Das Framework ist stark bei Navigation, CommandBar, Translation, Logging, Tenant-/Datenbankkontext und fachlicher Service-/Repository-Trennung.
- Die grössten Architekturspannungen entstehen dort, wo Workspace und Formulare Access-native Initialisierung nachträglich rekonstruieren:
  - `SourceObject`-basierter Formwechsel
  - verspätete Kontextübergabe über Pending Payload
  - Runtime-RecordSource-Umschaltung
  - Cleanup/Destruktion in `CanLeaveWorkspace`
  - Runtime-Event-Wiring im Formular
- Eine schrittweise Vereinfachung ist möglich. Eine Big-Bang-Neuschreibung ist nicht notwendig.

Wichtigste Befunde:
1. `frmAppShell` als globaler Host ist sinnvoll und architektonisch klar.
2. Die zentrale CommandBar ist sinnvoll, solange sie nur delegiert.
3. `modAppWorkspaceService` übernimmt aktuell teils echte Host-Verantwortung und teils formularinterne Initialisierung.
4. Das Pending-OpenArgs-System kompensiert ein natives Access-Limit des `SourceObject`-Hostings, führt aber zu verspäteter Initialisierung nach `Form_Load`.
5. `CanLeaveWorkspace` wird in mehreren Formularen nicht nur als Frage verwendet, sondern löst Seiteneffekte aus oder bereitet sie vor.
6. `frmOrderDetailNext` zeigt exemplarisch, wie stark die Formlogik inzwischen um Workspace-/Payload-Probleme herumgebaut wurde.

Empfohlene Stossrichtung:
1. Shell/Form-Verantwortung explizit trennen.
2. `CanLeaveWorkspace` auf reine Entscheidungslogik reduzieren.
3. Cleanup in eigene Lifecycle-Stufe verschieben.
4. Pending Payload nur als technische Brücke tolerieren, nicht als generischen Kontextstandard ausbauen.
5. Runtime-RecordSource- und Subform-SourceObject-Manipulation minimieren.
6. Neue oder kritische Detailformulare stärker an Access-native Binding-Regeln ausrichten.

Siehe:
- `app-shell-responsibility-map.md`
- `access-native-vs-framework.csv`
- `workspace-runtime-flow.md`
- `workspace-openargs-analysis.md`
- `app-shell-target-architecture.md`
- `app-shell-simplification-plan.md`
