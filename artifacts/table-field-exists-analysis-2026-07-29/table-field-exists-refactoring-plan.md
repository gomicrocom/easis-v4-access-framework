# TableExists / FieldExists Refactoring Plan

## Counts
- Actual TableExists definitions: 20
- Actual FieldExists definitions: 7

## Actual call counts per definition
- TableExists | modAddressRepository:205-228 | 1 call(s)
- TableExists | modAppNavigationService:668-686 | 2 call(s)
- FieldExists | modAppNavigationService:688-709 | 1 call(s)
- TableExists | modArticleGroupService:338-355 | 1 call(s)
- FieldExists | modArticleGroupService:357-377 | 1 call(s)
- TableExists | modBasicModuleSchema:1157-1172 | 10 call(s)
- FieldExists | modBasicModuleSchema:1174-1192 | 20 call(s)
- TableExists | modContactRepository:181-204 | 1 call(s)
- TableExists | modDemoDataSeeder:1269-1285 | 6 call(s)
- FieldExists | modDemoDataSeeder:1287-1297 | 4 call(s)
- TableExists | modDocumentCalculationService:576-592 | 2 call(s)
- FieldExists | modDocumentCalculationService:594-613 | 2 call(s)
- TableExists | modDocumentRepository:650-674 | 10 call(s)
- TableExists | modFwComposerService:1370-1390 | 2 call(s)
- TableExists | modFwTranslationAuditService:931-946 | 10 call(s)
- TableExists | modFwTranslationEditService:457-471 | 1 call(s)
- TableExists | modFwTranslationRuntime:681-701 | 2 call(s)
- TableExists | modFwTranslationTagGeneratorService:555-575 | 3 call(s)
- TableExists | modMigrationPaymentTerms:413-433 | 9 call(s)
- FieldExists | modMigrationPaymentTerms:435-462 | 3 call(s)
- TableExists | modMigrationTranslations:266-286 | 2 call(s)
- FieldExists | modMigrationTranslations:288-315 | 3 call(s)
- TableExists | modNumberRangeRepository:137-160 | 1 call(s)
- TableExists | modOrderRepository:1362-1377 | 1 call(s)
- TableExists | modOutputPathService:253-284 | 1 call(s)
- TableExists | modTenantRepository:102-126 | 1 call(s)
- TableExists | modUserRepository:136-160 | 1 call(s)

## Global identical implementation groups
- TableExists-EXACT-001: TableExists | EXACT | 6 definition(s) | modAddressRepository; modContactRepository; modDocumentRepository; modNumberRangeRepository; modTenantRepository; modUserRepository
- TableExists-EXACT-002: TableExists | EXACT | 2 definition(s) | modBasicModuleSchema; modOrderRepository
- TableExists-EXACT-003: TableExists | EXACT | 5 definition(s) | modFwComposerService; modFwTranslationRuntime; modFwTranslationTagGeneratorService; modMigrationPaymentTerms; modMigrationTranslations
- TableExists-SAME-SIGNATURE-DIFF-001: TableExists | SAME_SIGNATURE_DIFFERENT_BODY | 8 definition(s) | modAddressRepository; modArticleGroupService; modContactRepository; modDocumentRepository; modNumberRangeRepository; modOutputPathService; modTenantRepository; modUserRepository
- TableExists-SAME-SIGNATURE-DIFF-002: TableExists | SAME_SIGNATURE_DIFFERENT_BODY | 11 definition(s) | modBasicModuleSchema; modDemoDataSeeder; modDocumentCalculationService; modFwComposerService; modFwTranslationAuditService; modFwTranslationEditService; modFwTranslationRuntime; modFwTranslationTagGeneratorService; modMigrationPaymentTerms; modMigrationTranslations; modOrderRepository
- FieldExists-EXACT-001: FieldExists | EXACT | 2 definition(s) | modMigrationPaymentTerms; modMigrationTranslations
- FieldExists-SAME-SIGNATURE-DIFF-001: FieldExists | SAME_SIGNATURE_DIFFERENT_BODY | 5 definition(s) | modBasicModuleSchema; modDemoDataSeeder; modDocumentCalculationService; modMigrationPaymentTerms; modMigrationTranslations

## Functional differences
- Two main API families exist for each helper: CurrentDb-based and explicit DAO.Database-based.
- The explicit DAO.Database variants are the better base for a later shared schema API because they avoid hidden ambient database state.
- The CurrentDb variants are simpler for repository callers but couple the helper to the ambient frontend context.
- Parameter naming differences (tableName vs table_name, fieldName vs field_name) are cosmetic and should not drive separate long-term APIs.

## Recommended public signatures
- Preferred: Public Function TableExists(ByVal tableName As String, ByVal db As DAO.Database) As Boolean
- Preferred: Public Function FieldExists(ByVal tableName As String, ByVal fieldName As String, ByVal db As DAO.Database) As Boolean
- Assessment: an Optional DAO.Database parameter is technically possible in VBA only via an object/variant pattern, but it is less explicit and less readable than a required DAO.Database argument.
- Recommended CurrentDb caller pattern: caller resolves CurrentDb explicitly and passes db into the canonical helper.

## Recommended target module
- Preferred target module: modDbSchema
- Reason: both helpers inspect schema metadata and belong beside other field/index/table inspection responsibilities rather than in a broad generic db-access module.

## Later removal candidates
- modAddressRepository:205-228 Private Function TableExists(ByVal tableName As String) As Boolean
- modAppNavigationService:668-686 Private Function TableExists(ByVal table_name As String) As Boolean
- modAppNavigationService:688-709 Private Function FieldExists(ByVal table_name As String, ByVal field_name As String) As Boolean
- modArticleGroupService:338-355 Private Function TableExists(ByVal tableName As String) As Boolean
- modArticleGroupService:357-377 Private Function FieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
- modBasicModuleSchema:1157-1172 Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
- modBasicModuleSchema:1174-1192 Private Function FieldExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) As Boolean
- modContactRepository:181-204 Private Function TableExists(ByVal tableName As String) As Boolean
- modDemoDataSeeder:1269-1285 Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
- modDemoDataSeeder:1287-1297 Private Function FieldExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) As Boolean
- modDocumentCalculationService:576-592 Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
- modDocumentCalculationService:594-613 Private Function FieldExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) As Boolean
- modDocumentRepository:650-674 Private Function TableExists(ByVal tableName As String) As Boolean
- modFwComposerService:1370-1390 Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
- modFwTranslationAuditService:931-946 Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
- modFwTranslationEditService:457-471 Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
- modFwTranslationRuntime:681-701 Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
- modFwTranslationTagGeneratorService:555-575 Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
- modMigrationPaymentTerms:413-433 Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
- modMigrationPaymentTerms:435-462 Private Function FieldExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) As Boolean
- modMigrationTranslations:266-286 Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
- modMigrationTranslations:288-315 Private Function FieldExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) As Boolean
- modNumberRangeRepository:137-160 Private Function TableExists(ByVal tableName As String) As Boolean
- modOrderRepository:1362-1377 Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
- modOutputPathService:253-284 Private Function TableExists(ByVal tableName As String) As Boolean
- modTenantRepository:102-126 Private Function TableExists(ByVal tableName As String) As Boolean
- modUserRepository:136-160 Private Function TableExists(ByVal tableName As String) As Boolean

## Later call-site changes
- Module modAddressRepository: local calls to TableExists should later be redirected either through explicit CurrentDb resolution or an agreed wrapper policy.
- Module modAppNavigationService: local calls to FieldExists should later be redirected either through explicit CurrentDb resolution or an agreed wrapper policy.
- Module modAppNavigationService: local calls to TableExists should later be redirected either through explicit CurrentDb resolution or an agreed wrapper policy.
- Module modArticleGroupService: local calls to FieldExists should later be redirected either through explicit CurrentDb resolution or an agreed wrapper policy.
- Module modArticleGroupService: local calls to TableExists should later be redirected either through explicit CurrentDb resolution or an agreed wrapper policy.
- Module modBasicModuleSchema: local calls to FieldExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modBasicModuleSchema: local calls to TableExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modContactRepository: local calls to TableExists should later be redirected either through explicit CurrentDb resolution or an agreed wrapper policy.
- Module modDemoDataSeeder: local calls to FieldExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modDemoDataSeeder: local calls to TableExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modDocumentCalculationService: local calls to FieldExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modDocumentCalculationService: local calls to TableExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modDocumentRepository: local calls to TableExists should later be redirected either through explicit CurrentDb resolution or an agreed wrapper policy.
- Module modFwComposerService: local calls to TableExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modFwTranslationAuditService: local calls to TableExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modFwTranslationEditService: local calls to TableExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modFwTranslationRuntime: local calls to TableExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modFwTranslationTagGeneratorService: local calls to TableExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modMigrationPaymentTerms: local calls to FieldExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modMigrationPaymentTerms: local calls to TableExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modMigrationTranslations: local calls to FieldExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modMigrationTranslations: local calls to TableExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modNumberRangeRepository: local calls to TableExists should later be redirected either through explicit CurrentDb resolution or an agreed wrapper policy.
- Module modOrderRepository: local calls to TableExists should later be redirected to a shared schema helper with explicit DAO.Database.
- Module modOutputPathService: local calls to TableExists should later be redirected either through explicit CurrentDb resolution or an agreed wrapper policy.
- Module modTenantRepository: local calls to TableExists should later be redirected either through explicit CurrentDb resolution or an agreed wrapper policy.
- Module modUserRepository: local calls to TableExists should later be redirected either through explicit CurrentDb resolution or an agreed wrapper policy.

## Risks
- Private local helpers currently own unqualified calls inside their own modules; future migration must retarget those calls deliberately.
- CurrentDb-based and explicit-db-based call sites should not be merged blindly without checking backend-routing and transaction expectations.
- Any shared helper must preserve current False-on-missing behavior and existing early-exit/error-handling assumptions.
- Schema-sensitive modules that work against backend databases should keep explicit DAO.Database flow to avoid regressions.
