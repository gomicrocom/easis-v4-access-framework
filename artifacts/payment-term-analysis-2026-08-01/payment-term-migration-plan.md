# Payment Term Analysis

Generated: 2026-08-01

## Scope
- Live data exported read-only from C:\easis\easis.accdb via Access TransferText
- Repo references from src/access/exported/*
- No productive data was changed

## Current Schema
Current 	en_payment_term columns observed in live export:
$(payment_term_id payment_term_code language_code title terms_text days_net discount_days discount_percent is_default is_active sort_order created_at created_by updated_at updated_by -join ', ')

Current schema definition in repo is created in modBasicModuleSchema.EnsureReferenceLookupTables and mirrored in modMigrationPaymentTerms.

## Proposed Target Schema
$(payment_term_id payment_term_code payment_term_type_code days_net discount_days discount_percent is_default is_active sort_order created_at created_by updated_at updated_by -join ', ')

Fields planned for later removal after data + code migration:
- language_code
- 	itle
- 	erms_text

## Inventory Summary
- Current rows in 	en_payment_term: **9**
- Distinct current payment_term_code values: **7**
- Proposed canonical payment terms: **5**
- Language duplicate groups: **2**
- Code-variant groups: **2**

## Proposed Canonical Codes
- PREPAYMENT
- NET_30
- CASH_DISCOUNT_10_2_NET_30
- SPLIT_50_ORDER_50_DELIVERY
- MILESTONE_50_25_25

Proposed removals / remaps:
- NET30 -> NET_30
- 2D10N30 -> CASH_DISCOUNT_10_2_NET_30

## Live Findings
- PREPAYMENT exists in de-CH and en-US.
- NET_30 exists in de-CH and en-US.
- CASH_DISCOUNT_10_2_NET_30, SPLIT_50_ORDER_50_DELIVERY, and MILESTONE_50_25_25 currently exist only in de-CH.
- NET30 and 2D10N30 are additional active German rows with missing 	erms_text and missing audit metadata.
- is_default=True is only set on the German NET_30 row, which shows the current per-language duplication problem for an otherwise language-independent business flag.

## Translation Plan Summary
- Existing w_translation rows for PAYMENT_TERM.*: **0**
- Translation rows ready to migrate from current texts: **14**
- Translation rows with missing source text: **16**

Supported target languages for migration planning:
- de-CH
- en-US
- r-CH

## Reference Summary
### Live data references
- ord_order currently references CASH_DISCOUNT_10_2_NET_30 and one blank value.
- 	mp_order currently references NET_30 plus multiple blank values.
- doc_document currently references NET_30, CASH_DISCOUNT_10_2_NET_30, PREPAYMENT, and one blank value.
- No live references to NET30 or 2D10N30 were found outside 	en_payment_term itself.

### Repo references
- Active code/data references exist for NET_30, CASH_DISCOUNT_10_2_NET_30, and PREPAYMENT in modDemoDataSeeder and order/detail UI code.
- No repo references to NET30 or 2D10N30 were found outside the live data export.

## payment_term_type_code Assessment
No existing payment_term_type_code field or reference logic was found in the exported modules, queries, or forms.

Recommendation:
- add payment_term_type_code during the structural cleanup migration,
- seed it deterministically from the canonical code family,
- do **not** introduce a separate reference table in the same step.

Suggested initial mapping:
- PREPAYMENT -> PREPAYMENT
- NET_30 -> NET
- CASH_DISCOUNT_10_2_NET_30 -> CASH_DISCOUNT
- SPLIT_50_ORDER_50_DELIVERY -> INSTALLMENT
- MILESTONE_50_25_25 -> MILESTONE

## Recommended Migration Order
1. Backup frontend and backend ACCDBs.
2. Export current 	en_payment_term, ord_order, 	mp_order, doc_document, and relevant w_translation rows.
3. Insert missing PAYMENT_TERM.<code>.TITLE / .TERMS rows into w_translation for all reliable existing source texts.
4. Fill missing en-US / r-CH translations manually where no safe source text exists.
5. Add payment_term_type_code to 	en_payment_term.
6. Introduce a unique index on canonical payment_term_code only after duplicates are removed.
7. Update all live foreign-code references from NET30 to NET_30 and 2D10N30 to CASH_DISCOUNT_10_2_NET_30.
8. Collapse 	en_payment_term to one row per canonical code, preserving business parameters, is_active, and chosen default state.
9. Update forms, queries, reports, and document rendering to resolve title/terms via w_translation instead of 	itle / 	erms_text.
10. Remove language_code, 	itle, and 	erms_text only after the runtime and reports no longer depend on them.

## Required Code Changes Later
- rmOrderDetail: payment-term combo must resolve translated display text via w_translation instead of 	en_payment_term.title.
- Document output paths that currently read or persist payment_terms_text must move to translation-based resolution by document language.
- Seeder/migration logic must stop reseeding per-language business rows in 	en_payment_term.
- Translation Audit should include PAYMENT_TERM.* keys as ordinary framework translations.

## Risks
- Existing documents store payment_term_code plus free payment_terms_text; migration must preserve historic readability.
- Current language codes in live order/temp data still contain legacy values like DE-CH and r-FR; payment-term migration should not assume language normalization is already complete everywhere.
- is_default is currently attached to a language row, so collapsing duplicates needs an explicit rule for the surviving canonical row.
- Missing r-CH and several missing en-US texts mean translation seeding cannot be fully automated without business review.

## Rollback / Backup
- Copy C:\easis\easis.accdb
- Copy C:\easis\data\DEFAULT_be.accdb
- Copy C:\easis\data\sys_be.accdb if w_translation is stored there in the active environment
- Re-export 	en_payment_term, ord_order, 	mp_order, doc_document, w_translation before mutation

## Concrete Tests After Migration
1. 	en_payment_term has exactly one row per canonical payment_term_code.
2. rmOrderDetail.cboTenPaymentTerm lists only active canonical codes.
3. ord_order.payment_term_code, 	mp_order.payment_term_code, and doc_document.payment_term_code contain no NET30 or 2D10N30.
4. PAYMENT_TERM.<code>.TITLE and .TERMS resolve correctly for de-CH, en-US, and r-CH.
5. Translation Audit reports missing payment-term translations through the standard framework path.
6. Existing documents remain readable, especially if they rely on stored payment_terms_text.
