SELECT a.address_id, Nz(a.company_name, '') AS company_name, Trim(Nz(a.first_name, '') & ' ' & Nz(a.last_name, '')) AS person_name, IIf(Len(Trim(Nz(a.company_name, ''))) > 0,
        Trim(Nz(a.company_name, '')),
        Trim(Nz(a.first_name, '') & ' ' & Nz(a.last_name, ''))) AS display_name, Trim(Nz(a.street, '') & IIf(Len(Trim(Nz(a.house_no, ''))) > 0, ' ', '') & Nz(a.house_no, '')) AS address_line_1, Trim(Nz(a.zip_code, '') & IIf(Len(Trim(Nz(a.city, ''))) > 0, ' ', '') & Nz(a.city, '')) AS address_line_2, Trim(
        Nz(a.street, '') &
        IIf(Len(Trim(Nz(a.house_no, ''))) > 0, ' ', '') &
        Nz(a.house_no, '') &
        IIf(Len(Trim(Nz(a.zip_code, '') & Nz(a.city, ''))) > 0, ', ', '') &
        Trim(Nz(a.zip_code, '') & IIf(Len(Trim(Nz(a.city, ''))) > 0, ' ', '') & Nz(a.city, ''))
    ) AS address_compact, Nz(a.country_code, '') AS country_code, Nz(a.language_code, 'DE-CH') AS language_code, 'CHF' AS currency_code, IIf(Nz(a.is_active, True), 'ACTIVE', 'INACTIVE') AS status_code, IIf(Nz(a.is_active, True), 'ACTIVE', 'INACTIVE') AS status_text, False AS is_locked, Null AS lock_reason, Nz(a.preferred_name, '') AS contact_info, 0 AS open_invoices_count, 0 AS overdue_items_count, 0 AS dunning_count, 0 AS sales_current_year_amount, 0 AS open_orders_count, 0 AS active_subscriptions_count, 'n/a' AS last_activity_text
FROM adr_address AS a
ORDER BY IIf(Len(Trim(Nz(a.company_name, ''))) > 0,
              Trim(Nz(a.company_name, '')),
              Trim(Nz(a.first_name, '') & ' ' & Nz(a.last_name, '')));
