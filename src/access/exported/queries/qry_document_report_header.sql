SELECT
    d.document_id AS document_id,
    d.document_no AS document_no,
    d.document_date AS document_date,
    d.document_type_code AS document_type,
    d.document_status_code AS status,
    d.customer_address_id AS customer_address_id,
    d.customer_name AS customer_name,
    IIf(
        Nz (a.company_name, '') <> '',
        a.company_name,
        Nz (d.customer_name, '')
    ) AS customer_display_name,
    Nz (a.street, '') AS customer_street,
    Nz (a.zip_code, '') AS customer_postal_code,
    Nz (a.city, '') AS customer_city,
    Nz (a.country_code, '') AS customer_country,
    '' AS payment_terms,
    NULL AS due_date,
    Nz (d.subtotal_net_amount, 0) AS subtotal_net_amount,
    Nz (d.header_discount_amount, 0) AS header_discount_amount,
    Nz (d.header_surcharge_amount, 0) AS header_surcharge_amount,
    Nz (d.total_net, 0) AS net_amount,
    Nz (d.total_vat, 0) AS vat_amount,
    Nz (d.total_gross, 0) AS gross_amount,
    d.created_at AS created_at,
    NULL AS posted_at
FROM
    doc_document AS d
    LEFT JOIN adr_address AS a ON d.customer_address_id = a.address_id;