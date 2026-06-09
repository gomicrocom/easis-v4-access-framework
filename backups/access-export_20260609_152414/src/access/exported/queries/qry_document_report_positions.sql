SELECT
    p.document_position_id AS document_position_id,
    p.document_id AS document_id,
    Nz (p.line_no, 0) AS line_no,
    '' AS item_no,
    Nz (p.description, '') AS description,
    Nz (p.quantity, 0) AS quantity,
    Nz (p.unit_code, '') AS unit_code,
    Nz (p.unit_price, 0) AS unit_price,
    Nz (p.vat_rate, 0) AS vat_rate,
    Nz (p.discount_type, 'NONE') AS discount_type,
    Nz (p.discount_value, 0) AS discount_value,
    Nz (p.line_discount_amount, 0) AS line_discount_amount,
    Nz (p.surcharge_type, 'NONE') AS surcharge_type,
    Nz (p.surcharge_value, 0) AS surcharge_value,
    Nz (p.line_base_amount, 0) AS line_base_amount,
    Nz (p.line_surcharge_amount, 0) AS line_surcharge_amount,
    Nz (p.line_total_net, 0) AS net_amount,
    Nz (p.line_total_vat, 0) AS vat_amount,
    Nz (p.line_total_gross, 0) AS gross_amount
FROM
    doc_document_position AS p
ORDER BY
    document_id,
    line_no,
    document_position_id;