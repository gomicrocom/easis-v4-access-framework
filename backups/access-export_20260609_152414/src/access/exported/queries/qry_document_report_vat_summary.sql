SELECT
    p.document_id AS document_id,
    Nz(p.vat_rate, 0) AS vat_rate,
    Sum(Nz(p.line_total_net, 0)) AS vat_base_amount,
    Sum(Nz(p.line_total_vat, 0)) AS vat_amount,
    Sum(Nz(p.line_total_gross, 0)) AS gross_amount
FROM
    doc_document_position AS p
GROUP BY
    p.document_id,
    Nz(p.vat_rate, 0)
ORDER BY
    p.document_id,
    Nz(p.vat_rate, 0);
