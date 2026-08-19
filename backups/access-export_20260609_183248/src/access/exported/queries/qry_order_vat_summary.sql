SELECT
    l.order_id,
    l.vat_code,
    l.vat_rate,
    Sum(Nz(l.line_net_amount, 0)) AS vat_base_amount,
    Sum(Nz(l.line_vat_amount, 0)) AS vat_amount,
    Sum(Nz(l.line_gross_amount, 0)) AS gross_amount
FROM
    ord_order_line AS l
GROUP BY
    l.order_id,
    l.vat_code,
    l.vat_rate
ORDER BY
    l.order_id,
    l.vat_code,
    l.vat_rate;
