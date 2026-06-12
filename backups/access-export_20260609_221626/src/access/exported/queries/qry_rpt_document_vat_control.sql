SELECT p.document_id, p.vat_rate, p.line_no, p.description, p.quantity, p.unit_code, p.unit_price, Nz(p.line_discount_amount,0) AS line_discount_amount, p.line_total_net, p.line_total_vat, p.line_total_gross
FROM doc_document_position AS p
ORDER BY p.document_id, p.vat_rate, p.line_no;
