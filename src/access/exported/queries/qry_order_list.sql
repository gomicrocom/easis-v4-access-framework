SELECT o.order_id, o.order_no, o.order_type_code, o.order_status_code, o.customer_address_id, o.invoice_address_id, o.delivery_address_id, o.customer_name, o.order_date, o.delivery_date, o.valid_until, o.reference_text, o.external_reference, o.language_code, o.currency_code, o.payment_term_code, o.vat_mode, o.vat_code, o.vat_rate, o.subtotal_net_amount, o.header_discount_amount, o.header_surcharge_amount, o.net_amount, o.vat_amount, o.gross_amount, o.result_document_id, o.updated_at
FROM ord_order AS o
ORDER BY o.order_date DESC , o.order_no DESC;
