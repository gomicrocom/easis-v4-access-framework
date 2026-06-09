SELECT
    d.document_id AS document_id,
    d.document_no AS document_no,
    d.document_date AS document_date,
    d.document_type_code AS document_type,
    d.document_status_code AS status,
    d.vat_mode AS vat_mode,
    d.customer_address_id AS customer_address_id,
    d.customer_name AS customer_name,
    Nz (
        (
            SELECT
                TOP 1 CStr (param_value)
            FROM
                ten_parameter
            WHERE
                param_key = 'TENANT_NAME'
        ),
        ''
    ) AS tenant_name,
    Nz (
        (
            SELECT
                TOP 1 CStr (param_value)
            FROM
                ten_parameter
            WHERE
                param_key = 'TENANT_STREET'
        ),
        ''
    ) AS tenant_street,
    Nz (
        (
            SELECT
                TOP 1 CStr (param_value)
            FROM
                ten_parameter
            WHERE
                param_key = 'TENANT_ZIP_CODE'
        ),
        ''
    ) AS tenant_zip_code,
    Nz (
        (
            SELECT
                TOP 1 CStr (param_value)
            FROM
                ten_parameter
            WHERE
                param_key = 'TENANT_CITY'
        ),
        ''
    ) AS tenant_city,
    Nz (
        (
            SELECT
                TOP 1 CStr (param_value)
            FROM
                ten_parameter
            WHERE
                param_key = 'TENANT_COUNTRY_CODE'
        ),
        ''
    ) AS tenant_country_code,
    Nz (
        (
            SELECT
                TOP 1 CStr (param_value)
            FROM
                ten_parameter
            WHERE
                param_key = 'TENANT_PHONE'
        ),
        ''
    ) AS tenant_phone,
    Nz (
        (
            SELECT
                TOP 1 CStr (param_value)
            FROM
                ten_parameter
            WHERE
                param_key = 'TENANT_EMAIL'
        ),
        ''
    ) AS tenant_email,
    Nz (
        (
            SELECT
                TOP 1 CStr (param_value)
            FROM
                ten_parameter
            WHERE
                param_key = 'TENANT_VAT_NO'
        ),
        ''
    ) AS tenant_vat_no,
    IIf(
        UCase (
            Nz (
                (
                    SELECT
                        TOP 1 CStr (param_value)
                    FROM
                        ten_parameter
                    WHERE
                        param_key = 'ADDRESS_WINDOW_POSITION'
                ),
                ''
            )
        ) = 'RIGHT',
        'RIGHT',
        'LEFT'
    ) AS address_window_position,
    Trim(
        Nz (
            (
                SELECT
                    TOP 1 CStr (param_value)
                FROM
                    ten_parameter
                WHERE
                    param_key = 'TENANT_NAME'
            ),
            ''
        ) & IIf(
            Nz (
                (
                    SELECT
                        TOP 1 CStr (param_value)
                    FROM
                        ten_parameter
                    WHERE
                        param_key = 'TENANT_NAME'
                ),
                ''
            ) <> ''
            AND Nz (
                (
                    SELECT
                        TOP 1 CStr (param_value)
                    FROM
                        ten_parameter
                    WHERE
                        param_key = 'TENANT_STREET'
                ),
                ''
            ) <> '',
            ' | ',
            ''
        ) & Nz (
            (
                SELECT
                    TOP 1 CStr (param_value)
                FROM
                    ten_parameter
                WHERE
                    param_key = 'TENANT_STREET'
            ),
            ''
        ) & IIf(
            (
                Nz (
                    (
                        SELECT
                            TOP 1 CStr (param_value)
                        FROM
                            ten_parameter
                        WHERE
                            param_key = 'TENANT_NAME'
                    ),
                    ''
                ) <> ''
                OR Nz (
                    (
                        SELECT
                            TOP 1 CStr (param_value)
                        FROM
                            ten_parameter
                        WHERE
                            param_key = 'TENANT_STREET'
                    ),
                    ''
                ) <> ''
            )
            AND (
                Nz (
                    (
                        SELECT
                            TOP 1 CStr (param_value)
                        FROM
                            ten_parameter
                        WHERE
                            param_key = 'TENANT_ZIP_CODE'
                    ),
                    ''
                ) <> ''
                OR Nz (
                    (
                        SELECT
                            TOP 1 CStr (param_value)
                        FROM
                            ten_parameter
                        WHERE
                            param_key = 'TENANT_CITY'
                    ),
                    ''
                ) <> ''
            ),
            ' | ',
            ''
        ) & Trim(
            Nz (
                (
                    SELECT
                        TOP 1 CStr (param_value)
                    FROM
                        ten_parameter
                    WHERE
                        param_key = 'TENANT_ZIP_CODE'
                ),
                ''
            ) & ' ' & Nz (
                (
                    SELECT
                        TOP 1 CStr (param_value)
                    FROM
                        ten_parameter
                    WHERE
                        param_key = 'TENANT_CITY'
                ),
                ''
            )
        )
    ) AS sender_line,
    IIf(
        Nz (a.company_name, '') <> '',
        a.company_name,
        Nz (d.customer_name, '')
    ) AS customer_display_name,
    Nz (a.street, '') AS customer_street,
    Nz (a.zip_code, '') AS customer_postal_code,
    Nz (a.city, '') AS customer_city,
    Nz (a.country_code, '') AS customer_country,
    IIf(
        Nz (a.company_name, '') <> '',
        a.company_name,
        Nz (d.customer_name, '')
    ) AS billing_display_name,
    Nz (a.street, '') AS billing_street,
    Nz (a.zip_code, '') AS billing_zip_code,
    Nz (a.city, '') AS billing_city,
    Nz (a.country_code, '') AS billing_country_code,
    Trim(Nz (a.zip_code, '') & ' ' & Nz (a.city, '')) AS billing_zip_city,
    Nz (a.country_code, '') AS billing_country,
    '' AS delivery_display_name,
    '' AS delivery_street,
    '' AS delivery_zip_code,
    '' AS delivery_city,
    '' AS delivery_country_code,
    '' AS delivery_zip_city,
    '' AS delivery_country,
    False AS has_delivery_address,
    IIf(
        UCase (Nz (d.document_type_code, '')) = 'INVOICE',
        'Rechnung',
        IIf(
            UCase (Nz (d.document_type_code, '')) = 'QUOTE',
            'Offerte',
            IIf(
                UCase (Nz (d.document_type_code, '')) = 'OFFER',
                'Offerte',
                IIf(
                    UCase (Nz (d.document_type_code, '')) = 'REMINDER',
                    'Mahnung',
                    IIf(
                        UCase (Nz (d.document_type_code, '')) = 'CREDIT',
                        'Gutschrift',
                        IIf(
                            UCase (Nz (d.document_type_code, '')) = 'CREDIT_NOTE',
                            'Gutschrift',
                            Nz (d.document_type_code, '')
                        )
                    )
                )
            )
        )
    ) AS document_title,
    '' AS payment_terms,
    NULL AS due_date,
    CCur (Nz (d.subtotal_net_amount, 0)) AS subtotal_net_amount,
    CCur (Nz (d.header_discount_amount, 0)) AS header_discount_amount,
    CCur (Nz (d.header_surcharge_amount, 0)) AS header_surcharge_amount,
    CCur (Nz (d.total_net, 0)) AS net_amount,
    CCur (Nz (d.total_vat, 0)) AS vat_amount,
    CCur (Nz (d.total_gross, 0)) AS gross_amount,
    d.created_at AS created_at,
    NULL AS posted_at
FROM
    doc_document AS d
    LEFT JOIN adr_address AS a ON d.customer_address_id = a.address_id;