SELECT 1 AS join_key, Max(IIf(param_key='TENANT_NAME',param_value,Null)) AS tenant_name, Max(IIf(param_key='TENANT_STREET',param_value,Null)) AS tenant_street, Max(IIf(param_key='TENANT_ZIP_CODE',param_value,Null)) AS tenant_zip_code, Max(IIf(param_key='TENANT_CITY',param_value,Null)) AS tenant_city
FROM ten_parameter;
