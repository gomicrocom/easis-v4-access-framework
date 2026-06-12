SELECT fw_translation.translation_id, fw_translation.translation_key, fw_translation.language_code, fw_translation.translation_value
FROM fw_translation
WHERE (((fw_translation.translation_value)="<neu>"));
