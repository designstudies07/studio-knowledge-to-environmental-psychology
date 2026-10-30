# Restricted local inputs

This directory is intentionally excluded from version control. To reproduce the full analysis locally, provide:

- `Three_Stage_Analysis_Single_Researcher.xlsx`
- `Pre-Lecture_Test_Responses.xlsx`
- `hash_salt.txt`, containing the confidential linkage salt, or set `STUDY_LINKAGE_SALT`
- optionally, `name_aliases.json` for documented local-only name normalization

Example alias-file structure:

```json
{
  "normalized_variant": "normalized_canonical_value"
}
```

Never commit these files. The analysis workbook contains pseudonymous participant-level records, and the T0 export and linkage materials contain information that can enable re-identification.
