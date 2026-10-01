# Standard template

`standard-party-template.docx` is a sanitized, reusable DOCX template generated
by `scripts/build_standard_template.py`.

It contains only page settings, paragraph styles, fonts, spacing and other
format definitions. It does not contain real organization names, employee
names, meeting content, activity photos or source documents.

Regenerate it after changing the standard formatting profile:

```bash
python scripts/build_standard_template.py --overwrite
```
