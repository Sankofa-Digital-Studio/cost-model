## Description
<!-- Provide a concise summary of the change, issue addressed, or new model preset added. -->

## Type of Change
- [ ] New costing model preset (e.g. new mine / plant / municipal scope)
- [ ] Formula or calculation fix
- [ ] Visual styling / OpenPyXL formatting update
- [ ] Documentation update
- [ ] New test / verification audit utility

## Quality & Formula Integrity Checklist
- [ ] I have executed `python verify_universal_model.py <workbook.xlsx>` on all modified or generated workbooks.
- [ ] Total `#REF!` errors = 0
- [ ] Total `#VALUE!` errors = 0
- [ ] Total `#NAME?` errors = 0
- [ ] All formulas use uppercase syntax (`=SUM(...)`, `=ROUND(...)`, etc.).
- [ ] Executive Navy (`#1B365D`) and Gold (`#C5A059`) styling guidelines were followed.
- [ ] No client confidential credentials, private API keys, or unredacted personal information are included.
