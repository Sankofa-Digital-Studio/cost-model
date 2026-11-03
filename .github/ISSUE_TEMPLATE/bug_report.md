---
name: Bug Report
about: Create a report to help us fix formula errors, calculation issues, or script failures.
title: "[BUG]: "
labels: ["bug"]
assignees: ""
---

### Bug Description
A clear and concise description of what the formula error or script failure is.

### Reproduction Steps
1. Run command: `python build_universal_costing_model.py ...`
2. Open generated workbook: `...`
3. Inspect sheet: `...`, cell: `...`
4. Notice formula error: (`#REF!`, `#VALUE!`, `#NAME?`, or incorrect sum)

### Expected Behavior
A clear description of what the formula should evaluate to and the mathematical basis.

### Audit Log Output
Run `python verify_universal_model.py <workbook.xlsx>` and paste the terminal output below:
```text
(paste audit log here)
```

### Environment
- OS: [e.g. Windows 11, Ubuntu 22.04]
- Python Version: [e.g. 3.11.2]
- OpenPyXL Version: [e.g. 3.1.2]
