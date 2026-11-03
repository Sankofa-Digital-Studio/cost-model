# Contributing to Demolition & Rehabilitation Costing Models

Thank you for contributing to the **Mining & Heavy Industrial Decommissioning, Environmental Rehabilitation & Costing Automation Engine**.

This repository powers commercial tender submissions, bankable feasibility budgets, and engineering bill of quantities (BOQ) models for major mining asset closures. Accuracy, formula integrity, and executive presentation are critical.

---

## Code of Conduct

All contributors are expected to uphold our [Code of Conduct](CODE_OF_CONDUCT.md). Please report any violations to the project maintainers.

---

## Development Workflow

### 1. Branching Strategy
- `main`: Production-ready branch. Releases and verified client models live here.
- `dev`: Active integration branch.
- `feature/<issue-or-topic>`: For new models, features, or generator updates.
- `fix/<issue-or-topic>`: For bug fixes, rate adjustments, or formula corrections.

### 2. Commit Message Standards
We follow the Conventional Commits specification:
```
<type>(<scope>): <short description>
```
**Types**:
- `feat`: A new model preset, sheet structure, or calculation capability.
- `fix`: A formula fix, rate correction, or bug fix.
- `docs`: Documentation updates, README, or engineering notes.
- `style`: Formatting, OpenPyXL color palette, column width adjustments.
- `refactor`: Code reorganization without behavioral changes.
- `test`: Adding or updating verification scripts or formula audits.

**Examples**:
```
feat(generator): add waste rock dump mobile crushing calculation
fix(waterfall): resolve circular reference in contingencies row
docs(readme): add Klipfontein JSON configuration example
```

---

## Executive Workbook Standards

Any script that generates an Excel model (`.xlsx`) must strictly comply with the **Mining Industrial Executive Visual Standard**:

### 1. Color Palette
| Purpose | Hex Code | Usage |
| :--- | :---: | :--- |
| **Primary Navy Dark** | `#1B365D` | Top title banners, grand totals, table headers |
| **Secondary Navy Med** | `#2E5B88` | Section headers and category subheaders |
| **Gold Accent** | `#C5A059` | Subtle brand lines and editable parameter borders |
| **Zebra Striping** | `#F4F6F9` | Alternate row background for clean readability |
| **Subtotal Fill** | `#E9EEF4` | Subtotal rows across all BOQ and phase sheets |
| **Input Yellow** | `#FFF2CC` | All user-editable input cells (with `#D4AF37` border) |
| **Profit Box Green** | `#274E13` | Contractor margin highlight border |
| **Profit Text Red** | `#C00000` | Contractor profit margin percentage font |

### 2. Formula Integrity
- **100% Dynamic Formulas**: Never hardcode calculated values in generated Excel sheets.
- **Formula Casing**: Always use uppercase Excel function names (e.g. `=SUM(...)`, `=ROUND(...)`, `=IF(...)`).
- **Sheet Linking**: Use single-quoted sheet references (e.g. `='Rates & Parameters'!$D$6`).
- **Zero Error Tolerance**: Workbooks must evaluate with **0 `#REF!`**, **0 `#VALUE!`**, and **0 `#NAME?`** errors.

---

## Quality Gate (Mandatory Before Submitting a PR)

Before creating a pull request, run all verification scripts:

```bash
# 1. Verify Universal Models
python build_universal_costing_model.py --preset village
python verify_universal_model.py Senatla_5_Village_Demolition_Rehabilitation_Model.xlsx

python build_universal_costing_model.py --preset westgold
python verify_universal_model.py West_Gold_Plant_Universal_Model.xlsx

# 2. Verify Legacy Workbooks
python verify_model.py
python verify_capm_tau_model.py
```

Ensure all verification scripts exit with:
`SUCCESS: 100% VALIDATED WITH ZERO FORMULA ERRORS!`

---

## Pull Request Process

1. Fork or branch from `dev`.
2. Follow the workbook and code standards.
3. Verify all workbooks with the verification tools.
4. Fill out the [Pull Request Template](.github/PULL_REQUEST_TEMPLATE.md).
5. Ensure continuous integration (CI) tests pass.
6. A maintainer will review and approve your submission.
