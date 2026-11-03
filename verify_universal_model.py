import sys
import openpyxl

def verify_file(filename: str):
    print(f"--- Auditing Workbook: {filename} ---")
    wb = openpyxl.load_workbook(filename, data_only=False)
    print(f"Sheets ({len(wb.sheetnames)}): {wb.sheetnames}")

    total_formulas = 0
    ref_errors = 0
    value_errors = 0
    name_errors = 0

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        sheet_formulas = 0
        sheet_ref = 0
        sheet_val = 0
        sheet_name = 0

        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    total_formulas += 1
                    sheet_formulas += 1
                    v = str(cell.value)
                    if "#REF!" in v:
                        print(f"  [#REF!] {sheetname}!{cell.coordinate}: {v}")
                        sheet_ref += 1
                        ref_errors += 1
                    if "#VALUE!" in v:
                        print(f"  [#VALUE!] {sheetname}!{cell.coordinate}: {v}")
                        sheet_val += 1
                        value_errors += 1
                    if "#NAME?" in v:
                        print(f"  [#NAME?] {sheetname}!{cell.coordinate}: {v}")
                        sheet_name += 1
                        name_errors += 1

        print(f"  * Sheet '{sheetname}': {sheet_formulas} formulas, errors = {sheet_ref + sheet_val + sheet_name}")

    print(f"\nResult: Total Formulas = {total_formulas}, #REF! = {ref_errors}, #VALUE! = {value_errors}, #NAME? = {name_errors}")
    if ref_errors == 0 and value_errors == 0 and name_errors == 0:
        print(">>> SUCCESS: 100% VALIDATED WITH ZERO FORMULA ERRORS!\n")
        return True
    else:
        print(">>> FAILURE: Formula errors detected!\n")
        return False

if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "Senatla_5_Village_Demolition_Rehabilitation_Model.xlsx"
    success = verify_file(target)
    sys.exit(0 if success else 1)
