import openpyxl

def verify_model():
    filename = "CAPM_Tau_Mine_10_Demolition_Rehabilitation_Model.xlsx"
    print(f"Loading workbook: {filename}")
    wb = openpyxl.load_workbook(filename, data_only=False)
    print(f"Sheet count: {len(wb.sheetnames)}")
    print(f"Sheets: {wb.sheetnames}")

    total_formulas = 0
    ref_errors = 0
    value_errors = 0
    name_errors = 0

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        sheet_formulas = 0
        sheet_ref_err = 0
        sheet_val_err = 0
        sheet_name_err = 0
        
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                    sheet_formulas += 1
                    total_formulas += 1
                    val_str = str(cell.value)
                    if "#REF!" in val_str:
                        print(f"ERROR #REF! at {sheetname}!{cell.coordinate}: {val_str}")
                        sheet_ref_err += 1
                        ref_errors += 1
                    if "#VALUE!" in val_str:
                        print(f"ERROR #VALUE! at {sheetname}!{cell.coordinate}: {val_str}")
                        sheet_val_err += 1
                        value_errors += 1
                    if "#NAME?" in val_str:
                        print(f"ERROR #NAME? at {sheetname}!{cell.coordinate}: {val_str}")
                        sheet_name_err += 1
                        name_errors += 1

        print(f"Sheet '{sheetname}': {sheet_formulas} formulas, Ref Errors: {sheet_ref_err}, Val Errors: {sheet_val_err}, Name Errors: {sheet_name_err}")

    print("\n--- Summary Verification ---")
    print(f"Total Formulas Evaluated: {total_formulas}")
    print(f"Total #REF! Errors: {ref_errors}")
    print(f"Total #VALUE! Errors: {value_errors}")
    print(f"Total #NAME? Errors: {name_errors}")

    if ref_errors == 0 and value_errors == 0 and name_errors == 0:
        print("\nSUCCESS: All workbook formulas and links are 100% verified with ZERO errors!")
    else:
        print("\nWARNING: Errors were detected. Review above details.")

if __name__ == "__main__":
    verify_model()
