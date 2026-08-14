import sys
import openpyxl

KNOWN_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}

def scan_workbook(filepath):
    print(f"Scanning {filepath} for Excel formula errors...")

    # Load twice to map formulas to cached values
    wb_formulas = openpyxl.load_workbook(filepath, data_only=False)
    wb_values = openpyxl.load_workbook(filepath, data_only=True)

    errors_found = []
    formulas_scanned = 0

    for sheet_name in wb_formulas.sheetnames:
        ws_f = wb_formulas[sheet_name]
        ws_v = wb_values[sheet_name]

        for row in ws_f.iter_rows():
            is_empty_data_row = True
            for c in row:
                if c.data_type != 'f' and not (isinstance(c.value, str) and c.value.startswith('=')):
                    if c.value not in (None, ""):
                        is_empty_data_row = False
                        break

            for cell_f in row:
                if cell_f.data_type == 'f' or (isinstance(cell_f.value, str) and cell_f.value.startswith('=')):
                    formulas_scanned += 1
                    cell_v = ws_v[cell_f.coordinate]

                    if isinstance(cell_v.value, str) and cell_v.value in KNOWN_ERRORS:
                        if is_empty_data_row:
                            continue

                        errors_found.append({
                            "sheet": sheet_name,
                            "cell": cell_f.coordinate,
                            "formula": cell_f.value,
                            "result": cell_v.value
                        })

    if errors_found:
        print(f"\nFOUND {len(errors_found)} ERRORS in {filepath}:")
        for err in errors_found:
            print(f"{filepath} | {err['sheet']}!{err['cell']}")
            print(f"formula: {err['formula']}")
            print(f"result: {err['result']}\n")
        return False, formulas_scanned, len(errors_found)

    print(f"Scan complete. {formulas_scanned} formulas checked. 0 errors.")
    return True, formulas_scanned, 0

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: scan_excel_errors.py <workbook.xlsx>")
        sys.exit(1)

    filepath = sys.argv[1]
    success, scanned, errors = scan_workbook(filepath)
    if not success:
        sys.exit(1)
    sys.exit(0)
