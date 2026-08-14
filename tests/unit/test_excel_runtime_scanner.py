import os
import pytest
import xlsxwriter
from scripts.excel_runtime.scan_excel_errors import scan_workbook

def create_test_workbook(path, scenario):
    """
    Creates a workbook to test scanner behaviors.
    """
    workbook = xlsxwriter.Workbook(path)
    
    if scenario == "A":
        # A. Non-table formula error (worksheet formula cell has cached #VALUE!)
        ws = workbook.add_worksheet("Test")
        # xlsxwriter writes <v>#VALUE!</v> when value="#VALUE!" is passed
        ws.write_formula("A1", "=1/0", None, "#VALUE!")
        
    elif scenario == "B":
        # B. Dashboard-style formula-only row (cached #REF! or #VALUE!)
        ws = workbook.add_worksheet("Dashboard")
        # Empty row except for formulas
        ws.write_formula("A2", "=1/0", None, "#VALUE!")
        ws.write_formula("C2", "=A1+1", None, "#REF!")
        
    elif scenario == "C":
        # C. Empty table placeholder row (formula in table data row, all other cells blank)
        ws = workbook.add_worksheet("TableSheet")
        ws.write("A1", "Col1")
        ws.write("B1", "Col2")
        ws.write("C1", "FormulaCol")
        
        ws.add_table("A1:C2", {
            "name": "TestTable",
            "columns": [
                {"header": "Col1"},
                {"header": "Col2"},
                {"header": "FormulaCol"}
            ]
        })
        # Write after add_table so it doesn't overwrite our cached error value
        ws.write_formula("C2", "=A2+B2", None, "#VALUE!")
        
    elif scenario == "D":
        # D. Populated table data row
        ws = workbook.add_worksheet("TableSheet")
        ws.write("A1", "Col1")
        ws.write("B1", "Col2")
        ws.write("C1", "FormulaCol")
        
        ws.add_table("A1:C2", {
            "name": "TestTable",
            "columns": [
                {"header": "Col1"},
                {"header": "Col2"},
                {"header": "FormulaCol"}
            ]
        })
        
        # Row 2 is the data row. Col1 is populated.
        ws.write("A2", 10)
        # Write after add_table so it doesn't overwrite our cached error value
        ws.write_formula("C2", "=A2+B2", None, "#VALUE!")

    workbook.close()

def test_scanner_scenario_a(tmp_path):
    path = tmp_path / "scenario_a.xlsx"
    create_test_workbook(str(path), "A")
    success, scanned, errors = scan_workbook(str(path))
    assert not success, "Scanner MUST FAIL on non-table formula error"
    assert errors == 1

def test_scanner_scenario_b(tmp_path):
    path = tmp_path / "scenario_b.xlsx"
    create_test_workbook(str(path), "B")
    success, scanned, errors = scan_workbook(str(path))
    assert not success, "Scanner MUST FAIL on dashboard-style formula row error"
    assert errors == 2

def test_scanner_scenario_c(tmp_path):
    path = tmp_path / "scenario_c.xlsx"
    create_test_workbook(str(path), "C")
    success, scanned, errors = scan_workbook(str(path))
    assert success, "Scanner MAY ignore empty table placeholder row errors"
    assert errors == 0

def test_scanner_scenario_d(tmp_path):
    path = tmp_path / "scenario_d.xlsx"
    create_test_workbook(str(path), "D")
    success, scanned, errors = scan_workbook(str(path))
    assert not success, "Scanner MUST FAIL on populated table data row error"
    assert errors == 1
