import os
import zipfile
import re
from excel_saas.core.models.workbook_plan import WorkbookPlan, WorksheetPlan, CellPlan
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.engine.workbook_engine import WorkbookEngine
from excel_saas.themes.light import LightTheme

def test_workbook_engine_formula_cached_value_regression(tmp_path):
    """
    Test that the WorkbookEngine does not serialize `<v>None</v>` when a CellPlan
    with a formula has its `value` set to None (the default).
    Passing an explicit None to `write_formula` in XlsxWriter causes Excel to fail to open the file.
    """
    # Create a minimal workbook plan
    plan = WorkbookPlan(
        worksheets=[
            WorksheetPlan(
                name="TestSheet",
                cells=[
                    # Cell 1: Formula with default None value
                    CellPlan(row=0, col=0, formula=Formula("1+1"), value=None),
                    # Cell 2: Formula with explicit cached value
                    CellPlan(row=0, col=1, formula=Formula("2+2"), value=4)
                ]
            )
        ]
    )

    path = tmp_path / "test_engine.xlsx"
    engine = WorkbookEngine(plan, LightTheme())
    engine.render(str(path))

    # Read the generated XML to assert correctness
    with zipfile.ZipFile(path, 'r') as z:
        xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")

        # Cell A1 (row=0, col=0 in 0-indexed translates to A1, r="A1")
        # Ensure it does NOT contain `<v>None</v>`
        match_a1 = re.search(r'<c r="A1".*?</c>', xml)
        assert match_a1 is not None, "Cell A1 not found in XML"
        cell_a1_xml = match_a1.group(0)

        # It should contain the formula
        assert "<f>1+1</f>" in cell_a1_xml

        # It should NOT contain `<v>None</v>`
        assert "<v>None</v>" not in cell_a1_xml

        # In modern XlsxWriter, passing NO value generates `<v>0</v>`
        assert "<v>0</v>" in cell_a1_xml, "Expected XlsxWriter default 0 cached value"

        # Cell B1 (row=0, col=1 in 0-indexed translates to B1, r="B1")
        match_b1 = re.search(r'<c r="B1".*?</c>', xml)
        assert match_b1 is not None, "Cell B1 not found in XML"
        cell_b1_xml = match_b1.group(0)

        # It should contain the explicit cached value 4
        assert "<f>2+2</f>" in cell_b1_xml
        assert "<v>4</v>" in cell_b1_xml
