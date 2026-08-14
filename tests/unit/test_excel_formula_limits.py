import pytest
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.templates.finance_personal.template import FinancePersonalTemplate

def test_global_excel_formula_length_limit():
    """
    Ensure no formula generated in the production template exceeds Excel's 8192 character limit.
    """
    req = GenerationRequest(
        template_id="finance_personal",
        year=2026,
        theme="light",
        with_sample_data=False,
        projection_horizon=12
    )
    template = FinancePersonalTemplate()
    plan = template.build_workbook_plan(req)

    longest_length = 0
    longest_location = ""
    
    for ws in plan.worksheets:
        # Check cell formulas
        for cell in ws.cells:
            if cell.formula is not None:
                formula_text = str(cell.formula)
                if not formula_text.startswith("="):
                    formula_text = "=" + formula_text
                
                length = len(formula_text)
                location = f"{ws.name} / Cell({cell.row}, {cell.col})"
                
                if length > longest_length:
                    longest_length = length
                    longest_location = location
                
                if length > 8192:
                    pytest.fail(f"Excel formula exceeds 8192 characters:\n{location}: {length}")

        # Check table column formulas
        for tbl in ws.tables:
            for col in tbl.columns:
                if col.formula is not None:
                    formula_text = str(col.formula)
                    if not formula_text.startswith("="):
                        formula_text = "=" + formula_text
                    
                    length = len(formula_text)
                    location = f"{ws.name} / {tbl.name} / {col.header}"
                    
                    if length > longest_length:
                        longest_length = length
                        longest_location = location
                    
                    if length > 8192:
                        pytest.fail(f"Excel formula exceeds 8192 characters:\n{location}: {length}")
    
    print(f"Global formula budget check passed.")
    print(f"Longest formula: {longest_location} ({longest_length} characters)")
