import os
import openpyxl
from excel_saas.application.generate_workbook import generate
from excel_saas.core.models.generation_request import GenerationRequest

def test_full_pipeline_light_and_dark(tmp_path):
    output_dir = str(tmp_path)

    req_light = GenerationRequest(template_id="finance_personal", year=2026, theme="light")
    path_light = generate(req_light, output_dir)

    assert os.path.exists(path_light)

    # Verify via openpyxl
    wb = openpyxl.load_workbook(path_light, data_only=False)
    sheet_names = wb.sheetnames

    assert "Comece Aqui" in sheet_names
    assert "Dashboard" in sheet_names
    assert "Lan\u00e7amentos" in sheet_names
    assert "Configura\u00e7\u00f5es" in sheet_names

    ws = wb["Lan\u00e7amentos"]

    # Check table existence (openpyxl supports this)
    assert "tblLancamentos" in ws.tables

    table = ws.tables["tblLancamentos"]
    assert table.ref == "B4:L5"

    # Check that protection was disabled for natural expansion
    assert ws.protection.sheet is False
    assert wb["Comece Aqui"].protection.sheet is True
    assert wb["Dashboard"].protection.sheet is True
    assert wb["Configura\u00e7\u00f5es"].protection.sheet is False

    # Verify defined names
    assert "lista_categorias" in wb.defined_names
    assert "lista_contas" in wb.defined_names
    assert "lista_cartoes" in wb.defined_names
    assert wb.defined_names["lista_categorias"].value == "tblCategorias[Categoria]"

    # Bug 1: Verify Dashboard formulas are written as formulas, not text
    dash_ws = wb["Dashboard"]

    # Saldo Disponível Hoje
    saldo_cell = dash_ws["B5"]
    assert str(saldo_cell.value).startswith("=")
    assert "SUMIFS(tblLancamentos[Valor],tblLancamentos[Tipo],\"Receita\")" in str(saldo_cell.value)

    # Receitas do Mês
    receita_cell = dash_ws["D5"]
    assert str(receita_cell.value).startswith("=")
    assert "SUMIFS(tblLancamentos[Valor]" in str(receita_cell.value)

    # Despesas do Mês
    despesa_cell = dash_ws["E5"]
    assert str(despesa_cell.value).startswith("=")
    assert "SUMIFS(tblLancamentos[Valor]" in str(despesa_cell.value)

    req_dark = GenerationRequest(template_id="finance_personal", year=2026, theme="dark")
    path_dark = generate(req_dark, output_dir)
    assert os.path.exists(path_dark)

def test_literal_string_starts_with_equal(tmp_path):
    from excel_saas.core.models.workbook_plan import WorkbookPlan, WorksheetPlan, CellPlan
    from excel_saas.core.engine.workbook_engine import WorkbookEngine
    from excel_saas.themes.light import LightTheme

    plan = WorkbookPlan(
        worksheets=[
            WorksheetPlan(
                name="Test",
                cells=[
                    CellPlan(row=0, col=0, value="=not_a_formula"),
                ]
            )
        ]
    )

    engine = WorkbookEngine(plan, LightTheme())
    path = os.path.join(str(tmp_path), "test_string.xlsx")
    engine.render(path)

    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb["Test"]

    cell = ws["A1"]
    assert cell.value == "=not_a_formula"
    assert cell.data_type == "s" # String type, not formula ('f')
