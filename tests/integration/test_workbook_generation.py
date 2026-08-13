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

    assert sheet_names == [
        "Comece Aqui",
        "Dashboard",
        "Lançamentos",
        "Contas",
        "Cartões",
        "Configurações"
    ]

    ws = wb["Lançamentos"]

    # Check table existence (openpyxl supports this)
    assert "tblLancamentos" in ws.tables

    table = ws.tables["tblLancamentos"]
    assert table.ref == "B4:U5" # Columns B to U (20 columns now)

    # Check new tables
    contas_ws = wb["Contas"]
    assert "tblContas" in contas_ws.tables

    cartoes_ws = wb["Cartões"]
    assert "tblCartoes" in cartoes_ws.tables

    # Check exact columns for tblContas and tblCartoes
    contas_headers = [cell.value for cell in contas_ws[4] if cell.value]
    assert contas_headers == ["Nome", "Tipo", "Instituição", "Saldo inicial", "Saldo atual", "Incluir no saldo disponível?", "Ativa?", "Status"]

    cartoes_headers = [cell.value for cell in cartoes_ws[6] if cell.value]
    assert cartoes_headers == ["Nome", "Limite", "Dia fechamento", "Dia vencimento", "Conta de pagamento", "Ativo?"]

    config_ws = wb["Configurações"]
    assert "tblCategorias" in config_ws.tables
    assert "tblTiposConta" in config_ws.tables
    assert "tblContas" not in config_ws.tables
    assert "tblCartoes" not in config_ws.tables

    # Check that protection was disabled for natural expansion
    assert ws.protection.sheet is False
    assert contas_ws.protection.sheet is False
    assert cartoes_ws.protection.sheet is False
    assert config_ws.protection.sheet is False

    assert wb["Comece Aqui"].protection.sheet is True
    assert wb["Dashboard"].protection.sheet is True

    # Verify defined names
    assert "lista_categorias" in wb.defined_names
    assert "lista_contas" in wb.defined_names
    assert "lista_cartoes" in wb.defined_names
    assert "lista_tipos_conta" in wb.defined_names

    assert wb.defined_names["lista_categorias"].value == "tblCategorias[Categoria]"
    assert wb.defined_names["lista_tipos_conta"].value == "tblTiposConta[Tipo]"
    assert wb.defined_names["lista_contas"].value == "tblContas[Nome]"
    assert wb.defined_names["lista_cartoes"].value == "tblCartoes[Nome]"

    # Bug 1: Verify Dashboard formulas are written as formulas, not text
    dash_ws = wb["Dashboard"]

    # Saldo Disponível Hoje
    saldo_disponivel = dash_ws["B5"]
    assert str(saldo_disponivel.value).startswith("=")
    assert "SUMIFS(tblContas[Saldo atual]" in str(saldo_disponivel.value)

    # Resultado Registrado
    saldo_cell = dash_ws["B8"]
    assert str(saldo_cell.value).startswith("=")
    assert "SUM(tblLancamentos[sys_Receita])" in str(saldo_cell.value)

    # Receitas Registradas
    receita_cell = dash_ws["D8"]
    assert str(receita_cell.value).startswith("=")
    assert "SUM(tblLancamentos[sys_Receita])" in str(receita_cell.value)

    # Despesas Registradas
    despesa_cell = dash_ws["E8"]
    assert str(despesa_cell.value).startswith("=")
    assert "SUM(tblLancamentos[sys_Despesa])" in str(despesa_cell.value)

    # Check data validation logic for Contas (especially Saldo inicial on column E)
    saldo_val = None
    for val in contas_ws.data_validations.dataValidation:
        if "E" in val.sqref.__str__():
            saldo_val = val
            break

    assert saldo_val is not None
    assert saldo_val.type == "decimal"
    assert saldo_val.operator in (None, "between") # between is the default operator in Excel XML
    assert saldo_val.allowBlank is True

    # We must assert that it correctly handles negative, positive, and zero, and uses exact Excel technical bounds
    assert "-9.99999999999999e+307" in saldo_val.formula1.lower()
    assert "9.99999999999999e+307" in saldo_val.formula2.lower()
    assert "1000000000" not in saldo_val.formula1

    # Check Limite >= 0
    limite_val = None
    dia_fechamento_val = None
    dia_vencimento_val = None
    conta_pagamento_val = None
    for val in cartoes_ws.data_validations.dataValidation:
        sqref = val.sqref.__str__()
        if "C" in sqref:
            limite_val = val
        elif "D" in sqref:
            dia_fechamento_val = val
        elif "E" in sqref:
            dia_vencimento_val = val
        elif "F" in sqref:
            conta_pagamento_val = val

    assert limite_val is not None
    assert limite_val.type == "decimal"
    assert "0" in limite_val.formula1

    assert dia_fechamento_val is not None
    assert dia_fechamento_val.type == "whole"
    assert "1" in dia_fechamento_val.formula1
    assert "31" in dia_fechamento_val.formula2

    assert dia_vencimento_val is not None
    assert dia_vencimento_val.type == "whole"
    assert "1" in dia_vencimento_val.formula1
    assert "31" in dia_vencimento_val.formula2

    assert conta_pagamento_val is not None
    assert conta_pagamento_val.type == "list"
    assert "lista_contas" in conta_pagamento_val.formula1

    # Check Lançamentos validations
    categoria_val = None
    conta_val = None
    conta_dest_val = None
    cartao_val = None
    valor_val = None
    parcelas_val = None

    for val in ws.data_validations.dataValidation:
        sqref = val.sqref.__str__()
        if "E" in sqref:
            categoria_val = val
        elif "F" in sqref:
            conta_val = val
        elif "G" in sqref:
            conta_dest_val = val
        elif "H" in sqref:
            cartao_val = val
        elif "I" in sqref:
            valor_val = val
        elif "J" in sqref:
            parcelas_val = val

    assert categoria_val is not None and "lista_categorias" in categoria_val.formula1
    assert conta_val is not None and "lista_contas" in conta_val.formula1
    assert conta_dest_val is not None and "lista_contas" in conta_dest_val.formula1
    assert cartao_val is not None and "lista_cartoes" in cartao_val.formula1

    assert valor_val is not None
    assert valor_val.type == "decimal"
    assert valor_val.operator == "greaterThan"
    assert "0" in valor_val.formula1
    assert valor_val.allowBlank is True

    assert parcelas_val is not None
    assert parcelas_val.type == "whole"
    assert parcelas_val.operator == "greaterThanOrEqual"
    assert "1" in parcelas_val.formula1
    assert parcelas_val.allowBlank is True

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

def test_empty_mode_tables(tmp_path):
    output_dir = str(tmp_path)
    req = GenerationRequest(template_id="finance_personal", year=2026, with_sample_data=False)
    path = generate(req, output_dir)
    wb = openpyxl.load_workbook(path, data_only=False)

    # 1 row of headers, 1 row of blank data
    assert wb["Contas"].tables["tblContas"].ref == "B4:I5"
    assert wb["Cartões"].tables["tblCartoes"].ref == "B6:G7"

    # Assert data validations exist for the blank row in empty mode
    contas_ws = wb["Contas"]
    cartoes_ws = wb["Cartões"]
    # Verify validation is expanded
    assert len(contas_ws.data_validations.dataValidation) > 0
    assert len(cartoes_ws.data_validations.dataValidation) > 0

def test_sample_mode_tables(tmp_path):
    output_dir = str(tmp_path)
    req = GenerationRequest(template_id="finance_personal", year=2026, with_sample_data=True)
    path = generate(req, output_dir)
    wb = openpyxl.load_workbook(path, data_only=False)

    contas_ws = wb["Contas"]
    contas_ref = contas_ws.tables["tblContas"].ref
    # B4:I7 since default sample has 3 accounts (header + 3 data rows)
    assert contas_ref == "B4:I7"

    cartoes_ws = wb["Cartões"]
    cartoes_ref = cartoes_ws.tables["tblCartoes"].ref
    # B6:G8 since default sample has 2 cards (header + 2 data rows)
    assert cartoes_ref == "B6:G8"

    # Sample mode verification: check if Nubank is present
    assert contas_ws["B5"].value == "Nubank"
    assert cartoes_ws["B7"].value == "Nubank"
    # Conta de pagamento for the Nubank card is "Nubank"
    assert cartoes_ws["F7"].value == "Nubank"
