import pytest
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.templates.finance_personal.template import FinancePersonalTemplate
from tests.fixtures.realistic_household_state import inject_realistic_household
from excel_saas.templates.finance_personal.defaults import DEFAULT_ACCOUNT_TYPES
from excel_saas.templates.finance_personal.investimentos import CLASSES_PERMITIDAS
from excel_saas.templates.finance_personal.dividas import CATEGORIAS_DIVIDAS

def get_table(plan, name: str):
    for ws in plan.worksheets:
        for tbl in ws.tables:
            if tbl.name == name:
                return tbl
    raise ValueError(f"Table {name} not found")

def get_col_index(tbl, header_name: str) -> int:
    for i, col in enumerate(tbl.columns):
        if col.header == header_name:
            return i
    raise ValueError(f"Column {header_name} not found in {tbl.name}")

def test_realistic_fixture_contracts():
    req = GenerationRequest(template_id="finance_personal", year=2026, theme="light")
    template = FinancePersonalTemplate()
    plan = template.build_workbook_plan(req)
    
    inject_realistic_household(plan, 2026)
    
    # Check tblLancamentos
    tbl_lancamentos = get_table(plan, "tblLancamentos")
    idx_valor = get_col_index(tbl_lancamentos, "Valor")
    idx_tipo = get_col_index(tbl_lancamentos, "Tipo")
    
    assert len(tbl_lancamentos.data) > 0, "tblLancamentos must not be empty"
    
    tipos_presentes = set()
    for row in tbl_lancamentos.data:
        valor = row[idx_valor]
        if valor is not None:
            assert isinstance(valor, (int, float)), "Valor must be numeric"
            assert valor > 0, "Valor must be strictly positive"
            
        tipo = row[idx_tipo]
        if tipo is not None:
            tipos_presentes.add(tipo)
            
    assert "Estorno / Reembolso" in tipos_presentes, "Refund must use 'Estorno / Reembolso'"
    
    # Check tblContas
    tbl_contas = get_table(plan, "tblContas")
    idx_tipo_conta = get_col_index(tbl_contas, "Tipo")
    assert len(tbl_contas.data) > 0, "tblContas must not be empty"
    valid_account_types = [t[0] if isinstance(t, list) else t for t in DEFAULT_ACCOUNT_TYPES]
    for row in tbl_contas.data:
        t = row[idx_tipo_conta]
        if t is not None:
            assert t in valid_account_types, f"Invalid account type: {t}"
            
    # Check tblInvestimentos
    tbl_invest = get_table(plan, "tblInvestimentos")
    idx_classe = get_col_index(tbl_invest, "Classe")
    assert len(tbl_invest.data) > 0, "tblInvestimentos must not be empty"
    for row in tbl_invest.data:
        c = row[idx_classe]
        if c is not None:
            assert c in CLASSES_PERMITIDAS, f"Invalid investment class: {c}"
            
    # Check tblDividas
    tbl_dividas = get_table(plan, "tblDividas")
    idx_cat = get_col_index(tbl_dividas, "Categoria")
    assert len(tbl_dividas.data) > 0, "tblDividas must not be empty"
    for row in tbl_dividas.data:
        cat = row[idx_cat]
        if cat is not None:
            assert cat in CATEGORIAS_DIVIDAS, f"Invalid debt category: {cat}"
