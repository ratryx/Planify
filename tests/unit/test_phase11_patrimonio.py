import pytest
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.templates.finance_personal.patrimonio_semantics import (
    CATEGORIAS_PATRIMONIO,
    build_sys_valor_atual_valido,
    build_status
)
from excel_saas.templates.finance_personal.patrimonio import build_patrimonio_sheet

def test_category_model():
    assert CATEGORIAS_PATRIMONIO == [
        "Imóveis",
        "Veículos",
        "Empresas / participações",
        "Bens de valor",
        "Direitos / créditos",
        "Outros"
    ]
    
    assert "Ações" not in CATEGORIAS_PATRIMONIO
    assert "Renda fixa" not in CATEGORIAS_PATRIMONIO

def test_duplicate_identity():
    f = str(Formula(build_status()))
    assert 'COUNTIFS(tblBensPatrimoniais[Bem],[@Bem])>1' in f
    assert '"Bem duplicado"' in f

def test_validity_predicate_and_helper():
    f = str(Formula(build_sys_valor_atual_valido()))
    
    # Must check required fields
    assert 'ISBLANK([@Bem])' in f
    assert 'ISBLANK([@Categoria])' in f
    assert 'ISBLANK([@[Valor atual]])' in f
    
    # Must check numeric and >= 0
    assert 'ISNUMBER([@[Valor atual]])' in f
    assert '[@[Valor atual]]>=0' in f
    
    # Must NOT reference Status
    assert '[@Status]' not in f
    
    # Must evaluate to Valor atual or 0
    assert 'IF(' in f
    assert '[@[Valor atual]]' in f
    assert ',0)' in f or ',0,0)' in f or ',0, 0)' in f

def test_status_strings():
    f = str(Formula(build_status()))
    
    assert '"Informe o bem"' in f
    assert '"Informe a categoria"' in f
    assert '"Categoria inválida"' in f
    assert '"Bem duplicado"' in f
    assert '"Informe o valor atual"' in f
    assert '"Valor atual inválido"' in f
    assert '"OK"' in f
    assert '""' in f # for empty row

def test_domain_isolation():
    f1 = str(Formula(build_sys_valor_atual_valido()))
    f2 = str(Formula(build_status()))
    
    for f in [f1, f2]:
        assert 'tblContas' not in f
        assert 'tblInvestimentos' not in f
        assert 'tblLancamentos' not in f
        assert 'tblCartoes' not in f
        assert 'tblFaturas' not in f
        assert 'tblParcelamentos' not in f
        assert 'tblOrcamento' not in f
        assert 'tblMetas' not in f
        assert 'tblReserva' not in f
        assert 'tblResumoInvestimentos' not in f

def test_table_plan():
    sheet = build_patrimonio_sheet()
    
    assert sheet.is_protected is False
    assert sheet.freeze_panes == "B5"
    assert sheet.show_gridlines is False
    
    assert len(sheet.tables) == 1
    table = sheet.tables[0]
    
    assert len(table.columns) == 6
    
    headers = [col.header for col in table.columns]
    assert headers == ["Bem", "Categoria", "Valor atual", "Observação", "Status", "sys_ValorAtualValido"]
    
    for i in range(4):
        assert table.columns[i].role == CellRole.INPUT
        
    assert table.columns[4].role == CellRole.FORMULA
    assert table.columns[5].role == CellRole.SYSTEM
    assert table.columns[5].hidden is True
