import pytest
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.templates.finance_personal.dividas_semantics import (
    CATEGORIAS_DIVIDAS,
    _build_is_valid_debt,
    build_status,
    build_sys_saldo_devedor_valido,
    build_sys_parcela_mensal_valida
)
from excel_saas.templates.finance_personal.dividas import build_dividas_sheet

def test_category_model():
    assert len(CATEGORIAS_DIVIDAS) == 7
    assert "Financiamento imobiliário" in CATEGORIAS_DIVIDAS
    assert "Financiamento de veículo" in CATEGORIAS_DIVIDAS
    assert "Empréstimo pessoal" in CATEGORIAS_DIVIDAS
    assert "Consignado" in CATEGORIAS_DIVIDAS
    assert "Dívida tributária" in CATEGORIAS_DIVIDAS
    assert "Dívida com pessoa" in CATEGORIAS_DIVIDAS
    assert "Outros" in CATEGORIAS_DIVIDAS
    
    assert "Cartão de crédito" not in CATEGORIAS_DIVIDAS

def test_status_precedence():
    f = str(Formula(build_status()))
    assert 'COUNTIFS(tblDividas[Dívida],[@Dívida])>1' in f
    assert '"Dívida duplicada"' in f
    
    assert '"Informe a dívida"' in f
    assert '"Informe a categoria"' in f
    assert '"Categoria inválida"' in f
    assert '"Informe o saldo devedor"' in f
    assert '"Saldo devedor inválido"' in f
    assert '"Parcela mensal inválida"' in f
    assert '"Data final inválida"' in f
    assert '"OK"' in f
    assert '[@Status]' not in f

def test_is_valid_debt():
    f = str(Formula(_build_is_valid_debt()))
    assert 'NOT(ISBLANK([@Dívida]))' in f
    assert '[@Categoria]="Financiamento imobiliário"' in f
    assert 'COUNTIFS(tblDividas[Dívida],[@Dívida])=1' in f
    assert 'AND(NOT(ISBLANK([@[Saldo devedor atual]])),ISNUMBER([@[Saldo devedor atual]]),[@[Saldo devedor atual]]>0)' in f
    assert 'OR(ISBLANK([@[Parcela mensal atual]]),AND(ISNUMBER([@[Parcela mensal atual]]),[@[Parcela mensal atual]]>=0))' in f
    assert 'OR(ISBLANK([@[Data final]]),AND(ISNUMBER([@[Data final]]),[@[Data final]]>=1,[@[Data final]]<DATE(9999,12,31)+1))' in f
    assert '[@Status]' not in f

def test_valid_balance_helper():
    f = str(Formula(build_sys_saldo_devedor_valido()))
    assert 'IF(AND(NOT(ISBLANK([@Dívida]))' in f
    assert '[@[Saldo devedor atual]]' in f
    assert ',0)' in f

def test_valid_installment_helper():
    f = str(Formula(build_sys_parcela_mensal_valida()))
    assert 'IF(AND(NOT(ISBLANK([@Dívida]))' in f
    assert 'IF(ISBLANK([@[Parcela mensal atual]]),0,[@[Parcela mensal atual]])' in f
    assert ',0)' in f

def test_table_plan():
    sheet = build_dividas_sheet()
    
    assert sheet.is_protected is False
    assert sheet.show_gridlines is False
    assert sheet.freeze_panes == "B5"
    
    table = sheet.tables[0]
    assert len(table.columns) == 10
    
    headers = [c.header for c in table.columns]
    assert headers == [
        "Dívida", "Categoria", "Credor", "Saldo devedor atual",
        "Parcela mensal atual", "Data final", "Observação",
        "Status", "sys_SaldoDevedorValido", "sys_ParcelaMensalValida"
    ]
    
    for i in range(7):
        assert table.columns[i].role == CellRole.INPUT
        
    assert table.columns[7].role == CellRole.FORMULA
    assert table.columns[8].role == CellRole.SYSTEM
    assert table.columns[9].role == CellRole.SYSTEM
    
    assert table.columns[8].hidden is True
    assert table.columns[9].hidden is True

def test_domain_isolation():
    f1 = str(Formula(_build_is_valid_debt()))
    f2 = str(Formula(build_status()))
    f3 = str(Formula(build_sys_saldo_devedor_valido()))
    f4 = str(Formula(build_sys_parcela_mensal_valida()))
    
    for f in [f1, f2, f3, f4]:
        assert 'tblContas' not in f
        assert 'tblCartoes' not in f
        assert 'tblFaturas' not in f
        assert 'tblParcelamentos' not in f
        assert 'tblLancamentos' not in f
        assert 'tblOrcamento' not in f
        assert 'tblMetas' not in f
        assert 'tblReserva' not in f
        assert 'tblInvestimentos' not in f
        assert 'tblResumoInvestimentos' not in f
        assert 'tblBensPatrimoniais' not in f
        assert 'tblResumoPatrimonio' not in f
