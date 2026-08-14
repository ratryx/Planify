import pytest
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.templates.finance_personal.dividas_semantics import (
    CATEGORIAS_DIVIDAS,
    _build_is_valid_debt,
    _build_is_valid_date,
    build_status,
    build_sys_saldo_devedor_valido,
    build_sys_parcela_mensal_valida
)
from excel_saas.templates.finance_personal.dividas import build_dividas_sheet

def test_category_model():
    assert CATEGORIAS_DIVIDAS == [
        "Financiamento imobiliário",
        "Financiamento de veículo",
        "Empréstimo pessoal",
        "Consignado",
        "Dívida tributária",
        "Dívida com pessoa",
        "Outros",
    ]
    
    assert "Cartão de crédito" not in CATEGORIAS_DIVIDAS

def test_safe_date_predicate():
    from excel_saas.core.excel.references import ThisRowRef
    f = str(Formula(_build_is_valid_date(ThisRowRef("Data final"))))
    assert f == '=AND(ISNUMBER([@[Data final]]),[@[Data final]]>=1,[@[Data final]]<DATE(9999,12,31)+1)'
    assert 'TODAY' not in f

def test_status_precedence():
    f = str(Formula(build_status()))
    assert f == '=IF(AND(ISBLANK([@Dívida]),ISBLANK([@Categoria]),ISBLANK([@Credor]),ISBLANK([@[Saldo devedor atual]]),ISBLANK([@[Parcela mensal atual]]),ISBLANK([@[Data final]]),ISBLANK([@Observação])),"",IF(ISBLANK([@Dívida]),"Informe a dívida",IF(ISBLANK([@Categoria]),"Informe a categoria",IF(NOT(OR([@Categoria]="Financiamento imobiliário",[@Categoria]="Financiamento de veículo",[@Categoria]="Empréstimo pessoal",[@Categoria]="Consignado",[@Categoria]="Dívida tributária",[@Categoria]="Dívida com pessoa",[@Categoria]="Outros")),"Categoria inválida",IF(COUNTIFS(tblDividas[Dívida],[@Dívida])>1,"Dívida duplicada",IF(ISBLANK([@[Saldo devedor atual]]),"Informe o saldo devedor",IF(OR(NOT(ISNUMBER([@[Saldo devedor atual]])),[@[Saldo devedor atual]]<=0),"Saldo devedor inválido",IF(AND(NOT(ISBLANK([@[Parcela mensal atual]])),OR(NOT(ISNUMBER([@[Parcela mensal atual]])),[@[Parcela mensal atual]]<0)),"Parcela mensal inválida",IF(AND(NOT(ISBLANK([@[Data final]])),NOT(AND(ISNUMBER([@[Data final]]),[@[Data final]]>=1,[@[Data final]]<DATE(9999,12,31)+1))),"Data final inválida","OK")))))))))'
    assert '[@Status]' not in f

def test_is_valid_debt():
    f = str(Formula(_build_is_valid_debt()))
    assert f == '=AND(NOT(ISBLANK([@Dívida])),OR([@Categoria]="Financiamento imobiliário",[@Categoria]="Financiamento de veículo",[@Categoria]="Empréstimo pessoal",[@Categoria]="Consignado",[@Categoria]="Dívida tributária",[@Categoria]="Dívida com pessoa",[@Categoria]="Outros"),COUNTIFS(tblDividas[Dívida],[@Dívida])=1,AND(NOT(ISBLANK([@[Saldo devedor atual]])),ISNUMBER([@[Saldo devedor atual]]),[@[Saldo devedor atual]]>0),OR(ISBLANK([@[Parcela mensal atual]]),AND(ISNUMBER([@[Parcela mensal atual]]),[@[Parcela mensal atual]]>=0)),OR(ISBLANK([@[Data final]]),AND(ISNUMBER([@[Data final]]),[@[Data final]]>=1,[@[Data final]]<DATE(9999,12,31)+1)))'
    assert '[@Status]' not in f

def test_valid_balance_helper():
    f = str(Formula(build_sys_saldo_devedor_valido()))
    assert f == '=IF(AND(NOT(ISBLANK([@Dívida])),OR([@Categoria]="Financiamento imobiliário",[@Categoria]="Financiamento de veículo",[@Categoria]="Empréstimo pessoal",[@Categoria]="Consignado",[@Categoria]="Dívida tributária",[@Categoria]="Dívida com pessoa",[@Categoria]="Outros"),COUNTIFS(tblDividas[Dívida],[@Dívida])=1,AND(NOT(ISBLANK([@[Saldo devedor atual]])),ISNUMBER([@[Saldo devedor atual]]),[@[Saldo devedor atual]]>0),OR(ISBLANK([@[Parcela mensal atual]]),AND(ISNUMBER([@[Parcela mensal atual]]),[@[Parcela mensal atual]]>=0)),OR(ISBLANK([@[Data final]]),AND(ISNUMBER([@[Data final]]),[@[Data final]]>=1,[@[Data final]]<DATE(9999,12,31)+1))),[@[Saldo devedor atual]],0)'
    assert '[@Status]' not in f

def test_valid_installment_helper():
    f = str(Formula(build_sys_parcela_mensal_valida()))
    assert f == '=IF(AND(NOT(ISBLANK([@Dívida])),OR([@Categoria]="Financiamento imobiliário",[@Categoria]="Financiamento de veículo",[@Categoria]="Empréstimo pessoal",[@Categoria]="Consignado",[@Categoria]="Dívida tributária",[@Categoria]="Dívida com pessoa",[@Categoria]="Outros"),COUNTIFS(tblDividas[Dívida],[@Dívida])=1,AND(NOT(ISBLANK([@[Saldo devedor atual]])),ISNUMBER([@[Saldo devedor atual]]),[@[Saldo devedor atual]]>0),OR(ISBLANK([@[Parcela mensal atual]]),AND(ISNUMBER([@[Parcela mensal atual]]),[@[Parcela mensal atual]]>=0)),OR(ISBLANK([@[Data final]]),AND(ISNUMBER([@[Data final]]),[@[Data final]]>=1,[@[Data final]]<DATE(9999,12,31)+1))),IF(ISBLANK([@[Parcela mensal atual]]),0,[@[Parcela mensal atual]]),0)'
    assert '[@Status]' not in f

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

    cat_col = table.columns[1]
    assert cat_col.validation is not None
    assert cat_col.validation.validate == "list"
    assert cat_col.validation.ignore_blank is True

    saldo_col = table.columns[3]
    assert saldo_col.validation is not None
    assert saldo_col.validation.validate == "decimal"
    assert saldo_col.validation.criteria == ">"
    assert saldo_col.validation.minimum == 0
    assert saldo_col.validation.ignore_blank is True

    parcela_col = table.columns[4]
    assert parcela_col.validation is not None
    assert parcela_col.validation.validate == "decimal"
    assert parcela_col.validation.criteria == ">="
    assert parcela_col.validation.minimum == 0
    assert parcela_col.validation.ignore_blank is True

    data_col = table.columns[5]
    assert data_col.validation is not None
    assert data_col.validation.validate == "date"
    assert data_col.validation.criteria == "between"
    assert data_col.validation.minimum == 1
    assert data_col.validation.maximum == 2958465
    assert data_col.validation.ignore_blank is True

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
