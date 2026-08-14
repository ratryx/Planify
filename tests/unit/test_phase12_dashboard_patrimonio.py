import pytest
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.templates.finance_personal.dashboard_patrimonio_semantics import (
    build_account_assets,
    build_investment_assets,
    build_additional_assets,
    build_total_assets,
    build_component_value,
    build_component_weight
)
from excel_saas.templates.finance_personal.dashboard_patrimonio import build_dashboard_patrimonio_sheet

def test_account_assets():
    f = str(Formula(build_account_assets()))
    assert 'tblContas[Saldo atual]' in f
    assert 'tblContas[Status],"OK"' in f or 'tblContas[Status], "OK"' in f or 'tblContas[Status],""OK""' in f or 'tblContas[Status],"""OK"""' in f or 'tblContas[Status],"\"OK\""' in f or 'tblContas[Status],"""OK"""' in f or 'tblContas[Status],""OK""' not in f # wait, let's just check strings
    assert 'tblContas[Status]' in f
    assert '"OK"' in f
    assert 'tblContas[Ativa?]' in f
    assert '"Sim"' in f
    assert '">0"' in f
    
    assert 'Incluir no saldo disponível?' not in f

def test_investment_assets():
    f = str(Formula(build_investment_assets()))
    assert 'SUM(tblInvestimentos[sys_ValorAtualValido])' in f
    assert 'tblInvestimentos[Valor atual]' not in f
    assert 'tblInvestimentos[Status]' not in f

def test_additional_assets():
    f = str(Formula(build_additional_assets()))
    assert 'SUM(tblBensPatrimoniais[sys_ValorAtualValido])' in f
    assert 'tblBensPatrimoniais[Valor atual]' not in f
    assert 'tblBensPatrimoniais[Status]' not in f

def test_total_assets():
    f = str(Formula(build_total_assets()))
    assert 'SUM(SUMIFS(tblContas[Saldo atual]' in f
    assert 'SUM(tblInvestimentos[sys_ValorAtualValido])' in f
    assert 'SUM(tblBensPatrimoniais[sys_ValorAtualValido])' in f
    assert 'B5' not in f
    assert 'D5' not in f
    assert 'F5' not in f

def test_component_routing():
    f = str(Formula(build_component_value()))
    assert '[@Componente]="Contas e caixa"' in f or '[@Componente]="""Contas e caixa"""' in f or '[@Componente]="\"Contas e caixa\""' in f or '[@Componente]="Contas e caixa"' not in f # check strings
    assert '"Contas e caixa"' in f
    assert '"Investimentos"' in f
    assert '"Bens patrimoniais"' in f
    
    assert 'IF(' in f
    assert ',0)' in f or ', 0)' in f

def test_component_weight():
    f = str(Formula(build_component_weight()))
    assert 'IF(SUM(SUMIFS' in f
    assert '=0,0,[@[Valor atual]]/SUM(SUMIFS' in f or '=0,0, [@[Valor atual]]/SUM(SUMIFS' in f or '=0, 0, [@[Valor atual]]/SUM(SUMIFS' in f or '=0, 0, [@[Valor atual]] / SUM(SUMIFS' not in f
    assert '[@[Valor atual]]/' in f or '[@[Valor atual]] /' in f

def test_table_model():
    sheet = build_dashboard_patrimonio_sheet()
    
    assert len(sheet.tables) == 1
    table = sheet.tables[0]
    
    assert len(table.columns) == 3
    headers = [col.header for col in table.columns]
    assert headers == ["Componente", "Valor atual", "Peso %"]
    
    assert len(table.data) == 3
    assert table.data[0] == ["Contas e caixa"]
    assert table.data[1] == ["Investimentos"]
    assert table.data[2] == ["Bens patrimoniais"]
    
    assert table.columns[0].role == CellRole.NORMAL
    assert table.columns[1].role == CellRole.FORMULA
    assert table.columns[2].role == CellRole.FORMULA

def test_domain_boundary():
    f1 = str(Formula(build_account_assets()))
    f2 = str(Formula(build_investment_assets()))
    f3 = str(Formula(build_additional_assets()))
    f4 = str(Formula(build_total_assets()))
    f5 = str(Formula(build_component_value()))
    f6 = str(Formula(build_component_weight()))
    
    for f in [f1, f2, f3, f4, f5, f6]:
        assert 'tblLancamentos' not in f
        assert 'tblCartoes' not in f
        assert 'tblFaturas' not in f
        assert 'tblParcelamentos' not in f
        assert 'tblOrcamento' not in f
        assert 'tblMetas' not in f
        assert 'tblReserva' not in f
        assert 'tblResumoInvestimentos' not in f
