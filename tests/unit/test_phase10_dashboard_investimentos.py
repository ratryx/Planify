import pytest
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.templates.finance_personal.dashboard_investimentos_semantics import (
    build_portfolio_aportado, build_portfolio_recebido, build_portfolio_atual,
    build_portfolio_resultado, build_portfolio_retorno,
    build_class_aportado, build_class_recebido, build_class_atual,
    build_class_resultado, build_class_peso, build_class_retorno
)
from excel_saas.templates.finance_personal.dashboard_investimentos import build_dashboard_investimentos_sheet
from excel_saas.templates.finance_personal.investimentos_semantics import CLASSES_PERMITIDAS

def test_portfolio_totals():
    f1 = str(Formula(build_portfolio_aportado()))
    f2 = str(Formula(build_portfolio_recebido()))
    f3 = str(Formula(build_portfolio_atual()))
    
    assert 'SUM(tblInvestimentos[sys_AporteValido])' in f1
    assert 'SUM(tblInvestimentos[sys_RecebidoValido])' in f2
    assert 'SUM(tblInvestimentos[sys_ValorAtualValido])' in f3

def test_portfolio_resultado():
    f = str(Formula(build_portfolio_resultado()))
    assert '(SUM(tblInvestimentos[sys_ValorAtualValido])+SUM(tblInvestimentos[sys_RecebidoValido]))-SUM(tblInvestimentos[sys_AporteValido])' in f

def test_portfolio_retorno():
    f = str(Formula(build_portfolio_retorno()))
    assert 'IF(SUM(tblInvestimentos[sys_AporteValido])=0,0,' in f
    assert '/SUM(tblInvestimentos[sys_AporteValido])' in f

def test_class_aggregation():
    f1 = str(Formula(build_class_aportado()))
    f2 = str(Formula(build_class_recebido()))
    f3 = str(Formula(build_class_atual()))
    
    assert 'SUMIFS(tblInvestimentos[sys_AporteValido],tblInvestimentos[Classe],[@Classe])' in f1
    assert 'SUMIFS(tblInvestimentos[sys_RecebidoValido],tblInvestimentos[Classe],[@Classe])' in f2
    assert 'SUMIFS(tblInvestimentos[sys_ValorAtualValido],tblInvestimentos[Classe],[@Classe])' in f3

def test_class_resultado():
    f = str(Formula(build_class_resultado()))
    assert '([@[Valor atual]]+[@[Total recebido]])-[@[Total aportado]]' in f

def test_class_peso():
    f = str(Formula(build_class_peso()))
    assert 'IF(SUM(tblInvestimentos[sys_ValorAtualValido])=0,0,[@[Valor atual]]/SUM(tblInvestimentos[sys_ValorAtualValido]))' in f
    assert 'tblInvestimentos[Valor atual]' not in f

def test_class_retorno():
    f = str(Formula(build_class_retorno()))
    assert 'IF([@[Total aportado]]=0,0,[@[Resultado total]]/[@[Total aportado]])' in f

def test_domain_isolation():
    exprs = [
        build_portfolio_aportado(), build_portfolio_recebido(), build_portfolio_atual(),
        build_portfolio_resultado(), build_portfolio_retorno(),
        build_class_aportado(), build_class_recebido(), build_class_atual(),
        build_class_resultado(), build_class_peso(), build_class_retorno()
    ]
    for expr in exprs:
        f = str(Formula(expr))
        assert 'tblLancamentos' not in f
        assert 'tblContas' not in f
        assert 'tblCartoes' not in f
        assert 'tblFaturas' not in f
        assert 'tblParcelamentos' not in f
        assert 'tblOrcamento' not in f
        assert 'tblMetas' not in f
        assert 'tblReserva' not in f

def test_table_plan():
    sheet = build_dashboard_investimentos_sheet()
    
    assert sheet.is_protected is True
    
    assert len(sheet.tables) == 1
    table = sheet.tables[0]
    
    assert len(table.columns) == 7
    
    headers = [col.header for col in table.columns]
    assert headers == [
        "Classe", "Total aportado", "Total recebido", "Valor atual", "Resultado total", "Peso carteira %", "Retorno simples %"
    ]
    
    assert table.columns[0].role == CellRole.NORMAL
    
    for i in range(1, 7):
        assert table.columns[i].role == CellRole.FORMULA
        
    assert len(table.data) == 8
    
    for i, row in enumerate(table.data):
        assert row[0] == CLASSES_PERMITIDAS[i]
