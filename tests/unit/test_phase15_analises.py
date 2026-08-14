from excel_saas.templates.finance_personal.analises_semantics import (
    build_total_assets,
    build_negative_account_balances,
    build_structural_debt_balance,
    build_registered_card_position,
    build_registered_position,
    build_available_balance,
    build_reserve_coverage,
    build_reserve_gap,
    build_horizon_commitments,
    build_horizon_budget,
    build_horizon_margin,
    build_horizon_usage,
    build_goal_monthly_contribution,
)
from excel_saas.templates.finance_personal.analises import build_analises_sheet
from excel_saas.core.excel.formulas import Expression

def test_gross_assets_semantics():
    from excel_saas.templates.finance_personal.dashboard_patrimonio_semantics import build_total_assets as base_assets
    assert str(build_total_assets()) == str(base_assets())

def test_negative_accounts_semantics():
    expr = build_negative_account_balances()
    assert str(expr) == "-SUMIFS(tblContas[Saldo atual],tblContas[Status],\"OK\",tblContas[Ativa?],\"Sim\",tblContas[Saldo atual],\"<0\")"

def test_structural_debt_semantics():
    expr = build_structural_debt_balance()
    assert str(expr) == "SUM(tblDividas[sys_SaldoDevedorValido])"

def test_card_position_semantics():
    expr = build_registered_card_position()
    assert str(expr) == "SUM(tblLancamentos[sys_Cartao])"

def test_registered_position_semantics():
    expr = build_registered_position()
    assert str(expr) == "SUM(SUMIFS(tblContas[Saldo atual],tblContas[Status],\"OK\",tblContas[Ativa?],\"Sim\",tblContas[Saldo atual],\">0\"),SUM(tblInvestimentos[sys_ValorAtualValido]),SUM(tblBensPatrimoniais[sys_ValorAtualValido]))--SUMIFS(tblContas[Saldo atual],tblContas[Status],\"OK\",tblContas[Ativa?],\"Sim\",tblContas[Saldo atual],\"<0\")-SUM(tblDividas[sys_SaldoDevedorValido])-SUM(tblLancamentos[sys_Cartao])"

def test_available_balance_semantics():
    expr = build_available_balance()
    assert str(expr) == "SUMIFS(tblContas[Saldo atual],tblContas[Incluir no saldo disponível?],\"Sim\",tblContas[Ativa?],\"Sim\",tblContas[Status],\"OK\")"

def test_reserve_metrics_semantics():
    expr_cov = build_reserve_coverage()
    expr_gap = build_reserve_gap()
    assert str(expr_cov) == "SUM(tblReserva[Cobertura atual])"
    assert str(expr_gap) == "SUM(tblReserva[Falta])"

def test_goal_metrics_semantics():
    expr = build_goal_monthly_contribution()
    assert str(expr) == "SUM(tblMetas[Aporte mensal necessário])"

def test_horizon_metrics_semantics():
    expr_com = build_horizon_commitments()
    expr_bud = build_horizon_budget()
    expr_mar = build_horizon_margin()
    expr_use = build_horizon_usage()
    
    assert str(expr_com) == "SUM(tblProjecoes[Compromissos conhecidos])"
    assert str(expr_bud) == "SUM(tblProjecoes[Orçamento planejado])"
    assert str(expr_mar) == "SUM(tblProjecoes[Orçamento planejado])-SUM(tblProjecoes[Compromissos conhecidos])"
    assert str(expr_use) == "IF(AND(SUM(tblProjecoes[Orçamento planejado])=0,SUM(tblProjecoes[Compromissos conhecidos])=0),0,IF(SUM(tblProjecoes[Orçamento planejado])=0,\"\",SUM(tblProjecoes[Compromissos conhecidos])/SUM(tblProjecoes[Orçamento planejado])))"

def test_domain_isolation():
    sheet = build_analises_sheet()
    
    # Assert missing domains in ANY cell
    for cell in sheet.cells:
        if cell.role.name == "FORMULA" and isinstance(cell.formula, Expression):
            val = str(cell.formula)
            assert "tblCartoes" not in val
            assert "tblFaturas" not in val
            assert "tblParcelamentos" not in val
            assert "tblOrcamento[" not in val
            assert "tblResumoInvestimentos" not in val
            assert "tblResumoPatrimonio" not in val
            
    # Assert missing domains in position formula
    pos_val = str(build_registered_position())
    assert "tblReserva" not in pos_val
    assert "tblMetas" not in pos_val
    assert "tblProjecoes" not in pos_val

def test_analises_sheet_model():
    sheet = build_analises_sheet()
    
    assert sheet.name == "Análises"
    assert sheet.is_protected is True
    assert sheet.show_gridlines is False
    assert len(sheet.tables) == 0
    
    # Ensure no input cells (everything is either string/header or formula)
    for cell in sheet.cells:
        assert getattr(cell, "role", None) != "input"
