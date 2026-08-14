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
from excel_saas.core.models.cell_roles import CellRole

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
    
    assert sheet.column_widths == {
        1: 28,
        2: 4,
        3: 28,
        4: 4,
        5: 28,
        6: 4,
        7: 32,
    }
    
    # Ensure no input cells
    for cell in sheet.cells:
        assert cell.role != CellRole.INPUT

    # Exact coordinate contract (row/col are zero-based, B2 -> 1,1)
    cells_by_coord = {(c.row, c.col): c for c in sheet.cells}
    
    # Titles
    assert cells_by_coord[(1, 1)].value == "Análises Financeiras"
    assert cells_by_coord[(2, 1)].value == "Resumo executivo dos dados registrados no Planify. A posição patrimonial considera apenas ativos, dívidas e movimentos de cartão cadastrados; saldos anteriores de cartão não registrados não são incluídos."
    
    # Section A
    assert cells_by_coord[(4, 1)].value == "Posição atual"
    assert cells_by_coord[(5, 1)].value == "Ativos totais"
    assert cells_by_coord[(5, 3)].value == "Saldos negativos em contas"
    assert cells_by_coord[(5, 5)].value == "Dívidas estruturais"
    assert cells_by_coord[(5, 7)].value == "Saldo líquido registrado em cartões"
    assert cells_by_coord[(8, 1)].value == "Posição patrimonial registrada"
    
    # Section B
    assert cells_by_coord[(11, 1)].value == "Liquidez e segurança"
    assert cells_by_coord[(12, 1)].value == "Saldo disponível hoje"
    assert cells_by_coord[(12, 3)].value == "Cobertura da reserva"
    assert cells_by_coord[(12, 5)].value == "Falta para reserva"
    
    # Section C
    assert cells_by_coord[(16, 1)].value == "Planejamento"
    assert cells_by_coord[(17, 1)].value == "Compromissos no horizonte"
    assert cells_by_coord[(17, 3)].value == "Orçamento no horizonte"
    assert cells_by_coord[(17, 5)].value == "Margem no horizonte"
    assert cells_by_coord[(17, 7)].value == "Uso conhecido no horizonte"
    assert cells_by_coord[(20, 1)].value == "Aporte mensal para metas"
    
    # Formats
    brl_coords = [(6, 1), (6, 3), (6, 5), (6, 7), (9, 1), (13, 1), (13, 5), (18, 1), (18, 3), (18, 5), (21, 1)]
    for coord in brl_coords:
        assert cells_by_coord[coord].number_format == "R$ #,##0.00"
        
    assert cells_by_coord[(13, 3)].number_format == "0.0"
    assert cells_by_coord[(18, 7)].number_format == "0.0%"
    
    assert cells_by_coord[(9, 1)].bold is True
    assert cells_by_coord[(9, 1)].size == 14
