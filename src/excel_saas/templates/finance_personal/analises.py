from excel_saas.core.models.workbook_plan import WorksheetPlan, CellPlan, CellRole
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

BRL_FORMAT = "R$ #,##0.00"

def build_analises_sheet() -> WorksheetPlan:
    cells = []

    # Title & Instruction
    cells.append(CellPlan(row=1, col=1, value="Análises Financeiras", role=CellRole.TITLE))
    cells.append(CellPlan(
        row=2, 
        col=1, 
        value="Resumo executivo dos dados registrados no Planify. A posição patrimonial considera apenas ativos, dívidas e movimentos de cartão cadastrados; saldos anteriores de cartão não registrados não são incluídos.", 
        role=CellRole.NORMAL
    ))

    # SECTION A — Posição atual
    cells.append(CellPlan(row=4, col=1, value="Posição atual", role=CellRole.HEADER))
    
    cells.append(CellPlan(row=5, col=1, value="Ativos totais", role=CellRole.HEADER))
    cells.append(CellPlan(row=6, col=1, formula=build_total_assets(), role=CellRole.FORMULA, number_format=BRL_FORMAT))
    
    cells.append(CellPlan(row=5, col=3, value="Saldos negativos em contas", role=CellRole.HEADER))
    cells.append(CellPlan(row=6, col=3, formula=build_negative_account_balances(), role=CellRole.FORMULA, number_format=BRL_FORMAT))
    
    cells.append(CellPlan(row=5, col=5, value="Dívidas estruturais", role=CellRole.HEADER))
    cells.append(CellPlan(row=6, col=5, formula=build_structural_debt_balance(), role=CellRole.FORMULA, number_format=BRL_FORMAT))
    
    cells.append(CellPlan(row=5, col=7, value="Saldo líquido registrado em cartões", role=CellRole.HEADER))
    cells.append(CellPlan(row=6, col=7, formula=build_registered_card_position(), role=CellRole.FORMULA, number_format=BRL_FORMAT))
    
    cells.append(CellPlan(row=8, col=1, value="Posição patrimonial registrada", role=CellRole.HEADER))
    cells.append(CellPlan(row=9, col=1, formula=build_registered_position(), role=CellRole.FORMULA, number_format=BRL_FORMAT, bold=True, size=14))

    # SECTION B — Liquidez e segurança
    cells.append(CellPlan(row=11, col=1, value="Liquidez e segurança", role=CellRole.HEADER))
    
    cells.append(CellPlan(row=12, col=1, value="Saldo disponível hoje", role=CellRole.HEADER))
    cells.append(CellPlan(row=13, col=1, formula=build_available_balance(), role=CellRole.FORMULA, number_format=BRL_FORMAT))
    
    cells.append(CellPlan(row=12, col=3, value="Cobertura da reserva", role=CellRole.HEADER))
    cells.append(CellPlan(row=13, col=3, formula=build_reserve_coverage(), role=CellRole.FORMULA, number_format="0.0"))
    
    cells.append(CellPlan(row=12, col=5, value="Falta para reserva", role=CellRole.HEADER))
    cells.append(CellPlan(row=13, col=5, formula=build_reserve_gap(), role=CellRole.FORMULA, number_format=BRL_FORMAT))

    # SECTION C — Planejamento
    cells.append(CellPlan(row=16, col=1, value="Planejamento", role=CellRole.HEADER))
    
    cells.append(CellPlan(row=17, col=1, value="Compromissos no horizonte", role=CellRole.HEADER))
    cells.append(CellPlan(row=18, col=1, formula=build_horizon_commitments(), role=CellRole.FORMULA, number_format=BRL_FORMAT))
    
    cells.append(CellPlan(row=17, col=3, value="Orçamento no horizonte", role=CellRole.HEADER))
    cells.append(CellPlan(row=18, col=3, formula=build_horizon_budget(), role=CellRole.FORMULA, number_format=BRL_FORMAT))
    
    cells.append(CellPlan(row=17, col=5, value="Margem no horizonte", role=CellRole.HEADER))
    cells.append(CellPlan(row=18, col=5, formula=build_horizon_margin(), role=CellRole.FORMULA, number_format=BRL_FORMAT))
    
    cells.append(CellPlan(row=17, col=7, value="Uso conhecido no horizonte", role=CellRole.HEADER))
    cells.append(CellPlan(row=18, col=7, formula=build_horizon_usage(), role=CellRole.FORMULA, number_format="0.0%"))
    
    cells.append(CellPlan(row=20, col=1, value="Aporte mensal para metas", role=CellRole.HEADER))
    cells.append(CellPlan(row=21, col=1, formula=build_goal_monthly_contribution(), role=CellRole.FORMULA, number_format=BRL_FORMAT))


    return WorksheetPlan(
        name="Análises",
        is_protected=True,
        show_gridlines=False,
        tables=[],
        cells=cells,
        column_widths={2: 28, 3: 4, 4: 28, 5: 4, 6: 28, 7: 4, 8: 32}
    )
