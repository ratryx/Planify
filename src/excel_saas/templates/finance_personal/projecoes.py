from typing import List
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.excel.formulas import Formula
from .projecoes_semantics import (
    _effective_projection_horizon,
    build_competencia,
    build_orcamento_planejado,
    build_card_commitments,
    build_structural_debt_commitments,
    build_known_commitments,
    build_budget_margin,
    build_budget_usage
)

def build_projecoes_sheet(request: GenerationRequest) -> WorksheetPlan:
    horizon = _effective_projection_horizon(request.projection_horizon)
    
    rows_data: List[List[str]] = []
    for i in range(horizon):
        rows_data.append([
            "", # Competência
            "", # Orçamento planejado
            "", # Compromissos no cartão
            "", # Dívidas estruturais
            "", # Compromissos conhecidos
            "", # Margem vs orçamento
            "", # Uso conhecido %
            str(i) # sys_Offset
        ])
    
    table = TablePlan(
        name="tblProjecoes",
        start_cell="B4",
        columns=[
            ColumnPlan(header="Competência", role=CellRole.FORMULA, formula=Formula(build_competencia()), number_format="mmm/yyyy"),
            ColumnPlan(header="Orçamento planejado", role=CellRole.FORMULA, formula=Formula(build_orcamento_planejado()), number_format="R$ #,##0.00"),
            ColumnPlan(header="Compromissos no cartão", role=CellRole.FORMULA, formula=Formula(build_card_commitments()), number_format="R$ #,##0.00"),
            ColumnPlan(header="Dívidas estruturais", role=CellRole.FORMULA, formula=Formula(build_structural_debt_commitments()), number_format="R$ #,##0.00"),
            ColumnPlan(header="Compromissos conhecidos", role=CellRole.FORMULA, formula=Formula(build_known_commitments()), number_format="R$ #,##0.00"),
            ColumnPlan(header="Margem vs orçamento", role=CellRole.FORMULA, formula=Formula(build_budget_margin()), number_format="R$ #,##0.00"),
            ColumnPlan(header="Uso conhecido %", role=CellRole.FORMULA, formula=Formula(build_budget_usage()), number_format="0.0%"),
            ColumnPlan(header="sys_Offset", role=CellRole.SYSTEM, hidden=True)
        ],
        data=rows_data
    )
    
    cells = [
        CellPlan(row=1, col=1, value="Projeções", role=CellRole.TITLE),
        CellPlan(row=2, col=1, value="Compromissos conhecidos a partir do próximo mês, comparados ao orçamento cadastrado. Não representa previsão completa de receitas, gastos ou saldo futuro.", role=CellRole.NORMAL)
    ]
    
    return WorksheetPlan(
        name="Projeções",
        tables=[table],
        cells=cells,
        is_protected=True,
        show_gridlines=False,
        freeze_panes="B5"
    )
