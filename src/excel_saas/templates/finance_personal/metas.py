from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, DataValidationPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.excel.formulas import Formula

from .metas_semantics import (
    build_falta,
    build_progresso,
    build_meses_restantes,
    build_aporte_mensal,
    build_status,
    build_situacao
)

def build_metas_sheet() -> WorksheetPlan:
    alvo_val = DataValidationPlan(
        validate="decimal",
        criteria=">",
        minimum=0,
        ignore_blank=True,
        error_title="Valor alvo inválido",
        error_message="O valor alvo deve ser maior que zero."
    )
    
    atual_val = DataValidationPlan(
        validate="decimal",
        criteria=">=",
        minimum=0,
        ignore_blank=True,
        error_title="Valor atual inválido",
        error_message="O valor atual deve ser maior ou igual a zero."
    )
    
    data_val = DataValidationPlan(
        validate="date",
        criteria="between",
        minimum=1,
        maximum=2958465,
        ignore_blank=True,
        error_title="Data inválida",
        error_message="A data alvo deve ser uma data válida."
    )

    columns = [
        ColumnPlan(header="Meta", role=CellRole.INPUT),
        ColumnPlan(header="Valor alvo", role=CellRole.INPUT, validation=alvo_val, number_format="R$ #,##0.00"),
        ColumnPlan(header="Valor atual", role=CellRole.INPUT, validation=atual_val, number_format="R$ #,##0.00"),
        ColumnPlan(header="Data alvo", role=CellRole.INPUT, validation=data_val, number_format="mmm/yyyy"),
        ColumnPlan(header="Falta", role=CellRole.FORMULA, formula=Formula(build_falta()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Progresso %", role=CellRole.FORMULA, formula=Formula(build_progresso()), number_format="0.0%"),
        ColumnPlan(header="Meses restantes", role=CellRole.FORMULA, formula=Formula(build_meses_restantes())),
        ColumnPlan(header="Aporte mensal necessário", role=CellRole.FORMULA, formula=Formula(build_aporte_mensal()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Status", role=CellRole.FORMULA, formula=Formula(build_status())),
        ColumnPlan(header="Situação", role=CellRole.FORMULA, formula=Formula(build_situacao()))
    ]

    table = TablePlan(
        name="tblMetas",
        start_cell="B4",
        columns=columns,
        data=[[]]  # Empty initial row
    )

    cells = [
        CellPlan(row=1, col=1, value="Metas", role=CellRole.TITLE),
        CellPlan(row=2, col=1, value="Defina seus objetivos financeiros e atualize o valor já reservado para acompanhar o progresso.", role=CellRole.NORMAL)
    ]

    return WorksheetPlan(
        name="Metas",
        tables=[table],
        cells=cells,
        is_protected=False,
        show_gridlines=False,
        freeze_panes="B5"
    )
