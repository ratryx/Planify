from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, DataValidationPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.models.generation_request import GenerationRequest

from .reserva_semantics import (
    build_reserva_alvo,
    build_falta,
    build_cobertura_atual,
    build_progresso,
    build_status,
    build_situacao
)

def build_reserva_sheet(request: GenerationRequest) -> WorksheetPlan:
    # Seed validation
    try:
        if type(request.reserve_months) is int and request.reserve_months >= 1:
            seed = request.reserve_months
        else:
            seed = 6
    except Exception:
        seed = 6

    custo_val = DataValidationPlan(
        validate="decimal",
        criteria=">",
        minimum=0,
        ignore_blank=True,
        error_title="Custo essencial inválido",
        error_message="O custo essencial deve ser maior que zero."
    )
    
    meses_val = DataValidationPlan(
        validate="whole",
        criteria=">=",
        minimum=1,
        ignore_blank=True,
        error_title="Meses inválidos",
        error_message="Os meses desejados devem ser no mínimo 1."
    )
    
    reserva_val = DataValidationPlan(
        validate="decimal",
        criteria=">=",
        minimum=0,
        ignore_blank=True,
        error_title="Reserva atual inválida",
        error_message="A reserva atual deve ser maior ou igual a zero."
    )

    columns = [
        ColumnPlan(header="Custo essencial mensal", role=CellRole.INPUT, validation=custo_val, number_format="R$ #,##0.00"),
        ColumnPlan(header="Meses desejados", role=CellRole.INPUT, validation=meses_val, number_format="General"),
        ColumnPlan(header="Reserva atual", role=CellRole.INPUT, validation=reserva_val, number_format="R$ #,##0.00"),
        ColumnPlan(header="Reserva alvo", role=CellRole.FORMULA, formula=Formula(build_reserva_alvo()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Falta", role=CellRole.FORMULA, formula=Formula(build_falta()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Cobertura atual", role=CellRole.FORMULA, formula=Formula(build_cobertura_atual()), number_format="0.0"),
        ColumnPlan(header="Progresso %", role=CellRole.FORMULA, formula=Formula(build_progresso()), number_format="0.0%"),
        ColumnPlan(header="Status", role=CellRole.FORMULA, formula=Formula(build_status())),
        ColumnPlan(header="Situação", role=CellRole.FORMULA, formula=Formula(build_situacao()))
    ]

    table = TablePlan(
        name="tblReserva",
        start_cell="B4",
        columns=columns,
        data=[[None, seed, None]]
    )

    cells = [
        CellPlan(row=1, col=1, value="Reserva de Emergência", role=CellRole.TITLE),
        CellPlan(row=2, col=1, value="Informe seu custo essencial mensal e o valor já reservado para acompanhar sua cobertura de emergência. Use apenas uma linha.", role=CellRole.NORMAL)
    ]

    return WorksheetPlan(
        name="Reserva",
        tables=[table],
        cells=cells,
        is_protected=False,
        show_gridlines=False,
        freeze_panes="B5"
    )
