from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, DataValidationPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.excel.formulas import Formula

from .investimentos_semantics import (
    CLASSES_PERMITIDAS,
    build_status,
    build_resultado_total,
    build_retorno_simples,
    build_peso_carteira,
    build_situacao,
    build_sys_aporte_valido,
    build_sys_recebido_valido,
    build_sys_valor_atual_valido
)

def build_investimentos_sheet() -> WorksheetPlan:
    classe_val = DataValidationPlan(
        validate="list",
        source=CLASSES_PERMITIDAS,
        ignore_blank=True,
        error_title="Classe inválida",
        error_message="Selecione uma classe válida na lista."
    )
    
    aportado_val = DataValidationPlan(
        validate="decimal",
        criteria=">",
        minimum=0,
        ignore_blank=True,
        error_title="Total aportado inválido",
        error_message="O total aportado deve ser maior que zero."
    )
    
    recebido_val = DataValidationPlan(
        validate="decimal",
        criteria=">=",
        minimum=0,
        ignore_blank=True,
        error_title="Total recebido inválido",
        error_message="O total recebido deve ser maior ou igual a zero."
    )

    atual_val = DataValidationPlan(
        validate="decimal",
        criteria=">=",
        minimum=0,
        ignore_blank=True,
        error_title="Valor atual inválido",
        error_message="O valor atual deve ser maior ou igual a zero."
    )

    columns = [
        ColumnPlan(header="Ativo", role=CellRole.INPUT),
        ColumnPlan(header="Classe", role=CellRole.INPUT, validation=classe_val),
        ColumnPlan(header="Instituição", role=CellRole.INPUT),
        ColumnPlan(header="Total aportado", role=CellRole.INPUT, validation=aportado_val, number_format="R$ #,##0.00"),
        ColumnPlan(header="Total recebido", role=CellRole.INPUT, validation=recebido_val, number_format="R$ #,##0.00"),
        ColumnPlan(header="Valor atual", role=CellRole.INPUT, validation=atual_val, number_format="R$ #,##0.00"),
        ColumnPlan(header="Resultado total", role=CellRole.FORMULA, formula=Formula(build_resultado_total()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Retorno simples %", role=CellRole.FORMULA, formula=Formula(build_retorno_simples()), number_format="0.0%"),
        ColumnPlan(header="Peso carteira %", role=CellRole.FORMULA, formula=Formula(build_peso_carteira()), number_format="0.0%"),
        ColumnPlan(header="Status", role=CellRole.FORMULA, formula=Formula(build_status())),
        ColumnPlan(header="Situação", role=CellRole.FORMULA, formula=Formula(build_situacao())),
        ColumnPlan(header="sys_AporteValido", role=CellRole.SYSTEM, formula=Formula(build_sys_aporte_valido()), hidden=True),
        ColumnPlan(header="sys_RecebidoValido", role=CellRole.SYSTEM, formula=Formula(build_sys_recebido_valido()), hidden=True),
        ColumnPlan(header="sys_ValorAtualValido", role=CellRole.SYSTEM, formula=Formula(build_sys_valor_atual_valido()), hidden=True)
    ]

    table = TablePlan(
        name="tblInvestimentos",
        start_cell="B4",
        columns=columns,
        data=[[None] * 14]
    )

    cells = [
        CellPlan(row=1, col=1, value="Investimentos", role=CellRole.TITLE),
        CellPlan(row=2, col=1, value="Cadastre cada posição. Informe o total aportado, o total já recebido ou resgatado e o valor atual para acompanhar retorno e composição da carteira.", role=CellRole.NORMAL)
    ]

    return WorksheetPlan(
        name="Investimentos",
        tables=[table],
        cells=cells,
        is_protected=False,
        show_gridlines=False,
        freeze_panes="B5"
    )
