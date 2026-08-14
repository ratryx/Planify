from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, DataValidationPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.excel.formulas import Formula

from .dividas_semantics import (
    CATEGORIAS_DIVIDAS,
    build_status,
    build_sys_saldo_devedor_valido,
    build_sys_parcela_mensal_valida
)

def build_dividas_sheet() -> WorksheetPlan:
    cat_validation = DataValidationPlan(
        validate="list",
        source='"' + ",".join(CATEGORIAS_DIVIDAS) + '"',
        ignore_blank=True,
        error_message="Selecione uma categoria válida."
    )
    
    saldo_validation = DataValidationPlan(
        validate="decimal",
        criteria=">",
        minimum=0,
        ignore_blank=True,
        error_message="Insira um saldo devedor válido (maior que 0)."
    )
    
    parcela_validation = DataValidationPlan(
        validate="decimal",
        criteria=">=",
        minimum=0,
        ignore_blank=True,
        error_message="Insira um valor de parcela válido (maior ou igual a 0)."
    )
    
    data_validation = DataValidationPlan(
        validate="date",
        criteria="between",
        minimum=1,
        maximum=2958465,
        ignore_blank=True,
        error_message="Insira uma data válida."
    )
    
    columns = [
        ColumnPlan(header="Dívida", role=CellRole.INPUT, width=25),
        ColumnPlan(header="Categoria", role=CellRole.INPUT, validation=cat_validation, width=20),
        ColumnPlan(header="Credor", role=CellRole.INPUT, width=20),
        ColumnPlan(header="Saldo devedor atual", role=CellRole.INPUT, validation=saldo_validation, number_format="R$ #,##0.00", width=20),
        ColumnPlan(header="Parcela mensal atual", role=CellRole.INPUT, validation=parcela_validation, number_format="R$ #,##0.00", width=20),
        ColumnPlan(header="Data final", role=CellRole.INPUT, validation=data_validation, number_format="dd/mm/yyyy", width=15),
        ColumnPlan(header="Observação", role=CellRole.INPUT, width=25),
        ColumnPlan(header="Status", role=CellRole.FORMULA, formula=Formula(build_status()), width=25),
        ColumnPlan(header="sys_SaldoDevedorValido", role=CellRole.SYSTEM, formula=Formula(build_sys_saldo_devedor_valido()), hidden=True, number_format="R$ #,##0.00"),
        ColumnPlan(header="sys_ParcelaMensalValida", role=CellRole.SYSTEM, formula=Formula(build_sys_parcela_mensal_valida()), hidden=True, number_format="R$ #,##0.00")
    ]
    
    table = TablePlan(
        name="tblDividas",
        start_cell="B4",
        columns=columns,
        data=[[None] * 10]
    )
    
    cells = [
        CellPlan(row=1, col=1, value="Dívidas", role=CellRole.TITLE),
        CellPlan(row=2, col=1, value="Cadastre empréstimos, financiamentos e outras dívidas contratuais. Não repita faturas, compras parceladas ou saldos negativos já registrados em outras abas.", role=CellRole.NORMAL)
    ]
    
    return WorksheetPlan(
        name="Dívidas",
        tables=[table],
        cells=cells,
        is_protected=False,
        show_gridlines=False,
        freeze_panes="B5"
    )
