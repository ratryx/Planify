from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, DataValidationPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.excel.formulas import Formula

from .patrimonio_semantics import (
    CATEGORIAS_PATRIMONIO,
    build_status,
    build_sys_valor_atual_valido
)

def build_patrimonio_sheet() -> WorksheetPlan:
    categoria_val = DataValidationPlan(
        validate="list",
        source=CATEGORIAS_PATRIMONIO,
        ignore_blank=True,
        error_title="Categoria inválida",
        error_message="Selecione uma categoria válida na lista."
    )
    
    valor_atual_val = DataValidationPlan(
        validate="decimal",
        criteria=">=",
        minimum=0,
        ignore_blank=True,
        error_title="Valor atual inválido",
        error_message="O valor atual deve ser maior ou igual a zero."
    )
    
    columns = [
        ColumnPlan(header="Bem", role=CellRole.INPUT),
        ColumnPlan(header="Categoria", role=CellRole.INPUT, validation=categoria_val),
        ColumnPlan(header="Valor atual", role=CellRole.INPUT, validation=valor_atual_val, number_format="R$ #,##0.00"),
        ColumnPlan(header="Observação", role=CellRole.INPUT),
        ColumnPlan(header="Status", role=CellRole.FORMULA, formula=Formula(build_status())),
        ColumnPlan(header="sys_ValorAtualValido", role=CellRole.SYSTEM, formula=Formula(build_sys_valor_atual_valido()), hidden=True)
    ]
    
    table = TablePlan(
        name="tblBensPatrimoniais",
        start_cell="B4",
        columns=columns,
        data=[[None] * 6]
    )
    
    cells = [
        CellPlan(row=1, col=1, value="Bens Patrimoniais", role=CellRole.TITLE),
        CellPlan(row=2, col=1, value="Cadastre apenas bens que não estejam representados nas abas Contas ou Investimentos, como imóveis, veículos e participações em empresas. Informe o valor atual bruto estimado.", role=CellRole.NORMAL)
    ]
    
    return WorksheetPlan(
        name="Patrimônio",
        tables=[table],
        cells=cells,
        is_protected=False,
        show_gridlines=False,
        freeze_panes="B5"
    )
