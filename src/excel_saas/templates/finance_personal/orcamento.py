from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, DataValidationPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.excel.references import StringRef, DefinedNameRef

from .orcamento_semantics import (
    build_sys_competencia_normalizada,
    build_status,
    build_saidas_em_conta,
    build_cartao_parcelas,
    build_consumido_comprometido,
    build_disponivel,
    build_uso_pct,
    build_situacao
)

def build_orcamento_sheet() -> WorksheetPlan:
    comp_validation = DataValidationPlan(
        validate="date",
        criteria="between",
        minimum=1,
        maximum=2958465,
        ignore_blank=True,
        error_title="Data inválida",
        error_message="A competência deve ser uma data válida."
    )
    
    cat_validation = DataValidationPlan(
        validate="list",
        source=DefinedNameRef("lista_categorias"),
        ignore_blank=True,
        error_title="Categoria inválida",
        error_message="Selecione uma categoria da lista."
    )
    
    orc_validation = DataValidationPlan(
        validate="decimal",
        criteria=">=",
        minimum=0,
        ignore_blank=True,
        error_title="Orçamento inválido",
        error_message="O orçamento deve ser maior ou igual a zero."
    )

    columns = [
        ColumnPlan(header="Competência", role=CellRole.INPUT, validation=comp_validation, number_format="mmm/yyyy"),
        ColumnPlan(header="Categoria", role=CellRole.INPUT, validation=cat_validation),
        ColumnPlan(header="Orçamento", role=CellRole.INPUT, validation=orc_validation, number_format="R$ #,##0.00"),
        ColumnPlan(header="Saídas em conta", role=CellRole.FORMULA, formula=Formula(build_saidas_em_conta()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Cartão / Parcelas", role=CellRole.FORMULA, formula=Formula(build_cartao_parcelas()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Consumido / Comprometido", role=CellRole.FORMULA, formula=Formula(build_consumido_comprometido()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Disponível", role=CellRole.FORMULA, formula=Formula(build_disponivel()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Uso %", role=CellRole.FORMULA, formula=Formula(build_uso_pct()), number_format="0.0%"),
        ColumnPlan(header="Status", role=CellRole.FORMULA, formula=Formula(build_status())),
        ColumnPlan(header="Situação", role=CellRole.FORMULA, formula=Formula(build_situacao())),
        ColumnPlan(header="sys_CompetenciaNormalizada", role=CellRole.SYSTEM, formula=Formula(build_sys_competencia_normalizada()), hidden=True),
    ]

    table = TablePlan(
        name="tblOrcamento",
        start_cell="B4",
        columns=columns,
        data=[[]]  # Empty initial row
    )

    cells = [
        CellPlan(row=1, col=1, value="Orçamento", role=CellRole.HEADER),
        CellPlan(row=2, col=1, value="Defina um orçamento por categoria e competência. Gastos e parcelas são calculados automaticamente.", role=CellRole.NORMAL)
    ]

    return WorksheetPlan(
        name="Orçamento",
        tables=[table],
        cells=cells,
        is_protected=False,
        show_gridlines=False,
        freeze_panes="B5"
    )
