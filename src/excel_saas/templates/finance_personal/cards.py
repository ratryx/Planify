from typing import List
from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, DataValidationPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.core.excel.references import DefinedNameRef
from .defaults import DEFAULT_CARDS_SAMPLE

def build_cartoes(request: GenerationRequest) -> WorksheetPlan:
    sim_nao_validation = DataValidationPlan(validate="list", source=["Sim", "Não"])
    conta_validation = DataValidationPlan(validate="list", source=DefinedNameRef("lista_contas"), ignore_blank=True)
    
    limite_validation = DataValidationPlan(
        validate="decimal",
        criteria=">=",
        minimum=0,
        ignore_blank=True,
        error_message="O limite deve ser maior ou igual a zero."
    )
    
    dia_validation = DataValidationPlan(
        validate="integer",
        criteria="between",
        minimum=1,
        maximum=31,
        ignore_blank=True,
        error_message="O dia deve estar entre 1 e 31."
    )
    
    columns = [
        ColumnPlan(header="Nome", width=25),
        ColumnPlan(header="Limite", number_format="R$ #,##0.00", validation=limite_validation, width=15),
        ColumnPlan(header="Dia fechamento", validation=dia_validation, width=15),
        ColumnPlan(header="Dia vencimento", validation=dia_validation, width=15),
        ColumnPlan(header="Conta de pagamento", validation=conta_validation, width=25),
        ColumnPlan(header="Ativo?", validation=sim_nao_validation, width=12),
    ]
    
    data = DEFAULT_CARDS_SAMPLE if request.with_sample_data else []
    
    table = TablePlan(
        name="tblCartoes",
        start_cell="B6",
        columns=columns,
        data=data,
        show_total_row=False
    )
    
    cells = [
        CellPlan(row=1, col=1, value="Cartões", role=CellRole.TITLE, size=16, bold=True),
        CellPlan(row=2, col=1, value="Para cadastrar um cartão, basta informar um nome.", role=CellRole.NORMAL),
        CellPlan(row=3, col=1, value="Limite, fechamento, vencimento e conta de pagamento são opcionais.", role=CellRole.NORMAL),
    ]
    
    return WorksheetPlan(
        name="Cartões",
        is_protected=False,
        show_gridlines=False,
        cells=cells,
        tables=[table],
        column_widths={0: 3} # padding
    )
