from typing import List
from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, DataValidationPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.core.excel.references import DefinedNameRef
from .defaults import DEFAULT_ACCOUNTS_SAMPLE

def build_contas(request: GenerationRequest) -> WorksheetPlan:
    sim_nao_validation = DataValidationPlan(validate="list", source=["Sim", "Não"])
    tipo_validation = DataValidationPlan(validate="list", source=DefinedNameRef("lista_tipos_conta"))
    
    columns = [
        ColumnPlan(header="Nome", width=25),
        ColumnPlan(header="Tipo", validation=tipo_validation, width=20),
        ColumnPlan(header="Instituição", width=20),
        ColumnPlan(header="Saldo inicial", number_format="R$ #,##0.00", width=15),
        ColumnPlan(header="Incluir no saldo disponível?", validation=sim_nao_validation, width=28),
        ColumnPlan(header="Ativa?", validation=sim_nao_validation, width=12),
    ]
    
    data = DEFAULT_ACCOUNTS_SAMPLE if request.with_sample_data else []
    
    table = TablePlan(
        name="tblContas",
        start_cell="B4",
        columns=columns,
        data=data,
        show_total_row=False
    )
    
    cells = [
        CellPlan(row=1, col=1, value="Contas", role=CellRole.TITLE, size=16, bold=True),
        CellPlan(row=2, col=1, value="Cadastre aqui onde seu dinheiro fica.", role=CellRole.NORMAL),
    ]
    
    return WorksheetPlan(
        name="Contas",
        is_protected=False,
        show_gridlines=False,
        cells=cells,
        tables=[table],
        column_widths={0: 3} # padding
    )
