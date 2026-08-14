from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.excel.formulas import Formula

from .dashboard_patrimonio_semantics import (
    build_account_assets,
    build_investment_assets,
    build_additional_assets,
    build_total_assets,
    build_component_value,
    build_component_weight
)

def build_dashboard_patrimonio_sheet() -> WorksheetPlan:
    columns = [
        ColumnPlan(header="Componente", role=CellRole.NORMAL),
        ColumnPlan(header="Valor atual", role=CellRole.FORMULA, formula=Formula(build_component_value()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Peso %", role=CellRole.FORMULA, formula=Formula(build_component_weight()), number_format="0.0%")
    ]
    
    table = TablePlan(
        name="tblResumoPatrimonio",
        start_cell="B11",
        columns=columns,
        data=[
            ["Contas e caixa"],
            ["Investimentos"],
            ["Bens patrimoniais"]
        ]
    )
    
    cells = [
        CellPlan(row=1, col=1, value="Dashboard Patrimônio", role=CellRole.TITLE),
        CellPlan(row=2, col=1, value="Visão consolidada dos ativos atuais. Dívidas e outros passivos ainda não são descontados.", role=CellRole.NORMAL),
        
        CellPlan(row=3, col=1, value="Contas e caixa", role=CellRole.HEADER),
        CellPlan(row=4, col=1, formula=Formula(build_account_assets()), role=CellRole.FORMULA, number_format="R$ #,##0.00"),
        
        CellPlan(row=3, col=3, value="Investimentos", role=CellRole.HEADER),
        CellPlan(row=4, col=3, formula=Formula(build_investment_assets()), role=CellRole.FORMULA, number_format="R$ #,##0.00"),
        
        CellPlan(row=3, col=5, value="Bens patrimoniais", role=CellRole.HEADER),
        CellPlan(row=4, col=5, formula=Formula(build_additional_assets()), role=CellRole.FORMULA, number_format="R$ #,##0.00"),
        
        CellPlan(row=6, col=1, value="Ativos totais", role=CellRole.HEADER),
        CellPlan(row=7, col=1, formula=Formula(build_total_assets()), role=CellRole.FORMULA, number_format="R$ #,##0.00"),
        
        CellPlan(row=8, col=1, value="Contas representam caixa; investimentos e bens devem ser cadastrados apenas em suas abas próprias para evitar dupla contagem.", role=CellRole.NORMAL)
    ]
    
    return WorksheetPlan(
        name="Dashboard Patrimônio",
        tables=[table],
        cells=cells,
        is_protected=True,
        show_gridlines=False
    )
