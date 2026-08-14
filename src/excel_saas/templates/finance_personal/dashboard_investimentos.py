from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.excel.formulas import Formula
from .dashboard_investimentos_semantics import (
    build_portfolio_aportado, build_portfolio_recebido, build_portfolio_atual,
    build_portfolio_resultado, build_portfolio_retorno,
    build_class_aportado, build_class_recebido, build_class_atual,
    build_class_resultado, build_class_peso, build_class_retorno
)
from .investimentos_semantics import CLASSES_PERMITIDAS

def build_dashboard_investimentos_sheet() -> WorksheetPlan:
    cells = [
        CellPlan(row=1, col=1, value="Dashboard de Investimentos", role=CellRole.TITLE),
        CellPlan(row=2, col=1, value="Visão consolidada da carteira cadastrada na aba Investimentos.", role=CellRole.NORMAL),
        
        # KPIs
        CellPlan(row=3, col=1, value="Total aportado", role=CellRole.HEADER),
        CellPlan(row=4, col=1, formula=Formula(build_portfolio_aportado()), role=CellRole.FORMULA, number_format="R$ #,##0.00"),
        
        CellPlan(row=3, col=3, value="Total recebido", role=CellRole.HEADER),
        CellPlan(row=4, col=3, formula=Formula(build_portfolio_recebido()), role=CellRole.FORMULA, number_format="R$ #,##0.00"),
        
        CellPlan(row=3, col=5, value="Valor atual", role=CellRole.HEADER),
        CellPlan(row=4, col=5, formula=Formula(build_portfolio_atual()), role=CellRole.FORMULA, number_format="R$ #,##0.00"),
        
        CellPlan(row=6, col=1, value="Resultado total", role=CellRole.HEADER),
        CellPlan(row=7, col=1, formula=Formula(build_portfolio_resultado()), role=CellRole.FORMULA, number_format="R$ #,##0.00"),
        
        CellPlan(row=6, col=3, value="Retorno simples %", role=CellRole.HEADER),
        CellPlan(row=7, col=3, formula=Formula(build_portfolio_retorno()), role=CellRole.FORMULA, number_format="0.0%"),
    ]
    
    # Table cols
    columns = [
        ColumnPlan(header="Classe", role=CellRole.NORMAL),
        ColumnPlan(header="Total aportado", role=CellRole.FORMULA, formula=Formula(build_class_aportado()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Total recebido", role=CellRole.FORMULA, formula=Formula(build_class_recebido()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Valor atual", role=CellRole.FORMULA, formula=Formula(build_class_atual()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Resultado total", role=CellRole.FORMULA, formula=Formula(build_class_resultado()), number_format="R$ #,##0.00"),
        ColumnPlan(header="Peso carteira %", role=CellRole.FORMULA, formula=Formula(build_class_peso()), number_format="0.0%"),
        ColumnPlan(header="Retorno simples %", role=CellRole.FORMULA, formula=Formula(build_class_retorno()), number_format="0.0%")
    ]
    
    table_data = []
    for cls in CLASSES_PERMITIDAS:
        table_data.append([cls] + [None] * 6)
        
    table = TablePlan(
        name="tblResumoInvestimentos",
        start_cell="B11",
        columns=columns,
        data=table_data
    )
    
    return WorksheetPlan(
        name="Dashboard Investimentos",
        cells=cells,
        tables=[table],
        is_protected=True,
        show_gridlines=False
    )
