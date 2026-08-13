from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.templates.finance_personal.parcelamentos_semantics import (
    build_sys_indice_lancamento, build_projected_column,
    build_parcelas_restantes, build_proxima_competencia,
    build_compromisso_restante, build_situacao
)

def build_parcelamentos_sheet() -> WorksheetPlan:
    # 100 rows, 1 to 100, placed in the 12th column (sys_Ordem)
    # The columns with formulas can receive None or empty string. None is typical.
    # Col index 11 is sys_Ordem
    ordem_data = []
    for i in range(1, 101):
        row = [None] * 14
        row[11] = i
        ordem_data.append(row)

    table = TablePlan(
        name="tblParcelamentos",
        start_cell="B4",
        columns=[
            ColumnPlan(header="Descrição", formula=build_projected_column("Descrição"), role=CellRole.FORMULA, width=35),
            ColumnPlan(header="Cartão", formula=build_projected_column("Cartão"), role=CellRole.FORMULA, width=20),
            ColumnPlan(header="Valor original", formula=build_projected_column("Valor"), role=CellRole.FORMULA, width=15, number_format="R$ #,##0.00"),
            ColumnPlan(header="Parcelas", formula=build_projected_column("sys_ParcelasEfetivas"), role=CellRole.FORMULA, width=10),
            ColumnPlan(header="Valor da parcela", formula=build_projected_column("sys_ValorParcelaBase"), role=CellRole.FORMULA, width=18, number_format="R$ #,##0.00"),
            ColumnPlan(header="Primeira fatura", formula=build_projected_column("sys_FaturaInicial"), role=CellRole.FORMULA, width=15, number_format="mmm/yyyy"),
            ColumnPlan(header="Última fatura", formula=build_projected_column("sys_FaturaFinal"), role=CellRole.FORMULA, width=15, number_format="mmm/yyyy"),
            ColumnPlan(header="Parcelas restantes", formula=build_parcelas_restantes(), role=CellRole.FORMULA, width=18),
            ColumnPlan(header="Próxima competência", formula=build_proxima_competencia(), role=CellRole.FORMULA, width=20, number_format="mmm/yyyy"),
            ColumnPlan(header="Compromisso restante", formula=build_compromisso_restante(), role=CellRole.FORMULA, width=22, number_format="R$ #,##0.00"),
            ColumnPlan(header="Situação", formula=build_situacao(), role=CellRole.FORMULA, width=15),
            
            ColumnPlan(header="sys_Ordem", role=CellRole.SYSTEM, hidden=True),
            ColumnPlan(header="sys_IndiceLancamento", formula=build_sys_indice_lancamento(), role=CellRole.SYSTEM, hidden=True),
            ColumnPlan(header="sys_AjusteUltimaParcela", formula=build_projected_column("sys_AjusteUltimaParcela"), role=CellRole.SYSTEM, hidden=True)
        ],
        data=ordem_data,
        show_total_row=False
    )

    cells = [
        CellPlan(row=1, col=1, value="Parcelamentos", role=CellRole.TITLE, size=16),
        CellPlan(row=2, col=1, value="Visão automática das compras parceladas. Não edite esta tabela. Exibe até 100 compras.", role=CellRole.NORMAL)
    ]

    return WorksheetPlan(
        name="Parcelamentos",
        is_protected=True,
        freeze_panes="B5",
        show_gridlines=False,
        cells=cells,
        tables=[table]
    )
