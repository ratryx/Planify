from typing import List
from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, DataValidationPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.models.generation_request import GenerationRequest

def build_comece_aqui(request: GenerationRequest) -> WorksheetPlan:
    cells = [
        CellPlan(row=1, col=1, value="Bem-vindo ao Financeiro Pessoal", role=CellRole.TITLE, size=16, bold=True),
        CellPlan(row=3, col=1, value=f"Ano Base: {request.year}", role=CellRole.NORMAL, bold=True),
        CellPlan(row=5, col=1, value="Instruções Iniciais:", role=CellRole.HEADER, bold=True),
        CellPlan(row=6, col=1, value="1. Configure suas categorias e contas na aba 'Configurações'."),
        CellPlan(row=7, col=1, value="2. Registre suas movimentações na aba 'Lançamentos'."),
        CellPlan(row=8, col=1, value="3. Acompanhe seus resultados na aba 'Dashboard'."),
        CellPlan(row=10, col=1, value="Convenções Visuais:"),
        CellPlan(row=11, col=1, value="Células editáveis (Input)", role=CellRole.INPUT),
        CellPlan(row=12, col=1, value="Células automáticas (Fórmulas/Sistema) - Não edite", role=CellRole.FORMULA),
    ]
    
    return WorksheetPlan(
        name="Comece Aqui",
        is_protected=True,
        show_gridlines=False,
        cells=cells,
        column_widths={1: 60}
    )

def build_dashboard(request: GenerationRequest) -> WorksheetPlan:
    from excel_saas.core.excel.formulas import sumifs, subtract, literal
    from excel_saas.core.excel.references import TableRef
    
    val_ref = TableRef("tblLancamentos", "Valor")
    tipo_ref = TableRef("tblLancamentos", "Tipo")
    
    cells = [
        CellPlan(row=1, col=1, value="Dashboard Principal", role=CellRole.TITLE, size=16, bold=True),
        CellPlan(row=3, col=1, value="Saldo Disponível Hoje", role=CellRole.HEADER),
        CellPlan(row=4, col=1, formula=subtract(sumifs(val_ref, tipo_ref, literal('Receita')), sumifs(val_ref, tipo_ref, literal('Despesa'))), role=CellRole.FORMULA, number_format="R$ #,##0.00", size=14, bold=True),
        CellPlan(row=3, col=3, value="Receitas do Mês", role=CellRole.HEADER),
        CellPlan(row=4, col=3, formula=sumifs(val_ref, tipo_ref, literal('Receita')), role=CellRole.FORMULA, number_format="R$ #,##0.00", bold=True),
        CellPlan(row=3, col=4, value="Despesas do Mês", role=CellRole.HEADER),
        CellPlan(row=4, col=4, formula=sumifs(val_ref, tipo_ref, literal('Despesa')), role=CellRole.FORMULA, number_format="R$ #,##0.00", bold=True),
    ]
    
    return WorksheetPlan(
        name="Dashboard",
        is_protected=True,
        show_gridlines=False,
        cells=cells,
        column_widths={1: 25, 2: 5, 3: 20, 4: 20}
    )

def build_lancamentos(request: GenerationRequest) -> WorksheetPlan:
    from excel_saas.core.excel.references import DefinedNameRef
    
    tipo_validation = DataValidationPlan(
        validate="list",
        source=["Receita", "Despesa", "Transferência", "Investimento", "Resgate", "Pagamento de dívida"]
    )
    sim_nao_validation = DataValidationPlan(validate="list", source=["Sim", "Não"])
    
    cat_validation = DataValidationPlan(validate="list", source=DefinedNameRef("lista_categorias"))
    conta_validation = DataValidationPlan(validate="list", source=DefinedNameRef("lista_contas"))
    cartao_validation = DataValidationPlan(validate="list", source=DefinedNameRef("lista_cartoes"))
    
    columns = [
        ColumnPlan(header="Data", number_format="dd/mm/yyyy", width=12),
        ColumnPlan(header="Descrição", width=30),
        ColumnPlan(header="Tipo", validation=tipo_validation, width=15),
        ColumnPlan(header="Categoria", validation=cat_validation, width=20),
        ColumnPlan(header="Conta", validation=conta_validation, width=15),
        ColumnPlan(header="Cartão", validation=cartao_validation, width=15),
        ColumnPlan(header="Valor", number_format="R$ #,##0.00", width=15),
        ColumnPlan(header="Parcelas", width=10),
        ColumnPlan(header="Essencial?", validation=sim_nao_validation, width=12),
        ColumnPlan(header="Recorrente?", validation=sim_nao_validation, width=12),
        ColumnPlan(header="Observação", width=25),
    ]
    
    table = TablePlan(
        name="tblLancamentos",
        start_cell="B4",
        columns=columns,
        show_total_row=False
    )
    
    cells = [
        CellPlan(row=1, col=1, value="Lançamentos", role=CellRole.TITLE, size=16, bold=True)
    ]
    
    return WorksheetPlan(
        name="Lançamentos",
        is_protected=False,
        freeze_panes="B5",
        cells=cells,
        tables=[table],
        column_widths={0: 3} # padding
    )

def build_configuracoes(request: GenerationRequest) -> WorksheetPlan:
    cells = [
        CellPlan(row=1, col=1, value="Configurações", role=CellRole.TITLE, size=16, bold=True),
        CellPlan(row=3, col=1, value="Configure abaixo suas listas. Você pode adicionar novas linhas às tabelas sempre que precisar.", role=CellRole.NORMAL),
    ]
    
    default_categories = [["Moradia"], ["Alimentação"], ["Transporte"], ["Saúde"], ["Educação"], ["Lazer"], ["Assinaturas"], ["Compras"], ["Viagens"], ["Impostos"], ["Investimentos"], ["Dívidas"], ["Outros"]]
    
    tbl_categorias = TablePlan(
        name="tblCategorias",
        start_cell="B6",
        columns=[ColumnPlan(header="Categoria")],
        data=default_categories
    )
    
    tbl_contas = TablePlan(
        name="tblContas",
        start_cell="D6",
        columns=[ColumnPlan(header="Conta")],
        data=[["Conta Corrente"], ["Poupança"], ["Carteira"]]
    )
    
    tbl_cartoes = TablePlan(
        name="tblCartoes",
        start_cell="F6",
        columns=[ColumnPlan(header="Cartão")],
        data=[["Cartão Principal"]]
    )
        
    return WorksheetPlan(
        name="Configurações",
        is_protected=False, # Intentionally unprotected for Table expansion
        show_gridlines=False,
        cells=cells,
        tables=[tbl_categorias, tbl_contas, tbl_cartoes],
        column_widths={1: 25, 2: 5, 3: 20, 4: 5, 5: 20}
    )
