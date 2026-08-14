from typing import List
from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, DataValidationPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.models.generation_request import GenerationRequest

def build_comece_aqui(request: GenerationRequest) -> WorksheetPlan:
    cells = [
        CellPlan(row=1, col=1, value="Bem-vindo ao Financeiro Pessoal", role=CellRole.TITLE),
        CellPlan(row=3, col=1, value=f"Ano Base: {request.year}", role=CellRole.NORMAL),
        CellPlan(row=5, col=1, value="Como usar o Planify:", role=CellRole.HEADER),
        CellPlan(row=6, col=1, value="1. Revise as categorias na aba 'Configurações' e ajuste-as se necessário.", role=CellRole.NORMAL),
        CellPlan(row=7, col=1, value="2. Cadastre suas contas na aba 'Contas' e seus cartões na aba 'Cartões'.", role=CellRole.NORMAL),
        CellPlan(row=8, col=1, value="Para cartões, apenas o nome é obrigatório. Limite, fechamento, vencimento e conta de pagamento são opcionais.", role=CellRole.NORMAL),
        CellPlan(row=9, col=1, value="3. Registre receitas, despesas, transferências, investimentos, resgates, pagamentos e estornos na aba 'Lançamentos'.", role=CellRole.NORMAL),
        CellPlan(row=10, col=1, value="4. Planeje seu orçamento, metas e reserva de emergência nas abas 'Orçamento', 'Metas' e 'Reserva'.", role=CellRole.NORMAL),
        CellPlan(row=11, col=1, value="5. Cadastre investimentos, bens patrimoniais e dívidas estruturais em suas respectivas abas.", role=CellRole.NORMAL),
        CellPlan(row=12, col=1, value="6. Consulte faturas, parcelamentos, dashboards, projeções e a aba 'Análises' para acompanhar sua situação financeira.", role=CellRole.NORMAL),
        CellPlan(row=14, col=1, value="Importante:", role=CellRole.HEADER),
        CellPlan(row=15, col=1, value="As projeções mostram compromissos conhecidos a partir do próximo mês e não representam uma previsão completa de receitas, gastos ou saldo futuro.", role=CellRole.NORMAL),
        CellPlan(row=16, col=1, value="A posição patrimonial considera somente os dados registrados no Planify; saldos anteriores de cartão não cadastrados não são incluídos.", role=CellRole.NORMAL),
        CellPlan(row=18, col=1, value="Convenções visuais:", role=CellRole.HEADER),
        CellPlan(row=19, col=1, value="Células editáveis (Input)", role=CellRole.INPUT),
        CellPlan(row=20, col=1, value="Células automáticas (Fórmulas/Sistema) - Não edite", role=CellRole.FORMULA),
    ]

    return WorksheetPlan(
        name="Comece Aqui",
        is_protected=True,
        show_gridlines=False,
        tables=[],
        cells=cells,
        column_widths={1: 110}
    )

def build_dashboard(request: GenerationRequest) -> WorksheetPlan:
    from excel_saas.core.excel.formulas import subtract, sum_func, sumifs, literal
    from excel_saas.core.excel.references import TableRef

    sys_receita_ref = TableRef("tblLancamentos", "sys_Receita")
    sys_despesa_ref = TableRef("tblLancamentos", "sys_Despesa")

    saldo_disponivel_formula = sumifs(
        TableRef("tblContas", "Saldo atual"),
        TableRef("tblContas", "Incluir no saldo disponível?"), literal("Sim"),
        TableRef("tblContas", "Ativa?"), literal("Sim"),
        TableRef("tblContas", "Status"), literal("OK")
    )

    cells = [
        CellPlan(row=1, col=1, value="Dashboard Principal", role=CellRole.TITLE, size=16, bold=True),
        CellPlan(row=3, col=1, value="Saldo Disponível Hoje", role=CellRole.HEADER),
        CellPlan(row=4, col=1, formula=saldo_disponivel_formula, role=CellRole.FORMULA, number_format="R$ #,##0.00", size=14, bold=True),

        CellPlan(row=6, col=1, value="Resultado Registrado", role=CellRole.HEADER),
        CellPlan(row=7, col=1, formula=subtract(sum_func(sys_receita_ref), sum_func(sys_despesa_ref)), role=CellRole.FORMULA, number_format="R$ #,##0.00", size=14, bold=True),
        CellPlan(row=6, col=3, value="Receitas Registradas", role=CellRole.HEADER),
        CellPlan(row=7, col=3, formula=sum_func(sys_receita_ref), role=CellRole.FORMULA, number_format="R$ #,##0.00", bold=True),
        CellPlan(row=6, col=4, value="Despesas Registradas", role=CellRole.HEADER),
        CellPlan(row=7, col=4, formula=sum_func(sys_despesa_ref), role=CellRole.FORMULA, number_format="R$ #,##0.00", bold=True),
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
        source=["Receita", "Despesa", "Transferência", "Investimento", "Resgate", "Pagamento de fatura", "Pagamento de dívida", "Estorno / Reembolso"]
    )
    sim_nao_validation = DataValidationPlan(validate="list", source=["Sim", "Não"])

    cat_validation = DataValidationPlan(validate="list", source=DefinedNameRef("lista_categorias"))
    conta_validation = DataValidationPlan(validate="list", source=DefinedNameRef("lista_contas"))
    cartao_validation = DataValidationPlan(validate="list", source=DefinedNameRef("lista_cartoes"))

    valor_validation = DataValidationPlan(
        validate="decimal",
        criteria=">",
        minimum=0,
        ignore_blank=True,
        error_message="O valor deve ser maior que zero."
    )

    parcelas_validation = DataValidationPlan(
        validate="integer",
        criteria=">=",
        minimum=1,
        ignore_blank=True,
        error_message="Parcelas devem ser 1 ou maior."
    )

    comp_fatura_validation = DataValidationPlan(
        validate="date",
        criteria="between",
        minimum="1",
        maximum="2958465",
        ignore_blank=True,
        error_message="A data de competência é inválida."
    )

    from .lancamentos_semantics import (
        build_status_formula, build_sys_receita, build_sys_despesa,
        build_sys_valor_parcela, build_sys_compromisso_futuro,
        build_sys_caixa_conta, build_sys_caixa_destino, build_sys_cartao,
        build_status_fatura, build_sys_competencia_efetiva,
        build_sys_parcelas_efetivas, build_sys_fatura_inicial,
        build_sys_fatura_final, build_sys_valor_parcela_base,
        build_sys_valor_ultima_parcela, build_sys_ajuste_ultima_parcela,
        build_sys_credito_fatura, build_sys_pagamento_fatura
    )

    columns = [
        ColumnPlan(header="Data", number_format="dd/mm/yyyy", width=12),
        ColumnPlan(header="Descrição", width=30),
        ColumnPlan(header="Tipo", validation=tipo_validation, width=20),
        ColumnPlan(header="Categoria", validation=cat_validation, width=20),
        ColumnPlan(header="Conta", validation=conta_validation, width=15),
        ColumnPlan(header="Conta destino", validation=conta_validation, width=15),
        ColumnPlan(header="Cartão", validation=cartao_validation, width=15),
        ColumnPlan(header="Competência da fatura", validation=comp_fatura_validation, number_format="mmm/yyyy", width=20),
        ColumnPlan(header="Valor", validation=valor_validation, number_format="R$ #,##0.00", width=15),
        ColumnPlan(header="Parcelas", validation=parcelas_validation, width=10),
        ColumnPlan(header="Essencial?", validation=sim_nao_validation, width=12),
        ColumnPlan(header="Recorrente?", validation=sim_nao_validation, width=12),
        ColumnPlan(header="Status", formula=build_status_formula(), role=CellRole.FORMULA, width=25),
        ColumnPlan(header="Status fatura", formula=build_status_fatura(), role=CellRole.FORMULA, width=25),
        ColumnPlan(header="Observação", width=25),
        ColumnPlan(header="sys_CompetenciaEfetiva", formula=build_sys_competencia_efetiva(), role=CellRole.SYSTEM, hidden=True),
        ColumnPlan(header="sys_Receita", formula=build_sys_receita(), hidden=True),
        ColumnPlan(header="sys_Despesa", formula=build_sys_despesa(), hidden=True),
        ColumnPlan(header="sys_CaixaConta", formula=build_sys_caixa_conta(), hidden=True),
        ColumnPlan(header="sys_CaixaDestino", formula=build_sys_caixa_destino(), hidden=True),
        ColumnPlan(header="sys_Cartao", formula=build_sys_cartao(), hidden=True),
        ColumnPlan(header="sys_ValorParcela", formula=build_sys_valor_parcela(), hidden=True),
        ColumnPlan(header="sys_CompromissoFuturo", formula=build_sys_compromisso_futuro(), hidden=True),
        ColumnPlan(header="sys_ParcelasEfetivas", formula=build_sys_parcelas_efetivas(), hidden=True, role=CellRole.SYSTEM),
        ColumnPlan(header="sys_FaturaInicial", formula=build_sys_fatura_inicial(), hidden=True, role=CellRole.SYSTEM),
        ColumnPlan(header="sys_FaturaFinal", formula=build_sys_fatura_final(), hidden=True, role=CellRole.SYSTEM),
        ColumnPlan(header="sys_ValorParcelaBase", formula=build_sys_valor_parcela_base(), hidden=True, role=CellRole.SYSTEM),
        ColumnPlan(header="sys_ValorUltimaParcela", formula=build_sys_valor_ultima_parcela(), hidden=True, role=CellRole.SYSTEM),
        ColumnPlan(header="sys_AjusteUltimaParcela", formula=build_sys_ajuste_ultima_parcela(), hidden=True, role=CellRole.SYSTEM),
        ColumnPlan(header="sys_CreditoFatura", formula=build_sys_credito_fatura(), hidden=True, role=CellRole.SYSTEM),
        ColumnPlan(header="sys_PagamentoFatura", formula=build_sys_pagamento_fatura(), hidden=True, role=CellRole.SYSTEM)
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

    from .defaults import DEFAULT_ACCOUNT_TYPES

    tbl_tipos_conta = TablePlan(
        name="tblTiposConta",
        start_cell="D6",
        columns=[ColumnPlan(header="Tipo")],
        data=DEFAULT_ACCOUNT_TYPES
    )

    return WorksheetPlan(
        name="Configurações",
        is_protected=False, # Intentionally unprotected for Table expansion
        show_gridlines=False,
        cells=cells,
        tables=[tbl_categorias, tbl_tipos_conta],
        column_widths={1: 25, 2: 5, 3: 20}
    )
