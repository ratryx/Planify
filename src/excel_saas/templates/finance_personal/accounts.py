from typing import List
from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, DataValidationPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.core.excel.references import DefinedNameRef
from .defaults import DEFAULT_ACCOUNTS_SAMPLE
from excel_saas.core.excel.formulas import sum_func, if_func, isblank, literal, sumifs, greater_than, countifs, Expression
from excel_saas.core.excel.references import ThisRowRef, TableRef

def build_saldo_atual_formula() -> Expression:
    from excel_saas.core.excel.formulas import and_func, or_func, not_func, isnumber
    nome = ThisRowRef("Nome")
    saldo_inicial = ThisRowRef("Saldo inicial")
    sys_caixa_conta = TableRef("tblLancamentos", "sys_CaixaConta")
    sys_caixa_destino = TableRef("tblLancamentos", "sys_CaixaDestino")
    lanc_conta = TableRef("tblLancamentos", "Conta")
    lanc_conta_destino = TableRef("tblLancamentos", "Conta destino")

    is_safe = and_func(
        not_func(isblank(nome)),
        or_func(isblank(saldo_inicial), isnumber(saldo_inicial))
    )

    calc = sum_func(
        saldo_inicial,
        sumifs(sys_caixa_conta, lanc_conta, nome),
        sumifs(sys_caixa_destino, lanc_conta_destino, nome)
    )
    return if_func(is_safe, calc, literal(""))

def build_status_formula() -> Expression:
    from excel_saas.core.excel.formulas import and_func, isnumber, not_func

    nome = ThisRowRef("Nome")
    tipo = ThisRowRef("Tipo")
    inst = ThisRowRef("Instituição")
    saldo_ini = ThisRowRef("Saldo inicial")
    incluir = ThisRowRef("Incluir no saldo disponível?")
    ativa = ThisRowRef("Ativa?")

    is_empty_row = and_func(
        isblank(nome),
        isblank(tipo),
        isblank(inst),
        isblank(saldo_ini),
        isblank(incluir),
        isblank(ativa)
    )

    tbl_nome_col = TableRef("tblContas", "Nome")

    check_saldo = if_func(
        and_func(not_func(isblank(saldo_ini)), not_func(isnumber(saldo_ini))),
        literal("Saldo inicial inválido"),
        literal("OK")
    )

    check_dupe = if_func(
        greater_than(countifs(tbl_nome_col, nome), literal(1)),
        literal("Nome duplicado"),
        check_saldo
    )

    return if_func(
        is_empty_row,
        literal(""),
        if_func(isblank(nome), literal("Informe o nome"), check_dupe)
    )

def build_contas(request: GenerationRequest) -> WorksheetPlan:
    sim_nao_validation = DataValidationPlan(validate="list", source=["Sim", "Não"])
    tipo_validation = DataValidationPlan(validate="list", source=DefinedNameRef("lista_tipos_conta"))
    # We use Excel's exact technical limits for cell input rather than an arbitrary financial bound.
    # Microsoft documents the exact cell numeric limits as: -9.99999999999999E+307 to 9.99999999999999E+307.
    saldo_validation = DataValidationPlan(
        validate="decimal",
        criteria="between",
        minimum=-9.99999999999999e307,
        maximum=9.99999999999999e307,
        ignore_blank=True,
        error_message="O saldo inicial deve ser um número válido."
    )

    columns = [
        ColumnPlan(header="Nome", width=25),
        ColumnPlan(header="Tipo", validation=tipo_validation, width=20),
        ColumnPlan(header="Instituição", width=20),
        ColumnPlan(header="Saldo inicial", number_format="R$ #,##0.00", validation=saldo_validation, width=15),
        ColumnPlan(header="Saldo atual", formula=build_saldo_atual_formula(), role=CellRole.FORMULA, number_format="R$ #,##0.00", width=15),
        ColumnPlan(header="Incluir no saldo disponível?", validation=sim_nao_validation, width=28),
        ColumnPlan(header="Ativa?", validation=sim_nao_validation, width=12),
        ColumnPlan(header="Status", formula=build_status_formula(), role=CellRole.FORMULA, width=20),
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
