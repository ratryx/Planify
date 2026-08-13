from typing import List
from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, DataValidationPlan
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.core.excel.references import DefinedNameRef
from .defaults import DEFAULT_CARDS_SAMPLE
from excel_saas.core.excel.formulas import Expression
def build_status_formula() -> Expression:
    from excel_saas.core.excel.formulas import (
        if_func, isblank, and_func, not_func, isnumber, literal,
        greater_than, less_than, countifs, or_func, not_equals, int_func
    )
    from excel_saas.core.excel.references import ThisRowRef, TableRef

    nome = ThisRowRef("Nome")
    limite = ThisRowRef("Limite")
    fechamento = ThisRowRef("Dia fechamento")
    vencimento = ThisRowRef("Dia vencimento")
    conta = ThisRowRef("Conta de pagamento")
    ativo = ThisRowRef("Ativo?")

    is_empty = and_func(
        isblank(nome),
        isblank(limite),
        isblank(fechamento),
        isblank(vencimento),
        isblank(conta),
        isblank(ativo)
    )

    check_nome = if_func(isblank(nome), literal("Informe o nome"),
        if_func(greater_than(countifs(TableRef("tblCartoes", "Nome"), nome), literal(1)), literal("Nome duplicado"),
            literal("OK")))

    check_limite = if_func(not_func(isblank(limite)),
        if_func(or_func(not_func(isnumber(limite)), less_than(limite, literal(0))), literal("Limite inválido"), literal("OK")),
        literal("OK"))

    check_fechamento = if_func(not_func(isblank(fechamento)),
        if_func(or_func(not_func(isnumber(fechamento)), not_equals(fechamento, int_func(fechamento)), less_than(fechamento, literal(1)), greater_than(fechamento, literal(31))), literal("Fechamento inválido"), literal("OK")),
        literal("OK"))

    check_vencimento = if_func(not_func(isblank(vencimento)),
        if_func(or_func(not_func(isnumber(vencimento)), not_equals(vencimento, int_func(vencimento)), less_than(vencimento, literal(1)), greater_than(vencimento, literal(31))), literal("Vencimento inválido"), literal("OK")),
        literal("OK"))

    main_logic = if_func(not_equals(check_nome, literal("OK")), check_nome,
        if_func(not_equals(check_limite, literal("OK")), check_limite,
            if_func(not_equals(check_fechamento, literal("OK")), check_fechamento,
                if_func(not_equals(check_vencimento, literal("OK")), check_vencimento, literal("OK")))))

    return if_func(is_empty, literal(""), main_logic)

def build_sys_dia_fechamento_seguro() -> Expression:
    from excel_saas.core.excel.formulas import (
        if_func, isblank, not_func, isnumber, literal,
        greater_than, less_than, countifs, and_func, equals, int_func
    )
    from excel_saas.core.excel.references import ThisRowRef, TableRef
    
    nome = ThisRowRef("Nome")
    fechamento = ThisRowRef("Dia fechamento")
    
    nome_valid = and_func(not_func(isblank(nome)), equals(countifs(TableRef("tblCartoes", "Nome"), nome), literal(1)))
    fechamento_valid = and_func(not_func(isblank(fechamento)), isnumber(fechamento), equals(fechamento, int_func(fechamento)), greater_than(fechamento, literal(0)), less_than(fechamento, literal(32)))
    
    is_safe = and_func(nome_valid, fechamento_valid)
    return if_func(is_safe, fechamento, literal(""))

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
        ColumnPlan(header="Status", formula=build_status_formula(), role=CellRole.FORMULA, width=20),
        ColumnPlan(header="sys_DiaFechamentoSeguro", formula=build_sys_dia_fechamento_seguro(), role=CellRole.SYSTEM, hidden=True),
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
