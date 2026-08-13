from excel_saas.core.models.workbook_plan import WorksheetPlan, TablePlan, ColumnPlan, CellPlan, CellRole, DataValidationPlan
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.core.excel.formulas import (
    Expression,
    if_func, isblank, isnumber, and_func, not_func, literal,
    greater_than, less_than, greater_or_equal, less_or_equal, equals,
    int_func, countifs, sumifs, date_func, year_func, month_func, day_func,
    eomonth, edate, min_func, subtract, add, concat, today_func, multiply
)
from excel_saas.core.excel.references import ThisRowRef, TableRef, DefinedNameRef

def is_valid_date(val: Expression) -> Expression:
    return and_func(
        isnumber(val),
        greater_or_equal(val, literal(1)),
        less_than(val, add(date_func(literal(9999), literal(12), literal(31)), literal(1)))
    )

def build_sys_competencia_normalizada() -> Expression:
    comp = ThisRowRef("Competência")
    valid_comp = and_func(not_func(isblank(comp)), is_valid_date(comp))
    norm = date_func(year_func(int_func(comp)), month_func(int_func(comp)), literal(1))
    return if_func(valid_comp, norm, literal(""))

def _build_is_valid_fatura() -> Expression:
    cartao = ThisRowRef("Cartão")
    comp_norm = ThisRowRef("sys_CompetenciaNormalizada")
    
    cartao_populated = not_func(isblank(cartao))
    unique_card = equals(countifs(TableRef("tblCartoes", "Nome"), cartao), literal(1))
    comp_valid = not_func(equals(comp_norm, literal("")))
    unique_fatura = equals(countifs(TableRef("tblFaturas", "Cartão"), cartao, TableRef("tblFaturas", "sys_CompetenciaNormalizada"), comp_norm), literal(1))
    
    return and_func(cartao_populated, unique_card, comp_valid, unique_fatura)

def build_status_formula() -> Expression:
    cartao = ThisRowRef("Cartão")
    comp = ThisRowRef("Competência")
    comp_norm = ThisRowRef("sys_CompetenciaNormalizada")
    
    both_blank = and_func(isblank(cartao), isblank(comp))
    
    card_missing = isblank(cartao)
    
    card_count = countifs(TableRef("tblCartoes", "Nome"), cartao)
    card_unregistered = equals(card_count, literal(0))
    card_duplicate = greater_than(card_count, literal(1))
    
    comp_missing = isblank(comp)
    comp_invalid = not_func(is_valid_date(comp))
    
    fatura_count = countifs(TableRef("tblFaturas", "Cartão"), cartao, TableRef("tblFaturas", "sys_CompetenciaNormalizada"), comp_norm)
    fatura_duplicate = greater_than(fatura_count, literal(1))
    
    return if_func(both_blank, literal(""),
        if_func(card_missing, literal("Informe o cartão"),
            if_func(card_unregistered, literal("Cartão não cadastrado"),
                if_func(card_duplicate, literal("Cartão duplicado"),
                    if_func(comp_missing, literal("Informe a competência"),
                        if_func(comp_invalid, literal("Competência inválida"),
                            if_func(fatura_duplicate, literal("Fatura duplicada"), literal("OK"))))))))

def build_fechamento() -> Expression:
    cartao = ThisRowRef("Cartão")
    comp_norm = ThisRowRef("sys_CompetenciaNormalizada")
    is_valid = _build_is_valid_fatura()
    
    safe_fechamento = sumifs(TableRef("tblCartoes", "sys_DiaFechamentoSeguro"), TableRef("tblCartoes", "Nome"), cartao)
    has_fechamento = greater_than(safe_fechamento, literal(0))
    
    eff_fechamento = min_func(safe_fechamento, day_func(eomonth(comp_norm, literal(0))))
    date_result = date_func(year_func(comp_norm), month_func(comp_norm), eff_fechamento)
    
    return if_func(and_func(is_valid, has_fechamento), date_result, literal(""))

def build_vencimento() -> Expression:
    cartao = ThisRowRef("Cartão")
    comp_norm = ThisRowRef("sys_CompetenciaNormalizada")
    is_valid = _build_is_valid_fatura()
    
    safe_fechamento = sumifs(TableRef("tblCartoes", "sys_DiaFechamentoSeguro"), TableRef("tblCartoes", "Nome"), cartao)
    safe_vencimento = sumifs(TableRef("tblCartoes", "sys_DiaVencimentoSeguro"), TableRef("tblCartoes", "Nome"), cartao)
    
    has_both = and_func(greater_than(safe_fechamento, literal(0)), greater_than(safe_vencimento, literal(0)))
    
    target_month_is_same = greater_than(safe_vencimento, safe_fechamento)
    target_date = if_func(target_month_is_same, comp_norm, edate(comp_norm, literal(1)))
    
    eff_vencimento = min_func(safe_vencimento, day_func(eomonth(target_date, literal(0))))
    date_result = date_func(year_func(target_date), month_func(target_date), eff_vencimento)
    
    next_month_index = add(add(multiply(year_func(comp_norm), literal(12)), month_func(comp_norm)), literal(1))
    can_eval = if_func(not_func(target_month_is_same), less_or_equal(next_month_index, literal(120000)), literal(True))
    
    return if_func(and_func(is_valid, has_both, can_eval), date_result, literal(""))

def build_compras() -> Expression:
    cartao = ThisRowRef("Cartão")
    comp_norm = ThisRowRef("sys_CompetenciaNormalizada")
    is_valid = _build_is_valid_fatura()
    
    base = sumifs(
        TableRef("tblLancamentos", "sys_ValorParcelaBase"),
        TableRef("tblLancamentos", "Cartão"), cartao,
        TableRef("tblLancamentos", "sys_FaturaInicial"), concat(literal("<="), comp_norm),
        TableRef("tblLancamentos", "sys_FaturaFinal"), concat(literal(">="), comp_norm)
    )
    
    adj = sumifs(
        TableRef("tblLancamentos", "sys_AjusteUltimaParcela"),
        TableRef("tblLancamentos", "Cartão"), cartao,
        TableRef("tblLancamentos", "sys_FaturaFinal"), comp_norm
    )
    
    return if_func(is_valid, add(base, adj), literal(""))

def build_creditos() -> Expression:
    cartao = ThisRowRef("Cartão")
    comp_norm = ThisRowRef("sys_CompetenciaNormalizada")
    is_valid = _build_is_valid_fatura()
    
    cred = sumifs(
        TableRef("tblLancamentos", "sys_CreditoFatura"),
        TableRef("tblLancamentos", "Cartão"), cartao,
        TableRef("tblLancamentos", "sys_CompetenciaEfetiva"), comp_norm
    )
    return if_func(is_valid, cred, literal(""))

def build_total() -> Expression:
    compras = ThisRowRef("Compras / Parcelas")
    creditos = ThisRowRef("Créditos / Estornos")
    is_valid = _build_is_valid_fatura()
    
    return if_func(is_valid, subtract(compras, creditos), literal(""))

def build_pagamentos() -> Expression:
    cartao = ThisRowRef("Cartão")
    comp_norm = ThisRowRef("sys_CompetenciaNormalizada")
    is_valid = _build_is_valid_fatura()
    
    pag = sumifs(
        TableRef("tblLancamentos", "sys_PagamentoFatura"),
        TableRef("tblLancamentos", "Cartão"), cartao,
        TableRef("tblLancamentos", "sys_CompetenciaEfetiva"), comp_norm
    )
    return if_func(is_valid, pag, literal(""))

def build_em_aberto() -> Expression:
    total = ThisRowRef("Total da fatura")
    pagamentos = ThisRowRef("Pagamentos")
    is_valid = _build_is_valid_fatura()
    
    return if_func(is_valid, subtract(total, pagamentos), literal(""))

def build_situacao() -> Expression:
    is_valid = _build_is_valid_fatura()
    compras = ThisRowRef("Compras / Parcelas")
    creditos = ThisRowRef("Créditos / Estornos")
    pagamentos = ThisRowRef("Pagamentos")
    em_aberto = ThisRowRef("Em aberto")
    venc = ThisRowRef("Vencimento")
    
    no_activity = and_func(equals(compras, literal(0)), equals(creditos, literal(0)), equals(pagamentos, literal(0)))
    
    return if_func(not_func(is_valid), literal(""),
        if_func(no_activity, literal("Sem movimento"),
            if_func(less_than(em_aberto, literal(0)), literal("Crédito"),
                if_func(equals(em_aberto, literal(0)), literal("Paga"),
                    if_func(equals(venc, literal("")), literal("Em aberto — sem vencimento"),
                        if_func(greater_than(today_func(), venc), literal("Vencida"), literal("Em aberto")))))))

def build_faturas(request: GenerationRequest) -> WorksheetPlan:
    cartao_validation = DataValidationPlan(validate="list", source=DefinedNameRef("lista_cartoes"), ignore_blank=True)
    comp_fatura_validation = DataValidationPlan(
        validate="date",
        criteria="between",
        minimum=1,
        maximum=2958465,
        ignore_blank=True,
        error_message="Insira uma data válida para a competência."
    )
    
    columns = [
        ColumnPlan(header="Cartão", validation=cartao_validation, width=20),
        ColumnPlan(header="Competência", validation=comp_fatura_validation, number_format="mmm/yyyy", width=20),
        ColumnPlan(header="Fechamento", formula=build_fechamento(), role=CellRole.FORMULA, number_format="dd/mm/yyyy", width=15),
        ColumnPlan(header="Vencimento", formula=build_vencimento(), role=CellRole.FORMULA, number_format="dd/mm/yyyy", width=15),
        ColumnPlan(header="Compras / Parcelas", formula=build_compras(), role=CellRole.FORMULA, number_format="R$ #,##0.00", width=20),
        ColumnPlan(header="Créditos / Estornos", formula=build_creditos(), role=CellRole.FORMULA, number_format="R$ #,##0.00", width=20),
        ColumnPlan(header="Total da fatura", formula=build_total(), role=CellRole.FORMULA, number_format="R$ #,##0.00", width=20),
        ColumnPlan(header="Pagamentos", formula=build_pagamentos(), role=CellRole.FORMULA, number_format="R$ #,##0.00", width=20),
        ColumnPlan(header="Em aberto", formula=build_em_aberto(), role=CellRole.FORMULA, number_format="R$ #,##0.00", width=20),
        ColumnPlan(header="Status", formula=build_status_formula(), role=CellRole.FORMULA, width=25),
        ColumnPlan(header="Situação", formula=build_situacao(), role=CellRole.FORMULA, width=25),
        ColumnPlan(header="sys_CompetenciaNormalizada", formula=build_sys_competencia_normalizada(), role=CellRole.SYSTEM, hidden=True),
    ]
    
    table = TablePlan(
        name="tblFaturas",
        start_cell="B4",
        columns=columns,
        data=[],
        show_total_row=False
    )
    
    cells = [
        CellPlan(row=1, col=1, value="Faturas", role=CellRole.TITLE, size=16, bold=True),
        CellPlan(row=2, col=1, value="Adicione uma linha para cada cartão e competência que deseja acompanhar.", role=CellRole.NORMAL),
    ]
    
    return WorksheetPlan(
        name="Faturas",
        is_protected=False,
        show_gridlines=False,
        cells=cells,
        tables=[table],
        column_widths={0: 3} # padding
    )
