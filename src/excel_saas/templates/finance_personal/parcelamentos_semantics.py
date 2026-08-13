from excel_saas.core.excel.formulas import (
    index_func, aggregate_func, row_func, iferror_func, equals,
    and_func, not_func, literal, isblank, greater_than, not_equals, less_than,
    subtract, add, multiply, year_func, month_func, date_func, today_func,
    divide, min_func, if_func, group
)
from excel_saas.core.excel.references import TableRef, ThisRowRef

# Destination References (tblParcelamentos)
sys_ordem = ThisRowRef("sys_Ordem")
sys_indice_lancamento = ThisRowRef("sys_IndiceLancamento")
parcelas_restantes = ThisRowRef("Parcelas restantes")
valor_da_parcela = ThisRowRef("Valor da parcela")
ajuste_ultima_parcela = ThisRowRef("sys_AjusteUltimaParcela")
primeira_fatura = ThisRowRef("Primeira fatura")
ultima_fatura = ThisRowRef("Última fatura")

# Source References (tblLancamentos)
lanc_parcelas_efetivas = TableRef("tblLancamentos", "sys_ParcelasEfetivas")
lanc_fatura_inicial = TableRef("tblLancamentos", "sys_FaturaInicial")
lanc_fatura_final = TableRef("tblLancamentos", "sys_FaturaFinal")

def build_sys_indice_lancamento():
    qualifying_denominator = multiply(
        multiply(
            group(greater_than(lanc_parcelas_efetivas, literal(1))),
            group(not_equals(lanc_fatura_inicial, literal("")))
        ),
        group(not_equals(lanc_fatura_final, literal("")))
    )

    relative_source_index = add(
        subtract(
            row_func(lanc_parcelas_efetivas),
            min_func(row_func(lanc_parcelas_efetivas))
        ),
        literal(1)
    )

    array_computation = divide(group(relative_source_index), group(qualifying_denominator))
    
    return iferror_func(
        aggregate_func(literal(15), literal(6), array_computation, sys_ordem),
        literal("")
    )

def build_projected_column(source_col_name: str):
    return if_func(
        equals(sys_indice_lancamento, literal("")),
        literal(""),
        index_func(TableRef("tblLancamentos", source_col_name), sys_indice_lancamento)
    )

def _build_current_month():
    return date_func(year_func(today_func()), month_func(today_func()), literal(1))

def build_parcelas_restantes():
    current_month = _build_current_month()
    
    current_index = add(
        multiply(year_func(current_month), literal(12)),
        month_func(current_month)
    )
    
    final_index = add(
        multiply(year_func(ultima_fatura), literal(12)),
        month_func(ultima_fatura)
    )
    
    calc_remaining = add(subtract(group(final_index), group(current_index)), literal(1))
    
    logic = if_func(
        less_than(current_month, primeira_fatura),
        ThisRowRef("Parcelas"),
        if_func(
            greater_than(current_month, ultima_fatura),
            literal(0),
            calc_remaining
        )
    )
    
    return if_func(
        equals(sys_indice_lancamento, literal("")),
        literal(""),
        logic
    )

def build_proxima_competencia():
    current_month = _build_current_month()
    
    logic = if_func(
        less_than(current_month, primeira_fatura),
        primeira_fatura,
        if_func(
            greater_than(current_month, ultima_fatura),
            literal(""),
            current_month
        )
    )
    
    return if_func(
        equals(sys_indice_lancamento, literal("")),
        literal(""),
        logic
    )

def build_compromisso_restante():
    logic = if_func(
        equals(parcelas_restantes, literal(0)),
        literal(0),
        add(
            multiply(parcelas_restantes, valor_da_parcela),
            ajuste_ultima_parcela
        )
    )
    
    return if_func(
        equals(sys_indice_lancamento, literal("")),
        literal(""),
        logic
    )

def build_situacao():
    current_month = _build_current_month()
    
    logic = if_func(
        less_than(current_month, primeira_fatura),
        literal("A iniciar"),
        if_func(
            greater_than(current_month, ultima_fatura),
            literal("Concluído"),
            literal("Em andamento")
        )
    )
    
    return if_func(
        equals(sys_indice_lancamento, literal("")),
        literal(""),
        logic
    )
