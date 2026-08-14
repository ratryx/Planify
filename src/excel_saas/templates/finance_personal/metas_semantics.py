from excel_saas.core.excel.formulas import (
    if_func, equals, isblank, and_func, not_func, literal, or_func,
    date_func, year_func, month_func, int_func, isnumber, greater_or_equal, less_than,
    countifs, greater_than, add, subtract, multiply, divide, group, today_func, less_or_equal
)
from excel_saas.core.excel.references import ThisRowRef, TableRef

meta = ThisRowRef("Meta")
valor_alvo = ThisRowRef("Valor alvo")
valor_atual = ThisRowRef("Valor atual")
data_alvo = ThisRowRef("Data alvo")

falta = ThisRowRef("Falta")
meses_restantes = ThisRowRef("Meses restantes")

def is_safe_date(val):
    return and_func(
        isnumber(val),
        greater_or_equal(val, literal(1)),
        less_than(val, literal(2958466))
    )

def _build_is_valid_goal():
    meta_populated = not_func(isblank(meta))
    meta_unique = equals(countifs(TableRef("tblMetas", "Meta"), meta), literal(1))
    
    alvo_valid = and_func(not_func(isblank(valor_alvo)), isnumber(valor_alvo), greater_than(valor_alvo, literal(0)))
    atual_valid = and_func(not_func(isblank(valor_atual)), isnumber(valor_atual), greater_or_equal(valor_atual, literal(0)))
    
    data_valid = and_func(not_func(isblank(data_alvo)), is_safe_date(data_alvo))
    
    return and_func(meta_populated, meta_unique, alvo_valid, atual_valid, data_valid)

def build_status():
    all_blank = and_func(isblank(meta), isblank(valor_alvo), isblank(valor_atual), isblank(data_alvo))
    meta_count = countifs(TableRef("tblMetas", "Meta"), meta)
    
    return if_func(all_blank, literal(""),
        if_func(isblank(meta), literal("Informe a meta"),
            if_func(greater_than(meta_count, literal(1)), literal("Meta duplicada"),
                if_func(isblank(valor_alvo), literal("Informe o valor alvo"),
                    if_func(or_func(not_func(isnumber(valor_alvo)), less_or_equal(valor_alvo, literal(0))), literal("Valor alvo inválido"),
                        if_func(isblank(valor_atual), literal("Informe o valor atual"),
                            if_func(or_func(not_func(isnumber(valor_atual)), less_than(valor_atual, literal(0))), literal("Valor atual inválido"),
                                if_func(isblank(data_alvo), literal("Informe a data alvo"),
                                    if_func(not_func(is_safe_date(data_alvo)), literal("Data alvo inválida"),
                                        literal("OK")
                                    )
                                )
                            )
                        )
                    )
                )
            )
        )
    )

def build_falta():
    calc = if_func(greater_or_equal(valor_atual, valor_alvo), literal(0), subtract(valor_alvo, valor_atual))
    return if_func(_build_is_valid_goal(), calc, literal(""))

def build_progresso():
    calc = divide(valor_atual, valor_alvo)
    return if_func(_build_is_valid_goal(), calc, literal(""))

def build_meses_restantes():
    current_month_index = add(multiply(year_func(today_func()), literal(12)), month_func(today_func()))
    target_month_index = add(multiply(year_func(int_func(data_alvo)), literal(12)), month_func(int_func(data_alvo)))
    
    # Needs to be explicitly grouped
    math = add(subtract(group(target_month_index), group(current_month_index)), literal(1))
    
    calc = if_func(greater_or_equal(valor_atual, valor_alvo), literal(0),
        if_func(less_than(target_month_index, current_month_index), literal(0),
            math
        )
    )
    return if_func(_build_is_valid_goal(), calc, literal(""))

def build_aporte_mensal():
    calc = if_func(equals(falta, literal(0)), literal(0),
        if_func(equals(meses_restantes, literal(0)), literal(""),
            divide(falta, meses_restantes)
        )
    )
    return if_func(_build_is_valid_goal(), calc, literal(""))

def build_situacao():
    current_month_index = add(multiply(year_func(today_func()), literal(12)), month_func(today_func()))
    target_month_index = add(multiply(year_func(int_func(data_alvo)), literal(12)), month_func(int_func(data_alvo)))
    
    is_concluded = greater_or_equal(valor_atual, valor_alvo)
    is_overdue = greater_than(current_month_index, target_month_index)
    
    calc = if_func(is_concluded, literal("Concluída"),
        if_func(is_overdue, literal("Prazo vencido"),
            literal("Em andamento")
        )
    )
    return if_func(_build_is_valid_goal(), calc, literal(""))
