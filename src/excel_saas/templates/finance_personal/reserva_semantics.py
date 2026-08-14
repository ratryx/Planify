from excel_saas.core.excel.formulas import (
    if_func, equals, isblank, and_func, not_func, literal, or_func,
    int_func, isnumber, greater_or_equal, less_than,
    greater_than, add, subtract, multiply, divide, group,
    min_func, row_func, less_or_equal
)
from excel_saas.core.excel.references import ThisRowRef, TableRef

custo = ThisRowRef("Custo essencial mensal")
meses = ThisRowRef("Meses desejados")
reserva = ThisRowRef("Reserva atual")
alvo = ThisRowRef("Reserva alvo")
falta = ThisRowRef("Falta")

def build_row_index():
    return add(subtract(row_func(), min_func(row_func(TableRef("tblReserva", "Custo essencial mensal")))), literal(1))

def _build_is_valid_reserva():
    row_idx_is_one = equals(build_row_index(), literal(1))
    
    custo_valid = and_func(not_func(isblank(custo)), isnumber(custo), greater_than(custo, literal(0)))
    meses_valid = and_func(not_func(isblank(meses)), isnumber(meses), greater_or_equal(meses, literal(1)), equals(meses, int_func(meses)))
    reserva_valid = and_func(not_func(isblank(reserva)), isnumber(reserva), greater_or_equal(reserva, literal(0)))
    
    return and_func(row_idx_is_one, custo_valid, meses_valid, reserva_valid)

def build_status():
    row_idx_invalid = greater_than(build_row_index(), literal(1))
    all_blank = and_func(isblank(custo), isblank(meses), isblank(reserva))
    
    return if_func(row_idx_invalid, literal("Use apenas uma linha"),
        if_func(all_blank, literal(""),
            if_func(isblank(custo), literal("Informe o custo essencial"),
                if_func(or_func(not_func(isnumber(custo)), less_or_equal(custo, literal(0))), literal("Custo essencial inválido"),
                    if_func(isblank(meses), literal("Informe os meses"),
                        if_func(or_func(not_func(isnumber(meses)), less_than(meses, literal(1)), not_func(equals(meses, int_func(meses)))), literal("Meses inválidos"),
                            if_func(isblank(reserva), literal("Informe a reserva atual"),
                                if_func(or_func(not_func(isnumber(reserva)), less_than(reserva, literal(0))), literal("Reserva atual inválida"),
                                    literal("OK")
                                )
                            )
                        )
                    )
                )
            )
        )
    )

def build_reserva_alvo():
    calc = multiply(custo, meses)
    return if_func(_build_is_valid_reserva(), calc, literal(""))

def build_falta():
    calc = if_func(greater_or_equal(reserva, alvo), literal(0), subtract(alvo, reserva))
    return if_func(_build_is_valid_reserva(), calc, literal(""))

def build_cobertura_atual():
    calc = divide(reserva, custo)
    return if_func(_build_is_valid_reserva(), calc, literal(""))

def build_progresso():
    calc = divide(reserva, alvo)
    return if_func(_build_is_valid_reserva(), calc, literal(""))

def build_situacao():
    is_zero = equals(reserva, literal(0))
    is_concluded = greater_or_equal(reserva, alvo)
    
    calc = if_func(is_zero, literal("Não iniciada"),
        if_func(is_concluded, literal("Completa"),
            literal("Em formação")
        )
    )
    return if_func(_build_is_valid_reserva(), calc, literal(""))
