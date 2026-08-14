from excel_saas.core.excel.formulas import (
    if_func, and_func, or_func, isblank, isnumber, literal, equals, not_func,
    greater_than, less_than, greater_or_equal, less_or_equal, countifs, add, date_func
)
from excel_saas.core.excel.references import ThisRowRef, TableRef

CATEGORIAS_DIVIDAS = [
    "Financiamento imobiliário",
    "Financiamento de veículo",
    "Empréstimo pessoal",
    "Consignado",
    "Dívida tributária",
    "Dívida com pessoa",
    "Outros"
]

def _build_is_valid_date(val):
    return and_func(
        isnumber(val),
        greater_or_equal(val, literal(1)),
        less_than(val, add(date_func(literal(9999), literal(12), literal(31)), literal(1)))
    )

def _build_is_valid_debt():
    divida = ThisRowRef("Dívida")
    categoria = ThisRowRef("Categoria")
    saldo = ThisRowRef("Saldo devedor atual")
    parcela = ThisRowRef("Parcela mensal atual")
    data_final = ThisRowRef("Data final")
    
    divida_populated = not_func(isblank(divida))
    
    cat_checks = [equals(categoria, literal(cat)) for cat in CATEGORIAS_DIVIDAS]
    is_valid_cat = or_func(*cat_checks)
    
    is_unique = equals(countifs(TableRef("tblDividas", "Dívida"), divida), literal(1))
    
    is_valid_saldo = and_func(not_func(isblank(saldo)), isnumber(saldo), greater_than(saldo, literal(0)))
    
    is_valid_parcela = or_func(
        isblank(parcela),
        and_func(isnumber(parcela), greater_or_equal(parcela, literal(0)))
    )
    
    is_valid_data = or_func(
        isblank(data_final),
        _build_is_valid_date(data_final)
    )
    
    return and_func(
        divida_populated,
        is_valid_cat,
        is_unique,
        is_valid_saldo,
        is_valid_parcela,
        is_valid_data
    )

def build_status():
    divida = ThisRowRef("Dívida")
    categoria = ThisRowRef("Categoria")
    credor = ThisRowRef("Credor")
    saldo = ThisRowRef("Saldo devedor atual")
    parcela = ThisRowRef("Parcela mensal atual")
    data_final = ThisRowRef("Data final")
    obs = ThisRowRef("Observação")
    
    is_empty_row = and_func(
        isblank(divida),
        isblank(categoria),
        isblank(credor),
        isblank(saldo),
        isblank(parcela),
        isblank(data_final),
        isblank(obs)
    )
    
    cat_checks = [equals(categoria, literal(cat)) for cat in CATEGORIAS_DIVIDAS]
    is_valid_cat = or_func(*cat_checks)
    
    duplicate_divida = greater_than(countifs(TableRef("tblDividas", "Dívida"), divida), literal(1))
    
    invalid_saldo = or_func(not_func(isnumber(saldo)), less_or_equal(saldo, literal(0)))
    
    invalid_parcela = and_func(not_func(isblank(parcela)), or_func(not_func(isnumber(parcela)), less_than(parcela, literal(0))))
    
    invalid_data = and_func(not_func(isblank(data_final)), not_func(_build_is_valid_date(data_final)))
    
    return if_func(
        is_empty_row, literal(""),
        if_func(
            isblank(divida), literal("Informe a dívida"),
            if_func(
                isblank(categoria), literal("Informe a categoria"),
                if_func(
                    not_func(is_valid_cat), literal("Categoria inválida"),
                    if_func(
                        duplicate_divida, literal("Dívida duplicada"),
                        if_func(
                            isblank(saldo), literal("Informe o saldo devedor"),
                            if_func(
                                invalid_saldo, literal("Saldo devedor inválido"),
                                if_func(
                                    invalid_parcela, literal("Parcela mensal inválida"),
                                    if_func(
                                        invalid_data, literal("Data final inválida"),
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

def build_sys_saldo_devedor_valido():
    return if_func(
        _build_is_valid_debt(),
        ThisRowRef("Saldo devedor atual"),
        literal(0)
    )

def build_sys_parcela_mensal_valida():
    return if_func(
        _build_is_valid_debt(),
        if_func(
            isblank(ThisRowRef("Parcela mensal atual")),
            literal(0),
            ThisRowRef("Parcela mensal atual")
        ),
        literal(0)
    )
