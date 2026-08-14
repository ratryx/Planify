from excel_saas.core.excel.formulas import (
    if_func, and_func, or_func, isblank, isnumber, literal, equals, not_func, greater_than,
    less_than, add, greater_or_equal, concat, countifs
)
from excel_saas.core.excel.references import ThisRowRef, TableRef

CATEGORIAS_PATRIMONIO = [
    "Imóveis",
    "Veículos",
    "Empresas / participações",
    "Bens de valor",
    "Direitos / créditos",
    "Outros"
]

def _build_is_valid_asset():
    bem = ThisRowRef("Bem")
    categoria = ThisRowRef("Categoria")
    valor_atual = ThisRowRef("Valor atual")
    
    cat_checks = [equals(categoria, literal(cat)) for cat in CATEGORIAS_PATRIMONIO]
    is_valid_cat = or_func(*cat_checks)
    
    is_unique_bem = equals(countifs(TableRef("tblBensPatrimoniais", "Bem"), bem), literal(1))
    
    is_valid_value = and_func(
        not_func(isblank(valor_atual)),
        isnumber(valor_atual),
        greater_or_equal(valor_atual, literal(0))
    )
    
    return and_func(
        not_func(isblank(bem)),
        is_valid_cat,
        is_unique_bem,
        is_valid_value
    )

def build_sys_valor_atual_valido():
    return if_func(
        _build_is_valid_asset(),
        ThisRowRef("Valor atual"),
        literal(0)
    )

def build_status():
    bem = ThisRowRef("Bem")
    categoria = ThisRowRef("Categoria")
    valor_atual = ThisRowRef("Valor atual")
    observacao = ThisRowRef("Observação")
    
    is_empty_row = and_func(
        isblank(bem),
        isblank(categoria),
        isblank(valor_atual),
        isblank(observacao)
    )
    
    cat_checks = [equals(categoria, literal(cat)) for cat in CATEGORIAS_PATRIMONIO]
    is_valid_cat = or_func(*cat_checks)
    
    duplicate_bem = greater_than(countifs(TableRef("tblBensPatrimoniais", "Bem"), bem), literal(1))
    
    invalid_value = or_func(
        not_func(isnumber(valor_atual)),
        less_than(valor_atual, literal(0))
    )
    
    return if_func(
        is_empty_row, literal(""),
        if_func(
            isblank(bem), literal("Informe o bem"),
            if_func(
                isblank(categoria), literal("Informe a categoria"),
                if_func(
                    not_func(is_valid_cat), literal("Categoria inválida"),
                    if_func(
                        duplicate_bem, literal("Bem duplicado"),
                        if_func(
                            isblank(valor_atual), literal("Informe o valor atual"),
                            if_func(
                                invalid_value, literal("Valor atual inválido"),
                                literal("OK")
                            )
                        )
                    )
                )
            )
        )
    )
