from excel_saas.core.excel.formulas import (
    if_func, equals, isblank, and_func, not_func, literal, or_func,
    isnumber, greater_or_equal, less_than,
    countifs, greater_than, add, subtract, divide, group,
    sum_func, less_or_equal
)
from excel_saas.core.excel.references import ThisRowRef, TableRef

ativo = ThisRowRef("Ativo")
classe = ThisRowRef("Classe")
instituicao = ThisRowRef("Instituição")
aportado = ThisRowRef("Total aportado")
recebido = ThisRowRef("Total recebido")
atual = ThisRowRef("Valor atual")

CLASSES_PERMITIDAS = [
    "Renda fixa", "Ações", "FIIs", "ETFs", "Fundos", "Cripto", "Previdência", "Outros"
]

def _build_is_valid_classe():
    conditions = [equals(classe, literal(c)) for c in CLASSES_PERMITIDAS]
    return or_func(*conditions)

def _build_is_duplicate():
    return greater_than(
        countifs(TableRef("tblInvestimentos", "Ativo"), ativo, TableRef("tblInvestimentos", "Instituição"), instituicao),
        literal(1)
    )

def _build_is_valid_investment():
    ativo_valid = not_func(isblank(ativo))
    classe_valid = and_func(not_func(isblank(classe)), _build_is_valid_classe())
    duplicate_valid = not_func(_build_is_duplicate())
    aportado_valid = and_func(not_func(isblank(aportado)), isnumber(aportado), greater_than(aportado, literal(0)))
    recebido_valid = and_func(not_func(isblank(recebido)), isnumber(recebido), greater_or_equal(recebido, literal(0)))
    atual_valid = and_func(not_func(isblank(atual)), isnumber(atual), greater_or_equal(atual, literal(0)))
    
    return and_func(ativo_valid, classe_valid, duplicate_valid, aportado_valid, recebido_valid, atual_valid)

def build_status():
    all_blank = and_func(isblank(ativo), isblank(classe), isblank(instituicao), isblank(aportado), isblank(recebido), isblank(atual))
    
    return if_func(all_blank, literal(""),
        if_func(isblank(ativo), literal("Informe o ativo"),
            if_func(isblank(classe), literal("Informe a classe"),
                if_func(not_func(_build_is_valid_classe()), literal("Classe inválida"),
                    if_func(_build_is_duplicate(), literal("Investimento duplicado"),
                        if_func(isblank(aportado), literal("Informe o total aportado"),
                            if_func(or_func(not_func(isnumber(aportado)), less_or_equal(aportado, literal(0))), literal("Total aportado inválido"),
                                if_func(isblank(recebido), literal("Informe o total recebido"),
                                    if_func(or_func(not_func(isnumber(recebido)), less_than(recebido, literal(0))), literal("Total recebido inválido"),
                                        if_func(isblank(atual), literal("Informe o valor atual"),
                                            if_func(or_func(not_func(isnumber(atual)), less_than(atual, literal(0))), literal("Valor atual inválido"),
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
        )
    )

def build_resultado_total():
    # subtract(group(add(valor_atual, total_recebido)), total_aportado)
    calc = subtract(group(add(atual, recebido)), aportado)
    return if_func(_build_is_valid_investment(), calc, literal(""))

def build_retorno_simples():
    calc = divide(ThisRowRef("Resultado total"), aportado)
    return if_func(_build_is_valid_investment(), calc, literal(""))

def build_peso_carteira():
    # IF(NOT(valid), "", IF(SUM(sys_ValorAtualValido)=0, 0, Valor atual / SUM(sys_ValorAtualValido)))
    valid = _build_is_valid_investment()
    sum_valid_atual = sum_func(TableRef("tblInvestimentos", "sys_ValorAtualValido"))
    calc = if_func(equals(sum_valid_atual, literal(0)), literal(0), divide(atual, sum_valid_atual))
    
    return if_func(not_func(valid), literal(""), calc)

def build_situacao():
    resultado = ThisRowRef("Resultado total")
    calc = if_func(greater_than(resultado, literal(0)), literal("Ganho"),
        if_func(less_than(resultado, literal(0)), literal("Perda"),
            literal("No zero")
        )
    )
    return if_func(_build_is_valid_investment(), calc, literal(""))

def build_sys_aporte_valido():
    return if_func(_build_is_valid_investment(), aportado, literal(0))

def build_sys_recebido_valido():
    return if_func(_build_is_valid_investment(), recebido, literal(0))

def build_sys_valor_atual_valido():
    return if_func(_build_is_valid_investment(), atual, literal(0))
