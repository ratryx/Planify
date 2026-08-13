from excel_saas.core.excel.formulas import (
    if_func, equals, isblank, and_func, not_func, literal, or_func,
    date_func, year_func, month_func, int_func, isnumber, greater_or_equal, less_than,
    countifs, greater_than, add, subtract, multiply, divide, edate, group, sumifs, concat
)
from excel_saas.core.excel.references import ThisRowRef, TableRef

competencia = ThisRowRef("Competência")
categoria = ThisRowRef("Categoria")
orcamento = ThisRowRef("Orçamento")

# Output column refs needed for other calculations
saidas_em_conta = ThisRowRef("Saídas em conta")
cartao_parcelas = ThisRowRef("Cartão / Parcelas")
consumido_comprometido = ThisRowRef("Consumido / Comprometido")
disponivel = ThisRowRef("Disponível")

lanc_data = TableRef("tblLancamentos", "Data")
lanc_categoria = TableRef("tblLancamentos", "Categoria")
lanc_tipo = TableRef("tblLancamentos", "Tipo")
lanc_conta = TableRef("tblLancamentos", "Conta")
lanc_status = TableRef("tblLancamentos", "Status")
lanc_valor = TableRef("tblLancamentos", "Valor")
lanc_sys_despesa = TableRef("tblLancamentos", "sys_Despesa")
lanc_sys_caixaconta = TableRef("tblLancamentos", "sys_CaixaConta")
lanc_sys_valorparcelabase = TableRef("tblLancamentos", "sys_ValorParcelaBase")
lanc_sys_ajusteultimaparcela = TableRef("tblLancamentos", "sys_AjusteUltimaParcela")
lanc_sys_faturainicial = TableRef("tblLancamentos", "sys_FaturaInicial")
lanc_sys_faturafinal = TableRef("tblLancamentos", "sys_FaturaFinal")
lanc_sys_creditofatura = TableRef("tblLancamentos", "sys_CreditoFatura")
lanc_sys_competenciaefetiva = TableRef("tblLancamentos", "sys_CompetenciaEfetiva")

# 7. sys_CompetenciaNormalizada
def build_sys_competencia_normalizada():
    # DATE(YEAR(INT([@Competência])),MONTH(INT([@Competência])),1)
    is_date = and_func(not_func(isblank(competencia)), isnumber(competencia), greater_or_equal(competencia, literal(1)), less_than(competencia, literal(2958466)))
    date_val = int_func(competencia)
    normalized = date_func(year_func(date_val), month_func(date_val), literal(1))
    return if_func(is_date, normalized, literal(""))

sys_competencia_normalizada = ThisRowRef("sys_CompetenciaNormalizada")

# 11. DIRECT VALID-BUDGET PREDICATE
def _build_is_valid_budget():
    comp_valid = not_func(equals(sys_competencia_normalizada, literal("")))
    cat_populated = not_func(isblank(categoria))
    cat_exists = equals(countifs(TableRef("tblCategorias", "Categoria"), categoria), literal(1))
    
    orc_populated = not_func(isblank(orcamento))
    orc_valid = and_func(orc_populated, isnumber(orcamento), greater_or_equal(orcamento, literal(0)))
    
    duplicate_count = countifs(
        TableRef("tblOrcamento", "Categoria"), categoria,
        TableRef("tblOrcamento", "sys_CompetenciaNormalizada"), sys_competencia_normalizada
    )
    is_unique = equals(duplicate_count, literal(1))
    
    return and_func(comp_valid, cat_populated, cat_exists, orc_valid, is_unique)


# 10. STATUS PRECEDENCE
def build_status():
    all_blank = and_func(isblank(competencia), isblank(categoria), isblank(orcamento))
    
    # comp valid helper
    comp_populated = not_func(isblank(competencia))
    comp_safe = and_func(isnumber(competencia), greater_or_equal(competencia, literal(1)), less_than(competencia, literal(2958466)))
    
    cat_count = countifs(TableRef("tblCategorias", "Categoria"), categoria)
    
    duplicate_count = countifs(
        TableRef("tblOrcamento", "Categoria"), categoria,
        TableRef("tblOrcamento", "sys_CompetenciaNormalizada"), sys_competencia_normalizada
    )
    
    return if_func(all_blank, literal(""),
        if_func(isblank(competencia), literal("Informe a competência"),
            if_func(not_func(comp_safe), literal("Competência inválida"),
                if_func(isblank(categoria), literal("Informe a categoria"),
                    if_func(equals(cat_count, literal(0)), literal("Categoria não cadastrada"),
                        if_func(greater_than(cat_count, literal(1)), literal("Categoria duplicada"),
                            if_func(isblank(orcamento), literal("Informe o orçamento"),
                                if_func(or_func(not_func(isnumber(orcamento)), less_than(orcamento, literal(0))), literal("Orçamento inválido"),
                                    if_func(greater_than(duplicate_count, literal(1)), literal("Orçamento duplicado"),
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

def build_saidas_em_conta():
    # Ordinary expenses/refunds (Account only)
    account_expense = sumifs(
        lanc_sys_despesa,
        lanc_categoria, categoria,
        lanc_conta, literal("<>"),
        lanc_data, concat(literal(">="), sys_competencia_normalizada),
        lanc_data, concat(literal("<"), edate(sys_competencia_normalizada, literal(1)))
    )
    
    # Investments (Account only inherently, negated to make it positive)
    investments = multiply(sumifs(
        lanc_sys_caixaconta,
        lanc_tipo, literal("Investimento"),
        lanc_categoria, categoria,
        lanc_conta, literal("<>"),
        lanc_data, concat(literal(">="), sys_competencia_normalizada),
        lanc_data, concat(literal("<"), edate(sys_competencia_normalizada, literal(1)))
    ), literal(-1))

    # Debt payments (Account only inherently, negated)
    debts = multiply(sumifs(
        lanc_sys_caixaconta,
        lanc_tipo, literal("Pagamento de dívida"),
        lanc_categoria, categoria,
        lanc_conta, literal("<>"),
        lanc_data, concat(literal(">="), sys_competencia_normalizada),
        lanc_data, concat(literal("<"), edate(sys_competencia_normalizada, literal(1)))
    ), literal(-1))

    total = add(account_expense, add(investments, debts))
    
    return if_func(_build_is_valid_budget(), total, literal(""))

def build_cartao_parcelas():
    # Base installments
    base_installments = sumifs(
        lanc_sys_valorparcelabase,
        lanc_categoria, categoria,
        lanc_sys_faturainicial, concat(literal("<="), sys_competencia_normalizada),
        lanc_sys_faturafinal, concat(literal(">="), sys_competencia_normalizada)
    )

    # Final adjustment
    final_adj = sumifs(
        lanc_sys_ajusteultimaparcela,
        lanc_categoria, categoria,
        lanc_sys_faturafinal, sys_competencia_normalizada
    )

    # Card refunds
    refunds = sumifs(
        lanc_sys_creditofatura,
        lanc_categoria, categoria,
        lanc_sys_competenciaefetiva, sys_competencia_normalizada
    )
    
    total = subtract(add(base_installments, final_adj), refunds)
    
    return if_func(_build_is_valid_budget(), total, literal(""))


def build_consumido_comprometido():
    return if_func(_build_is_valid_budget(), add(saidas_em_conta, cartao_parcelas), literal(""))


def build_disponivel():
    return if_func(_build_is_valid_budget(), subtract(orcamento, consumido_comprometido), literal(""))


def build_uso_pct():
    calc = if_func(and_func(equals(orcamento, literal(0)), equals(consumido_comprometido, literal(0))), literal(0),
        if_func(equals(orcamento, literal(0)), literal(""),
            divide(consumido_comprometido, orcamento)
        )
    )
    return if_func(_build_is_valid_budget(), calc, literal(""))


def build_situacao():
    logic = if_func(less_than(consumido_comprometido, literal(0)), literal("Saldo a favor"),
        if_func(equals(consumido_comprometido, literal(0)), literal("Sem movimento"),
            if_func(less_than(disponivel, literal(0)), literal("Acima do orçamento"),
                if_func(equals(disponivel, literal(0)), literal("No limite"),
                    literal("Dentro do orçamento")
                )
            )
        )
    )
    return if_func(_build_is_valid_budget(), logic, literal(""))
