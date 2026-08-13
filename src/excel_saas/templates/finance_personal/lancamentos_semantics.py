from excel_saas.core.excel.formulas import (
    if_func, equals, isblank, or_func, and_func, not_func, literal,
    isnumber, divide, subtract, negate, add, Expression, not_equals,
    less_than, less_or_equal, int_func
)
from excel_saas.core.excel.references import ThisRowRef

tipo = ThisRowRef("Tipo")
conta = ThisRowRef("Conta")
conta_dest = ThisRowRef("Conta destino")
cartao = ThisRowRef("Cartão")
valor = ThisRowRef("Valor")
parcelas = ThisRowRef("Parcelas")
data = ThisRowRef("Data")
comp_fatura = ThisRowRef("Competência da fatura")
sys_valor_parcela = ThisRowRef("sys_ValorParcela")
comp_efetiva = ThisRowRef("sys_CompetenciaEfetiva")
parcelas_efetivas = ThisRowRef("sys_ParcelasEfetivas")
fatura_inicial = ThisRowRef("sys_FaturaInicial")
fatura_final = ThisRowRef("sys_FaturaFinal")
valor_parcela_base = ThisRowRef("sys_ValorParcelaBase")
valor_ultima_parcela = ThisRowRef("sys_ValorUltimaParcela")

def build_status_formula() -> Expression:
    is_empty_row = and_func(
        isblank(tipo),
        isblank(conta),
        isblank(conta_dest),
        isblank(cartao),
        isblank(valor),
        isblank(parcelas)
    )

    check_valor = if_func(isblank(valor), literal("Informe o valor"),
        if_func(not_func(isnumber(valor)), literal("Valor inválido"),
            if_func(less_or_equal(valor, literal(0)), literal("Valor deve ser maior que zero"),
                literal("OK"))))

    check_parcelas = if_func(not_func(isblank(parcelas)),
        if_func(not_func(isnumber(parcelas)), literal("Parcelas inválidas"),
            if_func(less_than(parcelas, literal(1)), literal("Parcelas inválidas"),
                if_func(not_equals(parcelas, int_func(parcelas)), literal("Parcelas inválidas"),
                    literal("OK")))),
        literal("OK"))

    check_receita = if_func(isblank(conta), literal("Informe a conta"),
        if_func(not_func(isblank(conta_dest)), literal("Conta destino deve ficar em branco"),
            if_func(not_func(isblank(cartao)), literal("Cartão deve ficar em branco"),
                if_func(not_func(isblank(parcelas)), literal("Parcelas não permitidas"), literal("OK")))))

    check_despesa = if_func(and_func(isblank(conta), isblank(cartao)), literal("Informe conta ou cartão"),
        if_func(and_func(not_func(isblank(conta)), not_func(isblank(cartao))), literal("Use conta ou cartão, não ambos"),
            if_func(not_func(isblank(conta_dest)), literal("Conta destino deve ficar em branco"),
                if_func(and_func(not_func(isblank(conta)), not_func(isblank(parcelas))), literal("Parcelas só no cartão"), literal("OK")))))

    check_transferencia = if_func(isblank(conta), literal("Informe a conta origem"),
        if_func(isblank(conta_dest), literal("Informe a conta destino"),
            if_func(equals(conta, conta_dest), literal("Contas devem ser diferentes"),
                if_func(not_func(isblank(cartao)), literal("Cartão deve ficar em branco"),
                    if_func(not_func(isblank(parcelas)), literal("Parcelas não permitidas"), literal("OK"))))))

    check_investimento = if_func(isblank(conta), literal("Informe a conta"),
        if_func(not_func(isblank(conta_dest)), literal("Conta destino deve ficar em branco"),
            if_func(not_func(isblank(cartao)), literal("Cartão deve ficar em branco"),
                if_func(not_func(isblank(parcelas)), literal("Parcelas não permitidas"), literal("OK")))))

    check_pagamento_fatura = if_func(isblank(conta), literal("Informe a conta"),
        if_func(isblank(cartao), literal("Informe o cartão"),
            if_func(not_func(isblank(conta_dest)), literal("Conta destino deve ficar em branco"),
                if_func(not_func(isblank(parcelas)), literal("Parcelas não permitidas"), literal("OK")))))

    check_estorno = if_func(and_func(isblank(conta), isblank(cartao)), literal("Informe conta ou cartão"),
        if_func(and_func(not_func(isblank(conta)), not_func(isblank(cartao))), literal("Use conta ou cartão, não ambos"),
            if_func(not_func(isblank(conta_dest)), literal("Conta destino deve ficar em branco"),
                if_func(not_func(isblank(parcelas)), literal("Parcelas não permitidas"), literal("OK")))))

    tipo_logic = if_func(equals(tipo, literal("Receita")), check_receita,
        if_func(equals(tipo, literal("Despesa")), check_despesa,
            if_func(equals(tipo, literal("Transferência")), check_transferencia,
                if_func(or_func(equals(tipo, literal("Investimento")), equals(tipo, literal("Resgate")), equals(tipo, literal("Pagamento de dívida"))), check_investimento,
                    if_func(equals(tipo, literal("Pagamento de fatura")), check_pagamento_fatura,
                        if_func(equals(tipo, literal("Estorno / Reembolso")), check_estorno,
                            literal("Tipo inválido")))))))

    main_logic = if_func(not_equals(check_valor, literal("OK")), check_valor,
        if_func(not_equals(check_parcelas, literal("OK")), check_parcelas,
            tipo_logic))

    return if_func(is_empty_row, literal(""), main_logic)

def build_is_valid_transaction() -> Expression:
    return equals(build_status_formula(), literal("OK"))

def build_sys_receita() -> Expression:
    return if_func(not_func(build_is_valid_transaction()), literal(0),
        if_func(equals(tipo, literal("Receita")), valor, literal(0)))

def build_sys_despesa() -> Expression:
    return if_func(not_func(build_is_valid_transaction()), literal(0),
        if_func(equals(tipo, literal("Despesa")), valor,
            if_func(equals(tipo, literal("Estorno / Reembolso")), negate(valor), literal(0))))

def build_sys_valor_parcela() -> Expression:
    safe_parcelas = if_func(or_func(isblank(parcelas), less_than(parcelas, literal(1)), not_func(isnumber(parcelas))), literal(1), parcelas)
    return if_func(not_func(build_is_valid_transaction()), literal(0),
        if_func(and_func(equals(tipo, literal("Despesa")), not_func(isblank(cartao))),
            if_func(or_func(isblank(parcelas), equals(parcelas, literal(1))), valor,
                divide(valor, safe_parcelas)), literal(0)))

def build_sys_compromisso_futuro() -> Expression:
    return if_func(not_func(build_is_valid_transaction()), literal(0),
        if_func(and_func(equals(tipo, literal("Despesa")), not_func(isblank(cartao))),
            if_func(or_func(isblank(parcelas), equals(parcelas, literal(1))), literal(0),
                subtract(valor, sys_valor_parcela)), literal(0)))

def build_sys_caixa_conta() -> Expression:
    return if_func(not_func(build_is_valid_transaction()), literal(0),
        if_func(and_func(or_func(equals(tipo, literal("Receita")), equals(tipo, literal("Resgate")), equals(tipo, literal("Estorno / Reembolso"))), not_func(isblank(conta))), valor,
            if_func(and_func(or_func(equals(tipo, literal("Despesa")), equals(tipo, literal("Transferência")), equals(tipo, literal("Investimento")), equals(tipo, literal("Pagamento de fatura")), equals(tipo, literal("Pagamento de dívida"))), not_func(isblank(conta))), negate(valor), literal(0))))

def build_sys_caixa_destino() -> Expression:
    return if_func(not_func(build_is_valid_transaction()), literal(0),
        if_func(equals(tipo, literal("Transferência")), valor, literal(0)))

def build_sys_cartao() -> Expression:
    return if_func(not_func(build_is_valid_transaction()), literal(0),
        if_func(and_func(equals(tipo, literal("Despesa")), not_func(isblank(cartao))), valor,
            if_func(and_func(or_func(equals(tipo, literal("Pagamento de fatura")), equals(tipo, literal("Estorno / Reembolso"))), not_func(isblank(cartao))), negate(valor), literal(0))))

def is_valid_date(val) -> Expression:
    from excel_saas.core.excel.formulas import and_func, not_func, isblank, isnumber, greater_or_equal, less_than, literal, date_func, add
    return and_func(
        not_func(isblank(val)),
        isnumber(val),
        greater_or_equal(val, literal(1)),
        less_than(val, add(date_func(literal(9999), literal(12), literal(31)), literal(1)))
    )

def build_status_fatura() -> Expression:
    from excel_saas.core.excel.formulas import (
        countifs, greater_than, sumifs, multiply, year_func, month_func, add, subtract
    )
    from excel_saas.core.excel.references import TableRef

    tx_invalid = not_func(build_is_valid_transaction())

    card_count = countifs(TableRef("tblCartoes", "Nome"), cartao)
    card_unregistered = equals(card_count, literal(0))
    card_duplicate = greater_than(card_count, literal(1))

    check_card = if_func(card_unregistered, literal("Cartão não cadastrado"),
        if_func(card_duplicate, literal("Cartão duplicado"), literal("OK")))

    comp_populated = not_func(isblank(comp_fatura))
    comp_valid = is_valid_date(comp_fatura)
    comp_invalid = and_func(comp_populated, not_func(comp_valid))

    check_pagamento = if_func(tx_invalid, literal("Lançamento inválido"),
        if_func(not_equals(check_card, literal("OK")), check_card,
            if_func(isblank(comp_fatura), literal("Informe a competência"),
                if_func(comp_invalid, literal("Competência inválida"), literal("OK")))))

    data_populated = not_func(isblank(data))
    data_valid = is_valid_date(data)
    data_invalid = and_func(data_populated, not_func(data_valid))

    safe_closing_day = sumifs(TableRef("tblCartoes", "sys_DiaFechamentoSeguro"), TableRef("tblCartoes", "Nome"), cartao)
    has_closing = greater_than(safe_closing_day, literal(0))

    check_card_tx = if_func(tx_invalid, literal("Lançamento inválido"),
        if_func(not_equals(check_card, literal("OK")), check_card,
            if_func(comp_invalid, literal("Competência inválida"),
                if_func(comp_valid, literal("OK"),
                    if_func(isblank(data), literal("Informe a data"),
                        if_func(data_invalid, literal("Data inválida"),
                            if_func(not_func(has_closing), literal("Sem fechamento"), literal("OK"))))))))

    valid_comp = not_func(equals(comp_efetiva, literal("")))
    safe_comp = if_func(valid_comp, comp_efetiva, literal(1))
    start_month = add(multiply(year_func(safe_comp), literal(12)), month_func(safe_comp))
    end_month = subtract(add(start_month, parcelas_efetivas), literal(1))

    schedule_overflow = and_func(
        equals(tipo, literal("Despesa")),
        valid_comp,
        greater_than(parcelas_efetivas, literal(0)),
        greater_than(end_month, literal(120000))
    )

    final_card_tx = if_func(equals(check_card_tx, literal("OK")),
        if_func(schedule_overflow, literal("Parcelamento fora do intervalo"), literal("OK")),
        check_card_tx
    )

    is_card_tx = and_func(not_func(isblank(cartao)), or_func(equals(tipo, literal("Despesa")), equals(tipo, literal("Estorno / Reembolso"))))

    return if_func(equals(tipo, literal("Pagamento de fatura")), check_pagamento,
        if_func(is_card_tx, final_card_tx, literal("")))

def build_sys_competencia_efetiva() -> Expression:
    from excel_saas.core.excel.formulas import (
        date_func, year_func, month_func, day_func, edate, eomonth, min_func, sumifs, greater_than
    )
    from excel_saas.core.excel.references import TableRef

    comp_populated = not_func(isblank(comp_fatura))
    comp_valid = is_valid_date(comp_fatura)

    override_date = int_func(comp_fatura)
    normalized_override = date_func(year_func(override_date), month_func(override_date), literal(1))

    pagamento_logic = if_func(comp_valid, normalized_override, literal(""))

    safe_closing_day = sumifs(TableRef("tblCartoes", "sys_DiaFechamentoSeguro"), TableRef("tblCartoes", "Nome"), cartao)
    has_closing = greater_than(safe_closing_day, literal(0))

    data_valid = is_valid_date(data)
    transaction_date = int_func(data)

    effective_closing_day = min_func(safe_closing_day, day_func(eomonth(transaction_date, literal(0))))
    effective_closing_date = date_func(year_func(transaction_date), month_func(transaction_date), effective_closing_day)

    nominal_competence = if_func(less_or_equal(transaction_date, effective_closing_date),
        date_func(year_func(transaction_date), month_func(transaction_date), literal(1)),
        edate(date_func(year_func(transaction_date), month_func(transaction_date), literal(1)), literal(1)))

    card_tx_logic = if_func(comp_valid, normalized_override,
        if_func(comp_populated, literal(""),
            if_func(and_func(has_closing, data_valid), nominal_competence, literal(""))))

    is_card_tx = and_func(not_func(isblank(cartao)), or_func(equals(tipo, literal("Despesa")), equals(tipo, literal("Estorno / Reembolso"))))

    return if_func(equals(tipo, literal("Pagamento de fatura")), pagamento_logic,
        if_func(is_card_tx, card_tx_logic, literal("")))

def build_sys_parcelas_efetivas() -> Expression:
    is_card_despesa = and_func(
        build_is_valid_transaction(),
        equals(tipo, literal("Despesa")),
        not_func(isblank(cartao))
    )
    return if_func(is_card_despesa,
        if_func(isblank(parcelas), literal(1), parcelas),
        literal(0)
    )

def _build_is_allocatable() -> Expression:
    from excel_saas.core.excel.formulas import multiply, year_func, month_func, greater_than, less_or_equal
    is_card_despesa = and_func(
        build_is_valid_transaction(),
        equals(tipo, literal("Despesa")),
        not_func(isblank(cartao))
    )
    valid_comp = not_func(equals(comp_efetiva, literal("")))
    has_installments = greater_than(parcelas_efetivas, literal(0))

    safe_comp = if_func(valid_comp, comp_efetiva, literal(1))
    start_month = add(
        multiply(year_func(safe_comp), literal(12)),
        month_func(safe_comp)
    )
    end_month = subtract(
        add(start_month, parcelas_efetivas),
        literal(1)
    )
    is_in_range = less_or_equal(end_month, literal(120000))

    return and_func(
        is_card_despesa,
        valid_comp,
        has_installments,
        is_in_range
    )

def build_sys_fatura_inicial() -> Expression:
    return if_func(_build_is_allocatable(), comp_efetiva, literal(""))

def build_sys_fatura_final() -> Expression:
    from excel_saas.core.excel.formulas import edate
    has_fatura_inicial = not_func(equals(fatura_inicial, literal("")))
    return if_func(has_fatura_inicial,
        edate(fatura_inicial, subtract(parcelas_efetivas, literal(1))),
        literal("")
    )

def build_sys_valor_parcela_base() -> Expression:
    from excel_saas.core.excel.formulas import round_func
    has_fatura_inicial = not_func(equals(fatura_inicial, literal("")))
    return if_func(has_fatura_inicial,
        round_func(divide(valor, parcelas_efetivas), literal(2)),
        literal(0)
    )

def build_sys_valor_ultima_parcela() -> Expression:
    from excel_saas.core.excel.formulas import multiply, group
    has_fatura_inicial = not_func(equals(fatura_inicial, literal("")))
    return if_func(has_fatura_inicial,
        subtract(valor, multiply(valor_parcela_base, group(subtract(parcelas_efetivas, literal(1))))),
        literal(0)
    )

def build_sys_ajuste_ultima_parcela() -> Expression:
    has_fatura_inicial = not_func(equals(fatura_inicial, literal("")))
    return if_func(has_fatura_inicial,
        subtract(valor_ultima_parcela, valor_parcela_base),
        literal(0)
    )

def build_sys_credito_fatura() -> Expression:
    is_card_refund = and_func(
        build_is_valid_transaction(),
        equals(tipo, literal("Estorno / Reembolso")),
        not_func(isblank(cartao)),
        not_func(equals(comp_efetiva, literal("")))
    )
    return if_func(is_card_refund, valor, literal(0))

def build_sys_pagamento_fatura() -> Expression:
    is_card_payment = and_func(
        build_is_valid_transaction(),
        equals(tipo, literal("Pagamento de fatura")),
        not_func(isblank(cartao)),
        not_func(equals(comp_efetiva, literal("")))
    )
    return if_func(is_card_payment, valor, literal(0))
