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
sys_valor_parcela = ThisRowRef("sys_ValorParcela")

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
