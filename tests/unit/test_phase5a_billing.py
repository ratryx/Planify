import pytest
from excel_saas.templates.finance_personal.lancamentos_semantics import (
    build_status_fatura, build_sys_competencia_efetiva, build_sys_caixa_conta, build_sys_cartao
)
from excel_saas.templates.finance_personal.cards import (
    build_status_formula as build_cards_status, build_sys_dia_fechamento_seguro
)

def test_phase4_regression_safety():
    """
    Pagamento de fatura WITHOUT Competência da fatura must STILL output correct cash/card semantics
    and those helpers MUST NOT reference Competência da fatura or Status fatura.
    """
    caixa_conta_formula = str(build_sys_caixa_conta())
    sys_cartao_formula = str(build_sys_cartao())

    assert "Pagamento de fatura" in caixa_conta_formula
    assert "Pagamento de fatura" in sys_cartao_formula

    assert "Competência da fatura" not in caixa_conta_formula
    assert "Status fatura" not in caixa_conta_formula
    assert "sys_CompetenciaEfetiva" not in caixa_conta_formula

    assert "Competência da fatura" not in sys_cartao_formula
    assert "Status fatura" not in sys_cartao_formula
    assert "sys_CompetenciaEfetiva" not in sys_cartao_formula

def test_card_status_and_closing_day():
    card_status = str(build_cards_status())

    assert "Informe o nome" in card_status
    assert "Nome duplicado" in card_status
    assert "Limite inválido" in card_status
    assert "Fechamento inválido" in card_status
    assert "Vencimento inválido" in card_status
    assert "COUNTIFS(tblCartoes[Nome],[@Nome])" in card_status

    safe_closing = str(build_sys_dia_fechamento_seguro())
    assert "COUNTIFS(tblCartoes[Nome],[@Nome])=1" in safe_closing
    assert "ISNUMBER([@[Dia fechamento]])" in safe_closing
    assert "INT([@[Dia fechamento]])" in safe_closing

    # Verify boundaries (our AST produces >0 and <32 which is equivalent to >=1 and <=31)
    assert ">0" in safe_closing
    assert "<32" in safe_closing

    # Must NOT reference other fields
    assert "Limite" not in safe_closing
    assert "Dia vencimento" not in safe_closing
    assert "Ativo?" not in safe_closing
    assert "Status" not in safe_closing

def test_excel_date_safety():
    status_fatura = str(build_status_fatura())
    comp_efetiva = str(build_sys_competencia_efetiva())

    for formula_str in [status_fatura, comp_efetiva]:
        assert "ISNUMBER(" in formula_str
        assert ">=1" in formula_str
        assert "<DATE(9999,12,31)+1" in formula_str

def test_calendar_normalization_and_nominal_competence():
    status_fatura = str(build_status_fatura())
    comp_efetiva = str(build_sys_competencia_efetiva())

    # Check overrides
    assert "Competência inválida" in status_fatura
    assert "Informe a competência" in status_fatura
    assert "Cartão não cadastrado" in status_fatura
    assert "Cartão duplicado" in status_fatura

    # Calendar normalization
    assert "INT([@Data])" in comp_efetiva
    assert "INT([@[Competência da fatura]])" in comp_efetiva

    # Override date extraction
    assert "DATE(YEAR(INT([@[Competência da fatura]])),MONTH(INT([@[Competência da fatura]])),1)" in comp_efetiva

    # Nominal competence calculation
    assert "MIN(SUMIFS(tblCartoes[sys_DiaFechamentoSeguro],tblCartoes[Nome],[@Cartão]),DAY(EOMONTH(INT([@Data]),0)))" in comp_efetiva
    assert "EDATE(" in comp_efetiva


def test_ast_grouping():
    from excel_saas.core.excel.formulas import multiply, group, subtract, literal, Formula
    # multiply(literal(10), group(subtract(literal(3), literal(1)))) -> 10*(3-1)
    expr = multiply(literal(10), group(subtract(literal(3), literal(1))))
    assert str(expr) == '10*(3-1)'

    # Also test sys_ValorUltimaParcela string
    from excel_saas.templates.finance_personal.lancamentos_semantics import build_sys_valor_ultima_parcela
    f_str = str(build_sys_valor_ultima_parcela())
    # Should not serialize as base*n-1
    assert "*([@sys_ParcelasEfetivas]-1)" in f_str

def test_effective_installments():
    from excel_saas.templates.finance_personal.lancamentos_semantics import build_sys_parcelas_efetivas
    f_str = str(build_sys_parcelas_efetivas())
    # blank -> 1, populated -> parcelas
    assert 'IF(ISBLANK([@Parcelas]),1,[@Parcelas])' in f_str
    # valid card purchase + invalid base transaction -> 0
    assert ',0)' in f_str
    assert '=[@Tipo]="Despesa"' not in f_str
    assert 'NOT(ISBLANK([@Cartão]))' in f_str

def test_exact_rounding_helpers():
    from excel_saas.templates.finance_personal.lancamentos_semantics import (
        build_sys_valor_parcela_base,
        build_sys_valor_ultima_parcela,
        build_sys_ajuste_ultima_parcela
    )
    base_str = str(build_sys_valor_parcela_base())
    assert 'ROUND([@Valor]/[@sys_ParcelasEfetivas],2)' in base_str

    last_str = str(build_sys_valor_ultima_parcela())
    # Valor - Base * (N - 1)
    assert '[@Valor]-[@sys_ValorParcelaBase]*([@sys_ParcelasEfetivas]-1)' in last_str

    adj_str = str(build_sys_ajuste_ultima_parcela())
    # Ultima - Base
    assert '[@sys_ValorUltimaParcela]-[@sys_ValorParcelaBase]' in adj_str

def test_schedule_range_safety():
    from excel_saas.templates.finance_personal.lancamentos_semantics import build_sys_fatura_inicial, build_status_fatura
    f_str = str(build_sys_fatura_inicial())

    # start month index uses YEAR() and MONTH()
    assert 'YEAR(IF(NOT([@sys_CompetenciaEfetiva]=""),[@sys_CompetenciaEfetiva],1))*12+MONTH(IF(NOT([@sys_CompetenciaEfetiva]=""),[@sys_CompetenciaEfetiva],1))' in f_str
    # maximum month 120000
    assert '<=120000' in f_str

    status_str = str(build_status_fatura())
    assert '"Parcelamento fora do intervalo"' in status_str
    assert '>120000' in status_str

def test_year_crossing():
    from excel_saas.templates.finance_personal.lancamentos_semantics import build_sys_fatura_final
    f_str = str(build_sys_fatura_final())
    assert 'EDATE([@sys_FaturaInicial],[@sys_ParcelasEfetivas]-1)' in f_str

def test_credits_and_payments():
    from excel_saas.templates.finance_personal.lancamentos_semantics import (
        build_sys_credito_fatura, build_sys_pagamento_fatura,
        build_sys_caixa_conta, build_sys_cartao
    )
    credit_str = str(build_sys_credito_fatura())
    assert 'NOT([@sys_CompetenciaEfetiva]="")' in credit_str
    assert '[@Tipo]="Estorno / Reembolso"' in credit_str
    assert ',[@Valor],0)' in credit_str

    payment_str = str(build_sys_pagamento_fatura())
    assert '[@Tipo]="Pagamento de fatura"' in payment_str
    assert ',[@Valor],0)' in payment_str

    # Phase 4 regression:
    caixa_str = str(build_sys_caixa_conta())
    cartao_str = str(build_sys_cartao())
    for f in [caixa_str, cartao_str]:
        assert 'sys_PagamentoFatura' not in f
        assert 'sys_FaturaInicial' not in f
        assert 'sys_FaturaFinal' not in f
        assert 'Status fatura' not in f
