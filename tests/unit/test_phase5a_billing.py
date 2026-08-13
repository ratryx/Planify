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

def test_competence_logic():
    status_fatura = str(build_status_fatura())
    comp_efetiva = str(build_sys_competencia_efetiva())
    
    # Check overrides
    assert "Competência inválida" in status_fatura
    assert "Informe a competência" in status_fatura
    assert "Cartão não cadastrado" in status_fatura
    assert "Cartão duplicado" in status_fatura
    
    assert "DATE(YEAR([@[Competência da fatura]]),MONTH([@[Competência da fatura]]),1)" in comp_efetiva
    assert "MIN(SUMIFS(tblCartoes[sys_DiaFechamentoSeguro],tblCartoes[Nome],[@Cartão]),DAY(EOMONTH([@Data],0)))" in comp_efetiva
    assert "EDATE(" in comp_efetiva
