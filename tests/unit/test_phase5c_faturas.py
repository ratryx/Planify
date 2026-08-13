import pytest
from excel_saas.core.excel.formulas import _val, literal, func

from excel_saas.core.excel.references import ThisRowRef
from excel_saas.templates.finance_personal.faturas import (
    build_sys_competencia_normalizada,
    _build_is_valid_fatura,
    build_status_formula,
    build_fechamento,
    build_vencimento,
    build_compras,
    build_creditos,
    build_total,
    build_pagamentos,
    build_em_aberto,
    build_situacao
)
from excel_saas.templates.finance_personal.cards import build_sys_dia_vencimento_seguro

def test_concat_ast():
    from excel_saas.core.excel.formulas import concat, literal
    expr = concat(literal("<="), literal("hello"))
    assert _val(expr) == '"<="&"hello"'

def test_sys_competencia_normalizada():
    expr = build_sys_competencia_normalizada()
    formula = _val(expr)
    
    assert "ISNUMBER([@Competência])" in formula
    assert "INT([@Competência])" in formula
    assert "1" in formula
    assert "2958465" in formula
    assert "DATE(YEAR(INT([@Competência])),MONTH(INT([@Competência])),1)" in formula

def test_sys_dia_vencimento_seguro():
    expr = build_sys_dia_vencimento_seguro()
    formula = _val(expr)
    
    assert "COUNTIFS(tblCartoes[Nome],[@Nome])=1" in formula
    assert "ISNUMBER([@[Dia vencimento]])" in formula
    assert "[@[Dia vencimento]]=INT([@[Dia vencimento]])" in formula
    assert "[@[Dia vencimento]]>0" in formula
    assert "[@[Dia vencimento]]<32" in formula
    assert "Limite" not in formula
    assert "Dia fechamento" not in formula
    assert "Ativo?" not in formula
    assert "Status" not in formula
    assert "Conta de pagamento" not in formula

def test_status_formula():
    expr = build_status_formula()
    formula = _val(expr)
    
    assert '"Informe o cartão"' in formula
    assert '"Cartão não cadastrado"' in formula
    assert '"Cartão duplicado"' in formula
    assert '"Informe a competência"' in formula
    assert '"Competência inválida"' in formula
    assert '"Fatura duplicada"' in formula
    assert '"OK"' in formula
    
    assert "tblFaturas[Cartão]" in formula
    assert "tblFaturas[sys_CompetenciaNormalizada]" in formula

def test_financial_validity_independence():
    compras = _val(build_compras())
    assert "[@Status]" not in compras
    assert "[@Situação]" not in compras
    
    creditos = _val(build_creditos())
    assert "[@Status]" not in creditos
    
    pagamentos = _val(build_pagamentos())
    assert "[@Status]" not in pagamentos

def test_fechamento_formula():
    expr = build_fechamento()
    formula = _val(expr)
    
    assert "MIN(" in formula
    assert "DAY(EOMONTH([@sys_CompetenciaNormalizada],0))" in formula
    assert "DATE(YEAR([@sys_CompetenciaNormalizada]),MONTH([@sys_CompetenciaNormalizada])" in formula

def test_vencimento_formula():
    expr = build_vencimento()
    formula = _val(expr)
    
    # Check configured due > configured closing -> same month vs next month
    assert ">" in formula
    
    # Check EDATE protection against > 120000
    assert "YEAR([@sys_CompetenciaNormalizada])*12+MONTH([@sys_CompetenciaNormalizada])+1<=120000" in formula

def test_compras_aggregation():
    expr = build_compras()
    formula = _val(expr)
    
    assert "tblLancamentos[sys_ValorParcelaBase]" in formula
    assert "tblLancamentos[sys_AjusteUltimaParcela]" in formula
    assert "tblLancamentos[sys_FaturaInicial]" in formula
    assert "tblLancamentos[sys_FaturaFinal]" in formula
    assert "tblLancamentos[Cartão]" in formula
    
    # Concat formatting
    assert '"<="&[@sys_CompetenciaNormalizada]' in formula
    assert '">="&[@sys_CompetenciaNormalizada]' in formula
    
    assert "tblLancamentos[sys_ValorParcela]" not in formula

def test_creditos_pagamentos_aggregation():
    creditos = _val(build_creditos())
    assert "tblLancamentos[sys_CreditoFatura]" in creditos
    assert "tblLancamentos[sys_CompetenciaEfetiva]" in creditos
    assert "[@sys_CompetenciaNormalizada]" in creditos
    
    pagamentos = _val(build_pagamentos())
    assert "tblLancamentos[sys_PagamentoFatura]" in pagamentos
    assert "tblLancamentos[sys_CompetenciaEfetiva]" in pagamentos

def test_total_and_em_aberto():
    total = _val(build_total())
    assert "[@[Compras / Parcelas]]-[@[Créditos / Estornos]]" in total
    
    em_aberto = _val(build_em_aberto())
    assert "[@[Total da fatura]]-[@Pagamentos]" in em_aberto

def test_situacao():
    expr = build_situacao()
    formula = _val(expr)
    
    assert '"Sem movimento"' in formula
    assert '"Crédito"' in formula
    assert '"Paga"' in formula
    assert '"Em aberto — sem vencimento"' in formula
    assert '"Vencida"' in formula
    assert '"Em aberto"' in formula
    
    # TODAY() > Vencimento, not >=
    assert "TODAY()>[@Vencimento]" in formula
