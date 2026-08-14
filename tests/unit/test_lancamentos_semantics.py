import pytest
from excel_saas.templates.finance_personal.lancamentos_semantics import (
    build_status_formula, build_sys_receita, build_sys_despesa,
    build_sys_valor_parcela, build_sys_compromisso_futuro,
    build_sys_caixa_conta, build_sys_caixa_destino, build_sys_cartao,
    build_is_valid_transaction
)

def test_semantic_builders_rely_on_status_column():
    """Ensure that the financial logic relies on the literal Status column."""
    formulas = [
        build_sys_receita(),
        build_sys_despesa(),
        build_sys_valor_parcela(),
        build_sys_compromisso_futuro(),
        build_sys_caixa_conta(),
        build_sys_caixa_destino(),
        build_sys_cartao()
    ]
    
    forbidden_strings = [
        "Informe o valor",
        "Valor inválido",
        "Valor deve ser maior que zero",
        "Informe conta ou cartão",
        "Use conta ou cartão, não ambos"
    ]
    
    for expr in formulas:
        f_str = str(expr)
        # Should use [@Status] for constant-time validity check
        assert "[@Status]" in f_str
        
        # Should NOT embed the old AST validations
        for forbidden in forbidden_strings:
            assert forbidden not in f_str

def test_is_valid_transaction_structure():
    """Ensure build_is_valid_transaction uses direct Status reference."""
    expr = build_is_valid_transaction()
    f_str = str(expr)
    
    # It should check if the precalculated Status column equals "OK"
    assert f_str == '([@Status]="OK")' or '[@Status]="OK"' in f_str
    assert "ISBLANK([@Tipo])" not in f_str

def test_sys_valor_parcela_safety():
    """sys_ValorParcela should guard against invalid parcelas."""
    expr = build_sys_valor_parcela()
    f_str = str(expr)
    
    # It should check if parcelas is blank, less than 1, or not a number, 
    # to avoid division by zero or invalid division
    assert "ISBLANK([@Parcelas])" in f_str
    assert "<1" in f_str
    assert "ISNUMBER([@Parcelas])" in f_str
    
def test_status_formula_structure():
    """Check that build_status_formula has the requested validations."""
    expr = build_status_formula()
    f_str = str(expr)
    
    assert "Valor inválido" in f_str
    assert "Valor deve ser maior que zero" in f_str
    assert "Parcelas inválidas" in f_str
    assert "ISNUMBER([@Valor])" in f_str
    assert "ISNUMBER([@Parcelas])" in f_str
    # Empty row check should be at the root
    assert "ISBLANK([@[Conta destino]])" in f_str
    assert 'IF(AND(ISBLANK([@Tipo])' in f_str
