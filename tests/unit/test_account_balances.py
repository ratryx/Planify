import pytest
from excel_saas.templates.finance_personal.accounts import build_saldo_atual_formula, build_status_formula
from excel_saas.templates.finance_personal.worksheets import build_dashboard
from excel_saas.core.models.generation_request import GenerationRequest

def test_saldo_atual_formula_structure():
    """Verify Saldo atual computes safely avoiding #DIV/0! or #VALUE! by using SUM function for all terms."""
    expr = build_saldo_atual_formula()
    f_str = str(expr)
    
    # Needs to handle blank rows cleanly and enforce safety
    assert "AND(NOT(ISBLANK([@Nome])),OR(ISBLANK([@[Saldo inicial]]),ISNUMBER([@[Saldo inicial]])))" in f_str
    
    # Needs to use SUM with multiple arguments rather than a + b
    assert "SUM([@[Saldo inicial]]" in f_str
    assert "SUMIFS(tblLancamentos[sys_CaixaConta]" in f_str
    assert "SUMIFS(tblLancamentos[sys_CaixaDestino]" in f_str
    
def test_status_formula_structure():
    """Verify Status handles duplicate detection properly."""
    expr = build_status_formula()
    f_str = str(expr)
    
    # Needs to handle blank rows cleanly with a full empty check
    assert "AND(ISBLANK([@Nome]),ISBLANK([@Tipo]),ISBLANK([@Instituição]),ISBLANK([@[Saldo inicial]]),ISBLANK([@[Incluir no saldo disponível?]]),ISBLANK([@[Ativa?]]))" in f_str
    
    # Needs to flag incomplete rows
    assert "ISBLANK([@Nome])" in f_str
    assert "Informe o nome" in f_str

    # Needs to flag duplicates
    assert "COUNTIFS(tblContas[Nome],[@Nome])>1" in f_str
    assert "Nome duplicado" in f_str

    # Needs to flag invalid numeric balances
    assert "ISNUMBER([@[Saldo inicial]])" in f_str
    assert "Saldo inicial inválido" in f_str

def test_dashboard_saldo_disponivel_structure():
    """Verify Dashboard correctly aggregates liquidity based on Ativa, Include, and Status."""
    req = GenerationRequest(template_id="finance_personal", year=2026)
    ws = build_dashboard(req)
    
    # Find the cell for Saldo Disponível Hoje
    saldo_cell = None
    for cell in ws.cells:
        if cell.row == 4 and cell.col == 1: # B5 -> row=4, col=1
            saldo_cell = cell
            break
            
    assert saldo_cell is not None
    f_str = str(saldo_cell.formula)
    
    # Must sum Saldo atual
    assert "SUMIFS(tblContas[Saldo atual]" in f_str
    
    # Must enforce Ativa = Sim
    assert "tblContas[Ativa?],\"Sim\"" in f_str
    
    # Must enforce Include = Sim
    assert "tblContas[Incluir no saldo disponível?],\"Sim\"" in f_str
    
    # Must enforce Status = OK
    assert "tblContas[Status],\"OK\"" in f_str
