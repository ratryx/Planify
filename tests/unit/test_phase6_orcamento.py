import pytest
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.templates.finance_personal.orcamento_semantics import (
    build_sys_competencia_normalizada,
    _build_is_valid_budget,
    build_status,
    build_saidas_em_conta,
    build_cartao_parcelas,
    build_consumido_comprometido,
    build_disponivel,
    build_uso_pct,
    build_situacao
)
from excel_saas.templates.finance_personal.orcamento import build_orcamento_sheet

def test_competencia_normalization():
    expr = build_sys_competencia_normalizada()
    f = str(Formula(expr))
    assert 'DATE(YEAR(INT([@Competência])),MONTH(INT([@Competência])),1)' in f
    assert '[@Competência]>=1' in f

def test_valid_budget_predicate():
    expr = _build_is_valid_budget()
    f = str(Formula(expr))
    assert 'COUNTIFS(tblCategorias[Categoria],[@Categoria])=1' in f
    assert 'COUNTIFS(tblOrcamento[Categoria],[@Categoria],tblOrcamento[sys_CompetenciaNormalizada],[@[sys_CompetenciaNormalizada]])=1' in f
    assert '[@Orçamento]>=0' in f

def test_status_precedence():
    expr = build_status()
    f = str(Formula(expr))
    assert '"Informe a competência"' in f
    assert '"Competência inválida"' in f
    assert '"Informe a categoria"' in f
    assert '"Categoria não cadastrada"' in f
    assert '"Categoria duplicada"' in f
    assert '"Informe o orçamento"' in f
    assert '"Orçamento inválido"' in f
    assert '"Orçamento duplicado"' in f
    assert '"OK"' in f

def test_account_consumption():
    expr = build_saidas_em_conta()
    f = str(Formula(expr))
    # General expenses and refunds
    assert 'SUMIFS(tblLancamentos[sys_Despesa],tblLancamentos[Categoria],[@Categoria],tblLancamentos[Conta],"<>",tblLancamentos[Data],">="&[@[sys_CompetenciaNormalizada]],tblLancamentos[Data],"<"&EDATE([@[sys_CompetenciaNormalizada]],1))' in f
    # Investment
    assert 'SUMIFS(tblLancamentos[sys_CaixaConta],tblLancamentos[Tipo],"Investimento"' in f
    # Debt
    assert 'SUMIFS(tblLancamentos[sys_CaixaConta],tblLancamentos[Tipo],"Pagamento de dívida"' in f
    # Must negate Investimento and Pagamento de dívida because CaixaConta is negative for outflows
    assert '*-1' in f or '*-1' in f.replace(' ', '')
    # Check that is_valid_budget gate is used, and it shouldn't contain Status text gate
    assert 'COUNTIFS(tblCategorias[Categoria],[@Categoria])=1' in f # inside valid budget logic
    assert '[@Status]' not in f

def test_exclusions():
    expr = build_saidas_em_conta()
    f = str(Formula(expr))
    assert '"Receita"' not in f
    assert '"Resgate"' not in f
    assert '"Transferência"' not in f
    assert '"Pagamento de fatura"' not in f
    
def test_card_consumption():
    expr = build_cartao_parcelas()
    f = str(Formula(expr))
    # Base installments
    assert 'SUMIFS(tblLancamentos[sys_ValorParcelaBase],tblLancamentos[Categoria],[@Categoria],tblLancamentos[sys_FaturaInicial],"<="&[@[sys_CompetenciaNormalizada]],tblLancamentos[sys_FaturaFinal],">="&[@[sys_CompetenciaNormalizada]])' in f
    # Final adjustment
    assert 'SUMIFS(tblLancamentos[sys_AjusteUltimaParcela],tblLancamentos[Categoria],[@Categoria],tblLancamentos[sys_FaturaFinal],[@[sys_CompetenciaNormalizada]])' in f
    # Refunds
    assert 'SUMIFS(tblLancamentos[sys_CreditoFatura],tblLancamentos[Categoria],[@Categoria],tblLancamentos[sys_CompetenciaEfetiva],[@[sys_CompetenciaNormalizada]])' in f
    # Exclusions
    assert 'tblFaturas' not in f
    assert 'sys_ValorParcela]' not in f
    assert 'sys_Despesa]' not in f
    assert 'tblLancamentos[Valor]' not in f
    assert 'sys_PagamentoFatura]' not in f
    # Valid budget gate
    assert '[@Status]' not in f

def test_uso_pct_zero_budget():
    expr = build_uso_pct()
    f = str(Formula(expr))
    assert 'IF(AND([@Orçamento]=0,[@[Consumido / Comprometido]]=0),0,IF([@Orçamento]=0,"",[@[Consumido / Comprometido]]/[@Orçamento]))' in f

def test_situation_strings():
    expr = build_situacao()
    f = str(Formula(expr))
    assert '"Saldo a favor"' in f
    assert '"Sem movimento"' in f
    assert '"Acima do orçamento"' in f
    assert '"No limite"' in f
    assert '"Dentro do orçamento"' in f
    assert 'TODAY()' not in f
    assert '[@Status]' not in f

def test_table_plan():
    sheet = build_orcamento_sheet()
    assert len(sheet.tables) == 1
    table = sheet.tables[0]
    
    assert len(table.columns) == 11
    
    headers = [col.header for col in table.columns]
    assert headers == [
        "Competência", "Categoria", "Orçamento", "Saídas em conta", "Cartão / Parcelas",
        "Consumido / Comprometido", "Disponível", "Uso %", "Status", "Situação", "sys_CompetenciaNormalizada"
    ]
    
    assert table.columns[10].role == CellRole.SYSTEM
    assert table.columns[10].hidden is True
    assert sheet.is_protected is False
