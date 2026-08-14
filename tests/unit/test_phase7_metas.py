import pytest
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.templates.finance_personal.metas_semantics import (
    _build_is_valid_goal,
    build_status,
    build_falta,
    build_progresso,
    build_meses_restantes,
    build_aporte_mensal,
    build_situacao
)
from excel_saas.templates.finance_personal.metas import build_metas_sheet

def test_validity_predicate():
    expr = _build_is_valid_goal()
    f = str(Formula(expr))
    assert 'COUNTIFS(tblMetas[Meta],[@Meta])=1' in f
    assert '[@[Valor alvo]]>0' in f
    assert '[@[Valor atual]]>=0' in f
    assert '[@Status]' not in f

def test_status():
    expr = build_status()
    f = str(Formula(expr))
    assert '"Informe a meta"' in f
    assert '"Meta duplicada"' in f
    assert '"Informe o valor alvo"' in f
    assert '"Valor alvo inválido"' in f
    assert '"Informe o valor atual"' in f
    assert '"Valor atual inválido"' in f
    assert '"Informe a data alvo"' in f
    assert '"Data alvo inválida"' in f
    assert '"OK"' in f

def test_falta():
    expr = build_falta()
    f = str(Formula(expr))
    assert 'IF([@[Valor atual]]>=[@[Valor alvo]],0,[@[Valor alvo]]-[@[Valor atual]])' in f
    # Must use validity predicate
    assert 'COUNTIFS(tblMetas[Meta],[@Meta])=1' in f
    assert '[@Status]' not in f

def test_progresso():
    expr = build_progresso()
    f = str(Formula(expr))
    assert '[@[Valor atual]]/[@[Valor alvo]]' in f
    assert 'MIN(' not in f
    assert 'MAX(' not in f

def test_meses_restantes():
    expr = build_meses_restantes()
    f = str(Formula(expr))
    
    # Must contain grouped month extraction
    assert '(YEAR(INT([@[Data alvo]]))*12+MONTH(INT([@[Data alvo]])))-(YEAR(TODAY())*12+MONTH(TODAY()))+1' in f or \
           '(YEAR(INT([@[Data alvo]]))*12+MONTH(INT([@[Data alvo]])))-(YEAR(TODAY())*12+MONTH(TODAY())))+1' in f or \
           '((YEAR(INT([@[Data alvo]]))*12+MONTH(INT([@[Data alvo]])))-(YEAR(TODAY())*12+MONTH(TODAY())))+1' in f
    
    assert 'DATEDIF' not in f
    assert 'IF(YEAR(INT([@[Data alvo]]))*12+MONTH(INT([@[Data alvo]]))<YEAR(TODAY())*12+MONTH(TODAY()),0' in f

def test_aporte_mensal():
    expr = build_aporte_mensal()
    f = str(Formula(expr))
    assert 'IF([@Falta]=0,0,IF([@[Meses restantes]]=0,"",[@Falta]/[@[Meses restantes]]))' in f

def test_situacao():
    expr = build_situacao()
    f = str(Formula(expr))
    assert '"Concluída"' in f
    assert '"Prazo vencido"' in f
    assert '"Em andamento"' in f
    
    # Check that concluded is evaluated before overdue
    assert 'IF([@[Valor atual]]>=[@[Valor alvo]],"Concluída"' in f

def test_domain_isolation():
    exprs = [
        _build_is_valid_goal(),
        build_status(),
        build_falta(),
        build_progresso(),
        build_meses_restantes(),
        build_aporte_mensal(),
        build_situacao()
    ]
    for expr in exprs:
        f = str(Formula(expr))
        assert 'tblLancamentos' not in f
        assert 'tblContas' not in f
        assert 'tblFaturas' not in f
        assert 'tblParcelamentos' not in f
        assert 'tblOrcamento' not in f

def test_table_plan():
    sheet = build_metas_sheet()
    assert len(sheet.tables) == 1
    table = sheet.tables[0]
    
    assert len(table.columns) == 10
    
    headers = [col.header for col in table.columns]
    assert headers == [
        "Meta", "Valor alvo", "Valor atual", "Data alvo", "Falta",
        "Progresso %", "Meses restantes", "Aporte mensal necessário", "Status", "Situação"
    ]
    
    for i in range(4):
        assert table.columns[i].role == CellRole.INPUT
        
    for i in range(4, 10):
        assert table.columns[i].role == CellRole.FORMULA
        
    assert sheet.is_protected is False
    assert sheet.freeze_panes == "B5"
