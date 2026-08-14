import pytest
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.templates.finance_personal.investimentos_semantics import (
    CLASSES_PERMITIDAS,
    _build_is_valid_classe,
    _build_is_duplicate,
    _build_is_valid_investment,
    build_status,
    build_resultado_total,
    build_retorno_simples,
    build_peso_carteira,
    build_situacao,
    build_sys_aporte_valido,
    build_sys_recebido_valido,
    build_sys_valor_atual_valido
)
from excel_saas.templates.finance_personal.investimentos import build_investimentos_sheet

def test_classes_permitidas():
    assert CLASSES_PERMITIDAS == ["Renda fixa", "Ações", "FIIs", "ETFs", "Fundos", "Cripto", "Previdência", "Outros"]
    assert "Exterior" not in CLASSES_PERMITIDAS
    
    expr = _build_is_valid_classe()
    f = str(Formula(expr))
    assert 'OR(' in f
    for c in CLASSES_PERMITIDAS:
        assert f'[@Classe]="{c}"' in f

def test_duplicate_identity():
    expr = _build_is_duplicate()
    f = str(Formula(expr))
    assert 'COUNTIFS(tblInvestimentos[Ativo],[@Ativo],tblInvestimentos[Instituição],[@Instituição])>1' in f

def test_validity_predicate():
    expr = _build_is_valid_investment()
    f = str(Formula(expr))
    assert '[@[Total aportado]]>0' in f
    assert '[@[Total recebido]]>=0' in f
    assert '[@[Valor atual]]>=0' in f
    assert '[@Status]' not in f

def test_status():
    expr = build_status()
    f = str(Formula(expr))
    assert '"Informe o ativo"' in f
    assert '"Informe a classe"' in f
    assert '"Classe inválida"' in f
    assert '"Investimento duplicado"' in f
    assert '"Informe o total aportado"' in f
    assert '"Total aportado inválido"' in f
    assert '"Informe o total recebido"' in f
    assert '"Total recebido inválido"' in f
    assert '"Informe o valor atual"' in f
    assert '"Valor atual inválido"' in f
    assert '"OK"' in f

def test_resultado_total():
    expr = build_resultado_total()
    f = str(Formula(expr))
    assert '([@[Valor atual]]+[@[Total recebido]])-[@[Total aportado]]' in f

def test_retorno_simples():
    expr = build_retorno_simples()
    f = str(Formula(expr))
    assert '[@[Resultado total]]/[@[Total aportado]]' in f
    assert 'MIN' not in f
    assert 'MAX' not in f

def test_helpers():
    f1 = str(Formula(build_sys_aporte_valido()))
    assert 'IF(' in f1 and ',[@[Total aportado]],0)' in f1
    
    f2 = str(Formula(build_sys_recebido_valido()))
    assert 'IF(' in f2 and ',[@[Total recebido]],0)' in f2
    
    f3 = str(Formula(build_sys_valor_atual_valido()))
    assert 'IF(' in f3 and ',[@[Valor atual]],0)' in f3

def test_peso_carteira():
    expr = build_peso_carteira()
    f = str(Formula(expr))
    assert 'SUM(tblInvestimentos[sys_ValorAtualValido])' in f
    assert 'IF(SUM(tblInvestimentos[sys_ValorAtualValido])=0,0,[@[Valor atual]]/SUM(tblInvestimentos[sys_ValorAtualValido]))' in f
    assert 'tblInvestimentos[Valor atual]' not in f

def test_situacao():
    expr = build_situacao()
    f = str(Formula(expr))
    assert '"Ganho"' in f
    assert '"Perda"' in f
    assert '"No zero"' in f

def test_domain_isolation():
    exprs = [
        _build_is_valid_investment(),
        build_status(),
        build_resultado_total(),
        build_retorno_simples(),
        build_peso_carteira(),
        build_situacao(),
        build_sys_aporte_valido(),
        build_sys_recebido_valido(),
        build_sys_valor_atual_valido()
    ]
    for expr in exprs:
        f = str(Formula(expr))
        assert 'tblLancamentos' not in f
        assert 'tblContas' not in f
        assert 'tblCartoes' not in f
        assert 'tblFaturas' not in f
        assert 'tblParcelamentos' not in f
        assert 'tblOrcamento' not in f
        assert 'tblMetas' not in f
        assert 'tblReserva' not in f

def test_table_plan():
    sheet = build_investimentos_sheet()
    assert len(sheet.tables) == 1
    table = sheet.tables[0]
    
    assert len(table.columns) == 14
    
    headers = [col.header for col in table.columns]
    assert headers == [
        "Ativo", "Classe", "Instituição", "Total aportado", "Total recebido", "Valor atual",
        "Resultado total", "Retorno simples %", "Peso carteira %", "Status", "Situação",
        "sys_AporteValido", "sys_RecebidoValido", "sys_ValorAtualValido"
    ]
    
    for i in range(6):
        assert table.columns[i].role == CellRole.INPUT
        
    for i in range(6, 11):
        assert table.columns[i].role == CellRole.FORMULA
        
    for i in range(11, 14):
        assert table.columns[i].role == CellRole.SYSTEM
        
    assert sheet.is_protected is False
    assert sheet.freeze_panes == "B5"
