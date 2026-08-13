import pytest
from excel_saas.core.excel.formulas import Formula
from excel_saas.templates.finance_personal.parcelamentos_semantics import (
    build_sys_indice_lancamento,
    build_projected_column,
    build_parcelas_restantes,
    build_proxima_competencia,
    build_compromisso_restante,
    build_situacao
)
from excel_saas.templates.finance_personal.parcelamentos import build_parcelamentos_sheet
from excel_saas.core.models.cell_roles import CellRole

def test_sys_indice_lancamento():
    expr = build_sys_indice_lancamento()
    f = str(Formula(expr))
    assert 'AGGREGATE(15,6,' in f
    assert 'ROW(tblLancamentos[sys_ParcelasEfetivas])' in f
    assert 'MIN(ROW(tblLancamentos[sys_ParcelasEfetivas]))' in f
    assert 'tblLancamentos[sys_ParcelasEfetivas]>1' in f
    assert 'tblLancamentos[sys_FaturaInicial]<>""' in f
    assert 'tblLancamentos[sys_FaturaFinal]<>""' in f
    assert '[@sys_Ordem]' in f
    assert 'IFERROR(' in f
    assert 'tblLancamentos[Status]' not in f
    assert 'tblLancamentos[Status fatura]' not in f
    assert 'tblFaturas' not in f

def test_source_projection():
    columns = ["Descrição", "Cartão", "Valor", "sys_ParcelasEfetivas", "sys_ValorParcelaBase", 
               "sys_FaturaInicial", "sys_FaturaFinal", "sys_AjusteUltimaParcela"]
    for col in columns:
        expr = build_projected_column(col)
        f = str(Formula(expr))
        assert f'INDEX(tblLancamentos[{col}],[@sys_IndiceLancamento])' in f
        assert 'IF([@sys_IndiceLancamento]="","",' in f

def test_current_month():
    f1 = str(Formula(build_parcelas_restantes()))
    f2 = str(Formula(build_proxima_competencia()))
    f3 = str(Formula(build_situacao()))
    
    assert 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)' in f1
    assert 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)' in f2
    assert 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)' in f3

def test_remaining_month():
    f = str(Formula(build_parcelas_restantes()))
    assert 'YEAR(' in f
    assert 'MONTH(' in f
    assert '*12' in f
    assert 'DATEDIF' not in f
    assert '+1' in f

def test_commitment():
    f = str(Formula(build_compromisso_restante()))
    assert '[@[Parcelas restantes]]*[@[Valor da parcela]]' in f
    assert '+[@sys_AjusteUltimaParcela]' in f

def test_lifecycle():
    f = str(Formula(build_situacao()))
    assert '"A iniciar"' in f
    assert '"Em andamento"' in f
    assert '"Concluído"' in f
    assert 'Pagamentos' not in f
    assert 'Faturas' not in f
    assert 'Status fatura' not in f

def test_table_plan():
    sheet = build_parcelamentos_sheet()
    assert len(sheet.tables) == 1
    table = sheet.tables[0]
    
    assert len(table.data) == 100
    assert table.data[0][11] == 1
    assert table.data[-1][11] == 100
    
    assert len(table.columns) == 14
    for i, col in enumerate(table.columns):
        if col.role == CellRole.SYSTEM:
            assert col.hidden is True
        else:
            assert col.formula is not None
            
    assert sheet.is_protected is True
