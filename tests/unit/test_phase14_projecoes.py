import pytest
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.templates.finance_personal.projecoes_semantics import (
    _effective_projection_horizon,
    build_competencia,
    build_orcamento_planejado,
    build_card_commitments,
    build_structural_debt_commitments,
    build_known_commitments,
    build_budget_margin,
    build_budget_usage
)
from excel_saas.templates.finance_personal.projecoes import build_projecoes_sheet

def test_effective_horizon():
    assert _effective_projection_horizon(0) == 1
    assert _effective_projection_horizon(1) == 1
    assert _effective_projection_horizon(12) == 12
    assert _effective_projection_horizon(60) == 60
    assert _effective_projection_horizon(61) == 60

def test_competence_formula():
    f = str(Formula(build_competencia()))
    assert f == "=EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),[@sys_Offset]+1)"
    assert "request.year" not in f

def test_budget_formula():
    f = str(Formula(build_orcamento_planejado()))
    assert f == '=SUMIFS(tblOrcamento[Orçamento],tblOrcamento[sys_CompetenciaNormalizada],[@Competência],tblOrcamento[Status],"OK")'

def test_card_formula():
    f = str(Formula(build_card_commitments()))
    expected = (
        '=SUMIFS(tblLancamentos[sys_ValorParcelaBase],tblLancamentos[sys_FaturaInicial],"<="&[@Competência],tblLancamentos[sys_FaturaFinal],">="&[@Competência])'
        '+SUMIFS(tblLancamentos[sys_AjusteUltimaParcela],tblLancamentos[sys_FaturaFinal],[@Competência])'
        '-SUMIFS(tblLancamentos[sys_CreditoFatura],tblLancamentos[sys_CompetenciaEfetiva],[@Competência])'
    )
    assert f == expected
    assert 'tblFaturas' not in f
    assert 'tblParcelamentos' not in f
    assert 'sys_PagamentoFatura' not in f
    assert 'Recorrente?' not in f

def test_debt_formula():
    f = str(Formula(build_structural_debt_commitments()))
    expected = (
        '=SUMIFS(tblDividas[sys_ParcelaMensalValida],tblDividas[Data final],">="&[@Competência])'
        '+SUMIFS(tblDividas[sys_ParcelaMensalValida],tblDividas[Data final],"")'
    )
    assert f == expected
    assert 'Saldo devedor atual' not in f

def test_known_total():
    f = str(Formula(build_known_commitments()))
    assert f == '=[@[Compromissos no cartão]]+[@[Dívidas estruturais]]'
    assert 'Orçamento planejado' not in f

def test_margin():
    f = str(Formula(build_budget_margin()))
    assert f == '=[@[Orçamento planejado]]-[@[Compromissos conhecidos]]'

def test_usage():
    f = str(Formula(build_budget_usage()))
    assert f == '=IF(AND([@[Orçamento planejado]]=0,[@[Compromissos conhecidos]]=0),0,IF([@[Orçamento planejado]]=0,"",[@[Compromissos conhecidos]]/[@[Orçamento planejado]]))'

def test_table_model_default():
    req = GenerationRequest(template_id="finance_personal", year=2026)
    sheet = build_projecoes_sheet(req)
    
    assert sheet.is_protected is True
    assert sheet.show_gridlines is False
    assert sheet.freeze_panes == "B5"
    
    table = sheet.tables[0]
    assert table.name == "tblProjecoes"
    assert table.start_cell == "B4"
    assert len(table.columns) == 8
    assert len(table.data) == 12
    
    headers = [c.header for c in table.columns]
    assert headers == [
        "Competência", "Orçamento planejado", "Compromissos no cartão",
        "Dívidas estruturais", "Compromissos conhecidos", "Margem vs orçamento",
        "Uso conhecido %", "sys_Offset"
    ]
    
    for i in range(7):
        assert table.columns[i].role == CellRole.FORMULA
    assert table.columns[7].role == CellRole.SYSTEM
    assert table.columns[7].hidden is True
    
    for c in table.columns:
        assert c.role != CellRole.INPUT
        
    offsets = [row[7] for row in table.data]
    assert offsets == list(range(12))
    for offset in offsets:
        assert isinstance(offset, int)

def test_custom_horizons():
    def get_len(h):
        req = GenerationRequest(template_id="finance_personal", year=2026, projection_horizon=h)
        return len(build_projecoes_sheet(req).tables[0].data)
    
    assert get_len(1) == 1
    assert get_len(60) == 60
    assert get_len(0) == 1
    assert get_len(61) == 60

def test_domain_isolation():
    funcs = [
        build_competencia, build_orcamento_planejado, build_card_commitments,
        build_structural_debt_commitments, build_known_commitments,
        build_budget_margin, build_budget_usage
    ]
    
    for func in funcs:
        f = str(Formula(func()))
        assert "tblContas" not in f
        assert "tblCartoes" not in f
        assert "tblFaturas" not in f
        assert "tblParcelamentos" not in f
        assert "tblMetas" not in f
        assert "tblReserva" not in f
        assert "tblInvestimentos" not in f
        assert "tblResumoInvestimentos" not in f
        assert "tblBensPatrimoniais" not in f
        assert "tblResumoPatrimonio" not in f
        assert "Recorrente?" not in f
        assert "sys_PagamentoFatura" not in f
