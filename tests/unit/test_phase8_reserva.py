import pytest
from excel_saas.core.excel.formulas import Formula
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.templates.finance_personal.reserva_semantics import (
    _build_is_valid_reserva,
    build_status,
    build_reserva_alvo,
    build_falta,
    build_cobertura_atual,
    build_progresso,
    build_situacao,
    build_row_index
)
from excel_saas.templates.finance_personal.reserva import build_reserva_sheet

def test_singleton_row_index():
    expr = build_row_index()
    f = str(Formula(expr))
    assert 'ROW()-MIN(ROW(tblReserva[Custo essencial mensal]))+1' in f or 'ROW()-(MIN(ROW(tblReserva[Custo essencial mensal])))+1' in f

def test_validity_predicate():
    expr = _build_is_valid_reserva()
    f = str(Formula(expr))
    assert 'ROW()-MIN(ROW(tblReserva[Custo essencial mensal]))+1=1' in f or 'ROW()-(MIN(ROW(tblReserva[Custo essencial mensal])))+1=1' in f
    assert '[@[Custo essencial mensal]]>0' in f
    assert '[@[Meses desejados]]>=1' in f
    assert '[@[Meses desejados]]=INT([@[Meses desejados]])' in f
    assert '[@[Reserva atual]]>=0' in f
    assert '[@Status]' not in f

def test_status():
    expr = build_status()
    f = str(Formula(expr))
    assert '"Use apenas uma linha"' in f
    assert '"Informe o custo essencial"' in f
    assert '"Custo essencial inválido"' in f
    assert '"Informe os meses"' in f
    assert '"Meses inválidos"' in f
    assert '"Informe a reserva atual"' in f
    assert '"Reserva atual inválida"' in f
    assert '"OK"' in f

def test_reserva_alvo():
    expr = build_reserva_alvo()
    f = str(Formula(expr))
    assert '[@[Custo essencial mensal]]*[@[Meses desejados]]' in f

def test_falta():
    expr = build_falta()
    f = str(Formula(expr))
    assert 'IF([@[Reserva atual]]>=[@[Reserva alvo]],0,[@[Reserva alvo]]-[@[Reserva atual]])' in f
    assert 'MAX' not in f

def test_cobertura_atual():
    expr = build_cobertura_atual()
    f = str(Formula(expr))
    assert '[@[Reserva atual]]/[@[Custo essencial mensal]]' in f

def test_progresso():
    expr = build_progresso()
    f = str(Formula(expr))
    assert '[@[Reserva atual]]/[@[Reserva alvo]]' in f

def test_situacao():
    expr = build_situacao()
    f = str(Formula(expr))
    assert '"Não iniciada"' in f
    assert '"Completa"' in f
    assert '"Em formação"' in f
    assert 'TODAY()' not in f
    assert 'IF([@[Reserva atual]]=0,"Não iniciada",IF([@[Reserva atual]]>=[@[Reserva alvo]],"Completa","Em formação"))' in f

def test_domain_isolation():
    exprs = [
        _build_is_valid_reserva(),
        build_status(),
        build_reserva_alvo(),
        build_falta(),
        build_cobertura_atual(),
        build_progresso(),
        build_situacao()
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

def test_table_plan():
    req = GenerationRequest(template_id="finance_personal", year=2026, reserve_months=6)
    sheet = build_reserva_sheet(req)
    assert len(sheet.tables) == 1
    table = sheet.tables[0]
    
    assert len(table.columns) == 9
    
    headers = [col.header for col in table.columns]
    assert headers == [
        "Custo essencial mensal", "Meses desejados", "Reserva atual", "Reserva alvo",
        "Falta", "Cobertura atual", "Progresso %", "Status", "Situação"
    ]
    
    for i in range(3):
        assert table.columns[i].role == CellRole.INPUT
        
    for i in range(3, 9):
        assert table.columns[i].role == CellRole.FORMULA
        
    assert sheet.is_protected is False
    assert sheet.freeze_panes == "B5"

def test_seed():
    # Valid
    sheet1 = build_reserva_sheet(GenerationRequest(template_id="finance_personal", year=2026, reserve_months=9))
    assert sheet1.tables[0].data[0][1] == 9
    
    sheet2 = build_reserva_sheet(GenerationRequest(template_id="finance_personal", year=2026, reserve_months=1))
    assert sheet2.tables[0].data[0][1] == 1
    
    # Invalid -> fallback 6
    sheet3 = build_reserva_sheet(GenerationRequest(template_id="finance_personal", year=2026, reserve_months=0))
    assert sheet3.tables[0].data[0][1] == 6
    
    sheet4 = build_reserva_sheet(GenerationRequest(template_id="finance_personal", year=2026, reserve_months=-1))
    assert sheet4.tables[0].data[0][1] == 6
    
    sheet5 = build_reserva_sheet(GenerationRequest(template_id="finance_personal", year=2026, reserve_months=2.7))
    assert sheet5.tables[0].data[0][1] == 6
    
    sheet6 = build_reserva_sheet(GenerationRequest(template_id="finance_personal", year=2026, reserve_months=True))
    assert sheet6.tables[0].data[0][1] == 6
    
    sheet7 = build_reserva_sheet(GenerationRequest(template_id="finance_personal", year=2026, reserve_months="12"))
    assert sheet7.tables[0].data[0][1] == 6
    
    sheet8 = build_reserva_sheet(GenerationRequest(template_id="finance_personal", year=2026, reserve_months=None))
    assert sheet8.tables[0].data[0][1] == 6
