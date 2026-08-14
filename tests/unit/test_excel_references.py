from excel_saas.core.excel.references import ThisRowRef
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.templates.finance_personal.template import FinancePersonalTemplate

def test_this_row_ref_serialization():
    # simple names remain compact
    assert str(ThisRowRef("Status")) == "[@Status]"
    assert str(ThisRowRef("Tipo")) == "[@Tipo]"

    # underscore names become fully bracketed
    assert str(ThisRowRef("sys_ValorParcela")) == "[@[sys_ValorParcela]]"
    assert str(ThisRowRef("sys_CompetenciaEfetiva")) == "[@[sys_CompetenciaEfetiva]]"

    # existing space-containing names remain correct
    assert str(ThisRowRef("Conta destino")) == "[@[Conta destino]]"
    assert str(ThisRowRef("Competência da fatura")) == "[@[Competência da fatura]]"

def test_workbook_sys_references():
    req = GenerationRequest(template_id="finance_personal", year=2026)
    plan = FinancePersonalTemplate().build_workbook_plan(req)
    
    target_checks = [
        '[@[sys_ValorParcela]]',
        '[@[sys_CompetenciaEfetiva]]',
        '[@[sys_ParcelasEfetivas]]',
        '[@[sys_FaturaInicial]]',
        '[@[sys_ValorParcelaBase]]',
        '[@[sys_ValorUltimaParcela]]'
    ]
    found = {c: False for c in target_checks}
    
    for ws in plan.worksheets:
        for tbl in ws.tables:
            for c in tbl.columns:
                if c.formula:
                    f_str = str(c.formula)
                    for tc in target_checks:
                        if tc in f_str:
                            found[tc] = True
                            
    for tc, was_found in found.items():
        assert was_found, f"{tc} not found in any formula"
