from excel_saas.core.excel.formulas import (
    edate, date_func, year_func, today_func, month_func, add, sumifs, subtract,
    if_func, and_func, equals, literal, divide, concat, Expression
)
from excel_saas.core.excel.references import ThisRowRef, TableRef

def _effective_projection_horizon(value: int) -> int:
    return max(1, min(value, 60))

def build_competencia() -> Expression:
    offset = ThisRowRef("sys_Offset")
    base_date = date_func(year_func(today_func()), month_func(today_func()), literal(1))
    return edate(base_date, add(offset, literal(1)))

def build_orcamento_planejado() -> Expression:
    return sumifs(
        TableRef("tblOrcamento", "Orçamento"),
        TableRef("tblOrcamento", "sys_CompetenciaNormalizada"),
        ThisRowRef("Competência"),
        TableRef("tblOrcamento", "Status"),
        literal("OK")
    )

def build_card_commitments() -> Expression:
    competencia = ThisRowRef("Competência")
    
    base = sumifs(
        TableRef("tblLancamentos", "sys_ValorParcelaBase"),
        TableRef("tblLancamentos", "sys_FaturaInicial"),
        concat(literal("<="), competencia),
        TableRef("tblLancamentos", "sys_FaturaFinal"),
        concat(literal(">="), competencia)
    )
    
    final_adj = sumifs(
        TableRef("tblLancamentos", "sys_AjusteUltimaParcela"),
        TableRef("tblLancamentos", "sys_FaturaFinal"),
        competencia
    )
    
    credit = sumifs(
        TableRef("tblLancamentos", "sys_CreditoFatura"),
        TableRef("tblLancamentos", "sys_CompetenciaEfetiva"),
        competencia
    )
    
    return subtract(add(base, final_adj), credit)

def build_structural_debt_commitments() -> Expression:
    competencia = ThisRowRef("Competência")
    
    open_debt = sumifs(
        TableRef("tblDividas", "sys_ParcelaMensalValida"),
        TableRef("tblDividas", "Data final"),
        concat(literal(">="), competencia)
    )
    
    blank_debt = sumifs(
        TableRef("tblDividas", "sys_ParcelaMensalValida"),
        TableRef("tblDividas", "Data final"),
        literal("")
    )
    
    return add(open_debt, blank_debt)

def build_known_commitments() -> Expression:
    return add(
        ThisRowRef("Compromissos no cartão"),
        ThisRowRef("Dívidas estruturais")
    )

def build_budget_margin() -> Expression:
    return subtract(
        ThisRowRef("Orçamento planejado"),
        ThisRowRef("Compromissos conhecidos")
    )

def build_budget_usage() -> Expression:
    orcamento = ThisRowRef("Orçamento planejado")
    commitments = ThisRowRef("Compromissos conhecidos")
    
    return if_func(
        and_func(equals(orcamento, literal(0)), equals(commitments, literal(0))),
        literal(0),
        if_func(
            equals(orcamento, literal(0)),
            literal(""),
            divide(commitments, orcamento)
        )
    )
