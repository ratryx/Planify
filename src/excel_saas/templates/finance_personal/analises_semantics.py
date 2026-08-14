from excel_saas.core.excel.formulas import (
    sumifs, literal, sum_func, if_func, and_func, subtract, divide, negate, Expression
)
from excel_saas.core.excel.references import TableRef
from excel_saas.templates.finance_personal.dashboard_patrimonio_semantics import build_total_assets

def build_negative_account_balances() -> Expression:
    return negate(
        sumifs(
            TableRef("tblContas", "Saldo atual"),
            TableRef("tblContas", "Status"), literal("OK"),
            TableRef("tblContas", "Ativa?"), literal("Sim"),
            TableRef("tblContas", "Saldo atual"), literal("<0")
        )
    )

def build_structural_debt_balance() -> Expression:
    return sum_func(TableRef("tblDividas", "sys_SaldoDevedorValido"))

def build_registered_card_position() -> Expression:
    return sum_func(TableRef("tblLancamentos", "sys_Cartao"))

def build_registered_position() -> Expression:
    return subtract(
        subtract(
            subtract(
                build_total_assets(),
                build_negative_account_balances()
            ),
            build_structural_debt_balance()
        ),
        build_registered_card_position()
    )

def build_available_balance() -> Expression:
    return sumifs(
        TableRef("tblContas", "Saldo atual"),
        TableRef("tblContas", "Incluir no saldo disponível?"), literal("Sim"),
        TableRef("tblContas", "Ativa?"), literal("Sim"),
        TableRef("tblContas", "Status"), literal("OK")
    )

def build_reserve_coverage() -> Expression:
    return sum_func(TableRef("tblReserva", "Cobertura atual"))

def build_reserve_gap() -> Expression:
    return sum_func(TableRef("tblReserva", "Falta"))

def build_goal_monthly_contribution() -> Expression:
    return sum_func(TableRef("tblMetas", "Aporte mensal necessário"))

def build_horizon_commitments() -> Expression:
    return sum_func(TableRef("tblProjecoes", "Compromissos conhecidos"))

def build_horizon_budget() -> Expression:
    return sum_func(TableRef("tblProjecoes", "Orçamento planejado"))

def build_horizon_margin() -> Expression:
    return subtract(
        build_horizon_budget(),
        build_horizon_commitments()
    )

def build_horizon_usage() -> Expression:
    orcamento = build_horizon_budget()
    compromissos = build_horizon_commitments()
    return if_func(
        and_func(
            equals_expr(orcamento, literal(0)),
            equals_expr(compromissos, literal(0))
        ),
        literal(0),
        if_func(
            equals_expr(orcamento, literal(0)),
            literal(""),
            divide(compromissos, orcamento)
        )
    )

def equals_expr(left, right) -> Expression:
    from excel_saas.core.excel.formulas import equals
    return equals(left, right)
