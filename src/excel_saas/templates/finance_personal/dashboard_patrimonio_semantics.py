from excel_saas.core.excel.formulas import (
    sumifs, literal, sum_func, if_func, equals, divide
)
from excel_saas.core.excel.references import TableRef, ThisRowRef

def build_account_assets():
    return sumifs(
        TableRef("tblContas", "Saldo atual"),
        TableRef("tblContas", "Status"), literal("OK"),
        TableRef("tblContas", "Ativa?"), literal("Sim"),
        TableRef("tblContas", "Saldo atual"), literal(">0")
    )

def build_investment_assets():
    return sum_func(TableRef("tblInvestimentos", "sys_ValorAtualValido"))

def build_additional_assets():
    return sum_func(TableRef("tblBensPatrimoniais", "sys_ValorAtualValido"))

def build_total_assets():
    return sum_func(
        build_account_assets(),
        build_investment_assets(),
        build_additional_assets()
    )

def build_component_value():
    componente = ThisRowRef("Componente")
    return if_func(
        equals(componente, literal("Contas e caixa")),
        build_account_assets(),
        if_func(
            equals(componente, literal("Investimentos")),
            build_investment_assets(),
            if_func(
                equals(componente, literal("Bens patrimoniais")),
                build_additional_assets(),
                literal(0)
            )
        )
    )

def build_component_weight():
    total_assets = build_total_assets()
    valor_atual = ThisRowRef("Valor atual")
    return if_func(
        equals(total_assets, literal(0)),
        literal(0),
        divide(valor_atual, total_assets)
    )
