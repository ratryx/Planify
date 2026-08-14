from excel_saas.core.excel.formulas import (
    sum_func, sumifs, if_func, equals, literal, add, subtract, divide, group
)
from excel_saas.core.excel.references import TableRef, ThisRowRef

PORTFOLIO_APORTE = sum_func(TableRef("tblInvestimentos", "sys_AporteValido"))
PORTFOLIO_RECEBIDO = sum_func(TableRef("tblInvestimentos", "sys_RecebidoValido"))
PORTFOLIO_ATUAL = sum_func(TableRef("tblInvestimentos", "sys_ValorAtualValido"))

def build_portfolio_aportado():
    return PORTFOLIO_APORTE

def build_portfolio_recebido():
    return PORTFOLIO_RECEBIDO

def build_portfolio_atual():
    return PORTFOLIO_ATUAL

def build_portfolio_resultado():
    return subtract(group(add(PORTFOLIO_ATUAL, PORTFOLIO_RECEBIDO)), PORTFOLIO_APORTE)

def build_portfolio_retorno():
    return if_func(
        equals(PORTFOLIO_APORTE, literal(0)),
        literal(0),
        divide(build_portfolio_resultado(), PORTFOLIO_APORTE)
    )

def build_class_aportado():
    return sumifs(TableRef("tblInvestimentos", "sys_AporteValido"), TableRef("tblInvestimentos", "Classe"), ThisRowRef("Classe"))

def build_class_recebido():
    return sumifs(TableRef("tblInvestimentos", "sys_RecebidoValido"), TableRef("tblInvestimentos", "Classe"), ThisRowRef("Classe"))

def build_class_atual():
    return sumifs(TableRef("tblInvestimentos", "sys_ValorAtualValido"), TableRef("tblInvestimentos", "Classe"), ThisRowRef("Classe"))

def build_class_resultado():
    atual = ThisRowRef("Valor atual")
    recebido = ThisRowRef("Total recebido")
    aportado = ThisRowRef("Total aportado")
    return subtract(group(add(atual, recebido)), aportado)

def build_class_peso():
    return if_func(
        equals(PORTFOLIO_ATUAL, literal(0)),
        literal(0),
        divide(ThisRowRef("Valor atual"), PORTFOLIO_ATUAL)
    )

def build_class_retorno():
    aportado = ThisRowRef("Total aportado")
    return if_func(
        equals(aportado, literal(0)),
        literal(0),
        divide(ThisRowRef("Resultado total"), aportado)
    )
