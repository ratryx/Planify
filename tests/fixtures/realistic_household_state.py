from datetime import date
from typing import Any, List
from excel_saas.core.models.workbook_plan import WorkbookPlan

def get_table(plan: WorkbookPlan, name: str):
    for ws in plan.worksheets:
        for tbl in ws.tables:
            if tbl.name == name:
                return tbl
    raise ValueError(f"Table {name} not found")

def add_row(tbl, **kwargs):
    row = [None] * len(tbl.columns)
    for k, v in kwargs.items():
        # Find column index by header
        idx = -1
        for i, col in enumerate(tbl.columns):
            if col.header == k:
                idx = i
                break
        if idx == -1:
            raise ValueError(f"Column '{k}' not found in table '{tbl.name}'")
        row[idx] = v
    tbl.data.append(row)

def inject_realistic_household(plan: WorkbookPlan, year: int):
    # tblContas
    # multiple positive active accounts;
    # one account excluded from available balance;
    # one negative active account.
    tbl_contas = get_table(plan, "tblContas")
    tbl_contas.data.clear()
    add_row(tbl_contas, Nome="Conta Corrente", Tipo="Conta Corrente", Instituição="Banco A", **{"Saldo inicial": 5000, "Incluir no saldo disponível?": "Sim", "Ativa?": "Sim"})
    add_row(tbl_contas, Nome="Poupança", Tipo="Conta Poupança", Instituição="Banco B", **{"Saldo inicial": 15000, "Incluir no saldo disponível?": "Não", "Ativa?": "Sim"})
    add_row(tbl_contas, Nome="Cheque Especial", Tipo="Conta Corrente", Instituição="Banco C", **{"Saldo inicial": -1000, "Incluir no saldo disponível?": "Sim", "Ativa?": "Sim"})

    # tblCartoes
    # at least two cards;
    # one fully configured;
    # one card demonstrating that card name alone is valid.
    tbl_cartoes = get_table(plan, "tblCartoes")
    tbl_cartoes.data.clear()
    add_row(tbl_cartoes, Nome="Cartão Principal", Limite=10000, **{"Dia fechamento": 15, "Dia vencimento": 25, "Conta de pagamento": "Conta Corrente", "Ativo?": "Sim"})
    add_row(tbl_cartoes, Nome="Cartão Secundário") # just name

    # tblCategorias (usually prepopulated but we might need to add if needed, let's assume default categories exist, but wait we might need to match them)
    # the template doesn't prepopulate tblCategorias with default in `data`, wait, defaults are applied in `with_sample_data=True`.
    # Let's populate tblCategorias to be safe.
    tbl_categorias = get_table(plan, "tblCategorias")
    tbl_categorias.data.clear()
    add_row(tbl_categorias, Categoria="Alimentação")
    add_row(tbl_categorias, Categoria="Moradia")
    add_row(tbl_categorias, Categoria="Lazer")
    add_row(tbl_categorias, Categoria="Salário")
    add_row(tbl_categorias, Categoria="Transporte")
    add_row(tbl_categorias, Categoria="Saúde")
    add_row(tbl_categorias, Categoria="Educação")

    # tblLancamentos
    # income; cash expense; card expense; installment card expense; transfer; investment; redemption; card invoice payment; refund; essential expense; dated activity.
    tbl_lancamentos = get_table(plan, "tblLancamentos")
    tbl_lancamentos.data.clear()
    # Income
    add_row(tbl_lancamentos, Data=date(year, 1, 5), Descrição="Salário", Tipo="Receita", Categoria="Salário", Conta="Conta Corrente", Valor=8000, **{"Essencial?": "Não", "Recorrente?": "Sim"})
    # Cash expense (essential)
    add_row(tbl_lancamentos, Data=date(year, 1, 10), Descrição="Aluguel", Tipo="Despesa", Categoria="Moradia", Conta="Conta Corrente", Valor=-2000, **{"Essencial?": "Sim", "Recorrente?": "Sim"})
    # Card expense
    add_row(tbl_lancamentos, Data=date(year, 1, 12), Descrição="Mercado", Tipo="Despesa", Categoria="Alimentação", Cartão="Cartão Principal", **{"Competência da fatura": date(year, 1, 1), "Valor": -800, "Essencial?": "Sim", "Recorrente?": "Não"})
    # Installment card expense
    add_row(tbl_lancamentos, Data=date(year, 1, 20), Descrição="TV Nova", Tipo="Despesa", Categoria="Lazer", Cartão="Cartão Principal", **{"Competência da fatura": date(year, 1, 1), "Valor": -3000, "Parcelas": 10, "Essencial?": "Não", "Recorrente?": "Não"})
    # Transfer
    add_row(tbl_lancamentos, Data=date(year, 1, 15), Descrição="Transferência para Poupança", Tipo="Transferência", Conta="Conta Corrente", **{"Conta destino": "Poupança", "Valor": 1000})
    # Investment
    add_row(tbl_lancamentos, Data=date(year, 1, 25), Descrição="Compra de Ações", Tipo="Investimento", Conta="Conta Corrente", Valor=500)
    # Redemption
    add_row(tbl_lancamentos, Data=date(year, 2, 5), Descrição="Resgate CDB", Tipo="Resgate", Conta="Poupança", Valor=200)
    # Card invoice payment
    add_row(tbl_lancamentos, Data=date(year, 2, 25), Descrição="Pagamento Fatura", Tipo="Pagamento de Fatura", Conta="Conta Corrente", Cartão="Cartão Principal", **{"Competência da fatura": date(year, 1, 1), "Valor": 3800})
    # Refund
    add_row(tbl_lancamentos, Data=date(year, 1, 22), Descrição="Estorno TV", Tipo="Receita", Categoria="Lazer", Cartão="Cartão Principal", **{"Competência da fatura": date(year, 1, 1), "Valor": 3000})

    # tblOrcamento
    # valid monthly budget rows.
    tbl_orcamento = get_table(plan, "tblOrcamento")
    tbl_orcamento.data.clear()
    add_row(tbl_orcamento, Competência=date(year, 1, 1), Categoria="Moradia", Orçamento=2500)
    add_row(tbl_orcamento, Competência=date(year, 1, 1), Categoria="Alimentação", Orçamento=1500)
    add_row(tbl_orcamento, Competência=date(year, 1, 1), Categoria="Lazer", Orçamento=500)

    # tblMetas
    # at least one valid active goal.
    tbl_metas = get_table(plan, "tblMetas")
    tbl_metas.data.clear()
    add_row(tbl_metas, Meta="Viagem Férias", **{"Valor alvo": 10000, "Valor atual": 2000, "Data alvo": date(year+1, 12, 1)})

    # tblReserva
    # valid reserve configuration.
    tbl_reserva = get_table(plan, "tblReserva")
    tbl_reserva.data.clear()
    add_row(tbl_reserva, **{"Custo essencial mensal": 4000, "Meses desejados": 6, "Reserva atual": 15000})

    # tblInvestimentos
    # multiple valid investment classes
    tbl_investimentos = get_table(plan, "tblInvestimentos")
    tbl_investimentos.data.clear()
    add_row(tbl_investimentos, Ativo="CDB Banco A", Classe="Renda Fixa", Instituição="Banco A", **{"Total aportado": 5000, "Total recebido": 0, "Valor atual": 5200})
    add_row(tbl_investimentos, Ativo="ITUB4", Classe="Ações", Instituição="Corretora B", **{"Total aportado": 2000, "Total recebido": 100, "Valor atual": 2300})

    # tblBensPatrimoniais
    # at least one valid asset
    tbl_bens = get_table(plan, "tblBensPatrimoniais")
    tbl_bens.data.clear()
    add_row(tbl_bens, Bem="Carro", Categoria="Veículos", **{"Valor atual": 50000})

    # tblDividas
    # one debt with final date; one debt with blank final date; positive monthly payments.
    tbl_dividas = get_table(plan, "tblDividas")
    tbl_dividas.data.clear()
    add_row(tbl_dividas, Dívida="Financiamento Carro", Categoria="Veículos", Credor="Banco A", **{"Saldo devedor atual": 25000, "Parcela mensal atual": 800, "Data final": date(year+2, 5, 1)})
    add_row(tbl_dividas, Dívida="Empréstimo Família", Categoria="Pessoal", Credor="Tio", **{"Saldo devedor atual": 5000, "Parcela mensal atual": 200})
