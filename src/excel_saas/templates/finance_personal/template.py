from excel_saas.templates.base import BaseTemplate
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.core.models.workbook_plan import WorkbookPlan
from excel_saas.core.registry.template_registry import registry
from .worksheets import build_comece_aqui, build_dashboard, build_lancamentos, build_configuracoes
from .accounts import build_contas
from .cards import build_cartoes

class FinancePersonalTemplate(BaseTemplate):
    def build_workbook_plan(self, request: GenerationRequest) -> WorkbookPlan:
        from excel_saas.core.models.workbook_plan import DefinedNamePlan
        from excel_saas.core.excel.references import TableRef

        plan = WorkbookPlan(
            worksheets=[
                build_comece_aqui(request),
                build_dashboard(request),
                build_lancamentos(request),
                build_contas(request),
                build_cartoes(request),
                build_configuracoes(request)
            ],
            defined_names=[
                DefinedNamePlan("lista_categorias", TableRef("tblCategorias", "Categoria")),
                DefinedNamePlan("lista_tipos_conta", TableRef("tblTiposConta", "Tipo")),
                DefinedNamePlan("lista_contas", TableRef("tblContas", "Nome")),
                DefinedNamePlan("lista_cartoes", TableRef("tblCartoes", "Nome"))
            ]
        )
        return plan

# Auto-register this template
registry.register("finance_personal", FinancePersonalTemplate)
