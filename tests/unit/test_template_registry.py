import pytest
from excel_saas.core.registry.template_registry import TemplateRegistry
from excel_saas.templates.base import BaseTemplate

class DummyTemplate(BaseTemplate):
    def build_workbook_plan(self, request):
        return None

def test_registry_registration():
    registry = TemplateRegistry()
    registry.register("dummy", DummyTemplate)

    template = registry.get("dummy")
    assert isinstance(template, DummyTemplate)

def test_registry_duplicate_registration():
    registry = TemplateRegistry()
    registry.register("dummy", DummyTemplate)

    with pytest.raises(ValueError):
        registry.register("dummy", DummyTemplate)

def test_registry_unknown_template():
    registry = TemplateRegistry()
    with pytest.raises(ValueError):
        registry.get("unknown")
