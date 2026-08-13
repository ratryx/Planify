from typing import Dict, Type
from excel_saas.templates.base import BaseTemplate

class TemplateRegistry:
    def __init__(self):
        self._templates: Dict[str, Type[BaseTemplate]] = {}

    def register(self, template_id: str, template_class: Type[BaseTemplate]) -> None:
        if template_id in self._templates:
            raise ValueError(f"Template '{template_id}' is already registered.")
        self._templates[template_id] = template_class

    def get(self, template_id: str) -> BaseTemplate:
        template_class = self._templates.get(template_id)
        if not template_class:
            raise ValueError(f"Unknown template ID: '{template_id}'.")
        return template_class()

# Global registry instance
registry = TemplateRegistry()
