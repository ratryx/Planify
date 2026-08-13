from abc import ABC, abstractmethod
from typing import Dict, Any

from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.core.models.workbook_plan import WorkbookPlan

class BaseTemplate(ABC):
    @abstractmethod
    def build_workbook_plan(self, request: GenerationRequest) -> WorkbookPlan:
        """
        Constructs the declarative WorkbookPlan based on the request.
        """
        pass
