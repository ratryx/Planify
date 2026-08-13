from abc import ABC, abstractmethod
from typing import Optional
from .naming import escape_sheet_name

class Reference(ABC):
    @abstractmethod
    def __str__(self) -> str:
        pass

class TableRef(Reference):
    """Reference to an Excel Structured Table or one of its columns."""
    def __init__(self, table_name: str, column: str = None):
        self.table_name = table_name
        self.column = column

    def __str__(self) -> str:
        if self.column:
            return f"{self.table_name}[{self.column}]"
        return self.table_name

class ThisRowRef(Reference):
    """Reference to the current row of a table column (e.g. [@Tipo] or [@[Conta destino]])."""
    def __init__(self, column: str):
        self.column = column

    def __str__(self) -> str:
        # If the column has spaces or special characters, Excel requires [@[Column Name]]
        # For safety, we can just always wrap the inner part in brackets if it has non-alphanumeric (excluding underscore)
        import re
        if not re.match(r'^[\w]+$', self.column):
            return f"[@[{self.column}]]"
        return f"[@{self.column}]"

class DefinedNameRef(Reference):
    """Reference to a workbook Defined Name."""
    def __init__(self, name: str):
        self.name = name

    def __str__(self) -> str:
        return self.name

class StringRef(Reference):
    """Raw string reference (e.g., 'A1:B10')."""
    def __init__(self, ref: str):
        self.ref = ref

    def __str__(self) -> str:
        return self.ref

class CellRef(Reference):
    """A reference to a specific cell, optionally with a sheet name."""
    def __init__(self, cell: str, sheet: Optional[str] = None):
        self.cell = cell
        self.sheet = sheet

    def __str__(self) -> str:
        if self.sheet:
            return f"{escape_sheet_name(self.sheet)}!{self.cell}"
        return self.cell
