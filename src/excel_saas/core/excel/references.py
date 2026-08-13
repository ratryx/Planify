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
            # Excel uses brackets for structured references.
            # Example: tblLancamentos[Valor]
            # If the column name has special characters like # or ', they might need escaping in some cases,
            # but usually just being inside brackets is enough for standard names.
            # If we wanted to be strictly safe for all possible column names (e.g. ones with brackets in them)
            # we would need more escaping, but for now we assume column names are sanitized.
            return f"{self.table_name}[{self.column}]"
        return self.table_name

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
