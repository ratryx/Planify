from abc import ABC, abstractmethod

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

class StringRef(Reference):
    """Raw string reference (e.g., 'A1:B10')."""
    def __init__(self, ref: str):
        self.ref = ref
        
    def __str__(self) -> str:
        return self.ref
