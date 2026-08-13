from enum import Enum, auto

class CellRole(Enum):
    """Semantic role of a cell, used to determine styling and protection."""
    INPUT = auto()      # Editable by end user
    FORMULA = auto()    # Generated calculation; protected
    SYSTEM = auto()     # Technical calculation; protected, possibly hidden
    HEADER = auto()     # Table or section header
    TITLE = auto()      # Main sheet title
    NORMAL = auto()     # Standard read-only text or label
