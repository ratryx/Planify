from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any, Union
from .cell_roles import CellRole

@dataclass
class DataValidationPlan:
    validate: str  # 'list', 'integer', 'decimal', etc.
    source: Union[List[str], str, Any]  # can be a list of options, formula string, or Formula object
    input_title: str = ""
    input_message: str = ""
    error_title: str = "Valor Inválido"
    error_message: str = "Por favor, selecione um valor válido da lista."

@dataclass
class ColumnPlan:
    header: str
    width: Optional[float] = None
    role: CellRole = CellRole.INPUT
    formula: Optional[str] = None  # if this column is entirely calculated
    validation: Optional[DataValidationPlan] = None
    number_format: Optional[str] = None

@dataclass
class TablePlan:
    name: str
    start_cell: str  # e.g., 'B4'. The engine calculates the full range.
    columns: List[ColumnPlan]
    data: List[List[Any]] = field(default_factory=list)  # Initial rows of data
    style: str = "Table Style Light 1"
    show_total_row: bool = False

@dataclass
class CellPlan:
    row: int  # 0-indexed
    col: int  # 0-indexed
    value: Any = None
    formula: Optional[str] = None
    role: CellRole = CellRole.NORMAL
    number_format: Optional[str] = None
    bold: bool = False
    size: Optional[int] = None
    validation: Optional[DataValidationPlan] = None

@dataclass
class WorksheetPlan:
    name: str
    is_protected: bool = True
    freeze_panes: Optional[str] = None  # e.g., 'A2' or 'B6'
    show_gridlines: bool = False
    tab_color: Optional[str] = None
    
    # Content
    cells: List[CellPlan] = field(default_factory=list)
    tables: List[TablePlan] = field(default_factory=list)
    
    # Columns config (0-indexed col index to width)
    column_widths: Dict[int, float] = field(default_factory=dict)

@dataclass
class WorkbookPlan:
    worksheets: List[WorksheetPlan] = field(default_factory=list)
    # metadata, defined names, etc., can be added here
