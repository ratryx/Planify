from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any, Union
from .cell_roles import CellRole
from excel_saas.core.excel.formulas import Formula, Expression

@dataclass
class DataValidationPlan:
    validate: str  # 'list', 'integer', 'decimal', etc.
    source: Union[List[str], str, Any] = None  # Optional now, since numeric uses criteria/min/max
    criteria: Optional[str] = None  # e.g., '>=', 'between'
    minimum: Optional[Any] = None
    maximum: Optional[Any] = None
    ignore_blank: bool = True
    input_title: str = ""
    input_message: str = ""
    error_title: str = "Valor Inválido"
    error_message: str = "Por favor, insira um valor válido."

@dataclass
class ColumnPlan:
    header: str
    width: Optional[float] = None
    role: CellRole = CellRole.INPUT
    formula: Optional[Union[str, Formula, Expression]] = None  # if this column is entirely calculated
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
    formula: Optional[Union[str, Formula, Expression]] = None
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
class DefinedNamePlan:
    name: str
    refers_to: Union[str, Any]  # Formula or Reference object

@dataclass
class WorkbookPlan:
    worksheets: List[WorksheetPlan]
    defined_names: List[DefinedNamePlan] = field(default_factory=list)

    def validate(self):
        from excel_saas.core.excel.naming import is_valid_defined_name, sanitize_worksheet_name, sanitize_table_name
        
        # Check worksheet uniqueness
        sheet_names = set()
        for ws in self.worksheets:
            s_name = sanitize_worksheet_name(ws.name).lower()
            if s_name in sheet_names:
                raise ValueError(f"Duplicate worksheet name detected: '{ws.name}'")
            sheet_names.add(s_name)
            
        # Check table uniqueness and validity
        table_names = set()
        for ws in self.worksheets:
            for tbl in ws.tables:
                t_name = sanitize_table_name(tbl.name).lower()
                if t_name in table_names:
                    raise ValueError(f"Duplicate table name detected: '{tbl.name}'")
                table_names.add(t_name)
                
        # Check defined names validity and uniqueness
        def_names = set()
        for dn in self.defined_names:
            if not is_valid_defined_name(dn.name):
                raise ValueError(f"Invalid defined name: '{dn.name}'")
            d_name = dn.name.lower()
            if d_name in def_names:
                raise ValueError(f"Duplicate defined name detected: '{dn.name}'")
            # Defined names can't conflict with tables if they are used globally without context
            if d_name in table_names:
                raise ValueError(f"Defined name '{dn.name}' conflicts with a table name")
            def_names.add(d_name)
