import pytest
from excel_saas.core.excel.naming import sanitize_worksheet_name, sanitize_table_name, is_valid_defined_name, escape_sheet_name
from excel_saas.core.models.workbook_plan import WorkbookPlan, WorksheetPlan, TablePlan, DefinedNamePlan

def test_sanitize_worksheet_name():
    assert sanitize_worksheet_name("Normal") == "Normal"
    assert sanitize_worksheet_name("Invalid/Chars*?[ ]\\:") == "InvalidChars"
    assert sanitize_worksheet_name("Receitas:2026") == "Receitas2026"
    assert sanitize_worksheet_name("A" * 50) == "A" * 31
    assert sanitize_worksheet_name("'Quoted'") == "Quoted"
    assert sanitize_worksheet_name("") == "Sheet"

def test_sanitize_table_name():
    assert sanitize_table_name("My Table") == "My_Table"
    assert sanitize_table_name("1InvalidStart") == "_1InvalidStart"
    assert sanitize_table_name("Valid_Name") == "Valid_Name"

def test_is_valid_defined_name():
    assert is_valid_defined_name("lista_categorias") == True
    assert is_valid_defined_name("Categorias") == True
    assert is_valid_defined_name("_hidden") == True
    
    assert is_valid_defined_name("lista categorias") == False  # space
    assert is_valid_defined_name("1lista") == False  # starts with number
    assert is_valid_defined_name("A1") == False  # looks like cell
    assert is_valid_defined_name("R1C1") == False  # looks like cell

def test_escape_sheet_name():
    assert escape_sheet_name("Normal") == "Normal"
    assert escape_sheet_name("With Spaces") == "'With Spaces'"
    assert escape_sheet_name("João's Data") == "'João''s Data'"

def test_workbook_structural_validation():
    # Valid plan
    plan = WorkbookPlan(
        worksheets=[
            WorksheetPlan(name="Sheet1", cells=[], tables=[TablePlan("tbl1", "A1", [], False)])
        ],
        defined_names=[DefinedNamePlan("lista_1", "tbl1[col1]")]
    )
    plan.validate()  # should not raise
    
    # Duplicate sheet
    with pytest.raises(ValueError, match="Duplicate worksheet"):
        WorkbookPlan(
            worksheets=[
                WorksheetPlan(name="Sheet1", cells=[]),
                WorksheetPlan(name="SHEET1", cells=[])
            ]
        ).validate()
        
    # Duplicate table
    with pytest.raises(ValueError, match="Duplicate table"):
        WorkbookPlan(
            worksheets=[
                WorksheetPlan(name="Sheet1", cells=[], tables=[TablePlan("tbl1", "A1", [], False)]),
                WorksheetPlan(name="Sheet2", cells=[], tables=[TablePlan("TBL1", "A1", [], False)])
            ]
        ).validate()
        
    # Duplicate defined name
    with pytest.raises(ValueError, match="Duplicate defined name"):
        WorkbookPlan(
            worksheets=[],
            defined_names=[
                DefinedNamePlan("lista_1", "ref"),
                DefinedNamePlan("LISTA_1", "ref")
            ]
        ).validate()

    # Invalid defined name
    with pytest.raises(ValueError, match="Invalid defined name"):
        WorkbookPlan(
            worksheets=[],
            defined_names=[DefinedNamePlan("invalid name", "ref")]
        ).validate()
        
    # Defined name conflicts with table
    with pytest.raises(ValueError, match="conflicts with a table"):
        WorkbookPlan(
            worksheets=[WorksheetPlan(name="Sheet1", cells=[], tables=[TablePlan("my_table", "A1", [], False)])],
            defined_names=[DefinedNamePlan("my_table", "ref")]
        ).validate()
