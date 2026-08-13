import pytest
from excel_saas.core.excel.naming import sanitize_worksheet_name, sanitize_table_name, escape_worksheet_reference

def test_sanitize_worksheet_name():
    assert sanitize_worksheet_name("Normal Name") == "Normal Name"
    assert sanitize_worksheet_name("Name with / and ?") == "Name with  and"
    assert sanitize_worksheet_name("A" * 40) == "A" * 31
    assert sanitize_worksheet_name("'Quoted'") == "Quoted"
    assert sanitize_worksheet_name("[]*?") == "Sheet" # Falls back to Sheet when empty

def test_sanitize_table_name():
    assert sanitize_table_name("My Table") == "My_Table"
    assert sanitize_table_name("123Table") == "_123Table"
    assert sanitize_table_name("tblLançamentos") == "tblLançamentos" # Should allow letters (including accents if handled by regex \w, but regex uses a-zA-Z, let's see)

def test_escape_worksheet_reference():
    assert escape_worksheet_reference("Sheet1") == "Sheet1"
    assert escape_worksheet_reference("My Sheet") == "'My Sheet'"
    assert escape_worksheet_reference("Sheet's Data") == "'Sheet''s Data'"
