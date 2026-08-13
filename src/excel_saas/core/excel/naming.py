import re

# Excel restricts these characters in worksheet names
INVALID_SHEET_CHARS = r'[\\/?*\[\]:]'

def sanitize_worksheet_name(name: str) -> str:
    """
    Sanitizes a worksheet name according to Excel rules:
    - Removes invalid characters: \\ / ? * [ ] :
    - Removes leading and trailing apostrophes
    - Truncates to 31 characters
    - Provides a default if empty
    """
    safe_name = re.sub(INVALID_SHEET_CHARS, '', name)
    safe_name = safe_name.strip(" '")

    if not safe_name:
        return "Sheet"

    return safe_name[:31]

def sanitize_table_name(name: str) -> str:
    """
    Sanitizes an Excel Table name:
    - Must not contain spaces
    - Must start with a letter or underscore
    - (simplified rule for this phase)
    """
    safe_name = name.replace(' ', '_')
    if safe_name and not safe_name[0].isalpha() and safe_name[0] not in ('_', '\\'):
        safe_name = '_' + safe_name
    return safe_name

def is_valid_defined_name(name: str) -> bool:
    """
    Validates an Excel Defined Name.
    Rules:
    - Non-empty
    - Starts with letter or underscore
    - No spaces or invalid characters
    - Does not look like a cell reference (e.g. A1, R1C1)
    """
    if not name:
        return False

    if not re.match(r'^[a-zA-Z_\\][a-zA-Z0-9_\.\?]*$', name):
        return False

    # Check for cell reference pattern (A1 to XFD1048576 or R1C1)
    # simplified regex
    if re.match(r'^[A-Za-z]{1,3}[0-9]{1,7}$', name):
        return False
    if re.match(r'^R[0-9]*C[0-9]*$', name, re.IGNORECASE):
        return False

    return True

def escape_sheet_name(sheet_name: str) -> str:
    """
    Escapes a worksheet name for use in a formula reference.
    - If it contains spaces or special characters, it must be enclosed in single quotes.
    - Existing single quotes must be doubled.
    """
    # Double the single quotes
    escaped = sheet_name.replace("'", "''")
    # Enclose in single quotes if it contains spaces or other non-alphanumeric chars
    if not re.match(r'^[A-Za-z0-9_]+$', sheet_name):
        return f"'{escaped}'"
    return escaped
