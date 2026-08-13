import re

def sanitize_worksheet_name(name: str) -> str:
    """
    Sanitizes a worksheet name for Excel.
    - Max 31 characters.
    - Cannot contain: \\, /, ?, *, [, ]
    - Cannot start or end with an apostrophe.
    """
    # Remove invalid characters
    invalid_chars = r'[\\/?*\[\]]'
    clean_name = re.sub(invalid_chars, '', name)
    
    # Strip starting/ending apostrophes and spaces
    clean_name = clean_name.strip(" '")
    
    # Truncate to 31 chars
    clean_name = clean_name[:31]
    
    if not clean_name:
        return "Sheet"
    return clean_name

def sanitize_table_name(name: str) -> str:
    """
    Sanitizes a table name for Excel.
    - Must start with a letter or underscore.
    - Cannot contain spaces or special characters other than underscore.
    - Cannot be a cell reference (e.g., A1, R1C1).
    """
    # Replace spaces and invalid chars with underscores (allow unicode word chars)
    clean_name = re.sub(r'[^\w]', '_', name)
    
    # Must start with letter or underscore (unicode)
    if not re.match(r'^[^\W\d_]|_', clean_name):
        clean_name = '_' + clean_name
        
    return clean_name

def escape_worksheet_reference(name: str) -> str:
    """
    Escapes a worksheet name for use in a formula.
    If the name contains spaces or special characters, it must be wrapped in single quotes.
    Single quotes inside the name must be doubled.
    """
    if re.search(r'[^a-zA-Z0-9_.]', name):
        escaped = name.replace("'", "''")
        return f"'{escaped}'"
    return name
