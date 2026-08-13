from typing import Union, Any
from .references import Reference

class Formula:
    """Represents a safely constructed Excel formula."""
    def __init__(self, expression: str):
        self.expression = expression
        
    def __str__(self) -> str:
        # Ensure it starts with '='
        if not self.expression.startswith("="):
            return f"={self.expression}"
        return self.expression

def _val(value: Any) -> str:
    """Converts a Python value/Reference into an Excel formula fragment."""
    if isinstance(value, Reference):
        return str(value)
    if isinstance(value, Formula):
        # strip leading = for embedding
        expr = str(value)
        return expr[1:] if expr.startswith("=") else expr
    if isinstance(value, str):
        # Quote string literals
        escaped = value.replace('"', '""')
        return f'"{escaped}"'
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if value is None:
        return '""'
    return str(value)

def sumifs(sum_range: Union[Reference, str], *criteria_pairs) -> Formula:
    """
    Constructs a SUMIFS formula.
    criteria_pairs should be alternating criteria_range and criteria.
    """
    if len(criteria_pairs) % 2 != 0:
        raise ValueError("SUMIFS requires an even number of criteria arguments.")
        
    args = [_val(sum_range)]
    for i in range(0, len(criteria_pairs), 2):
        args.append(_val(criteria_pairs[i]))
        args.append(_val(criteria_pairs[i+1]))
        
    return Formula(f"SUMIFS({','.join(args)})")

def indirect(ref_text: Union[Reference, str]) -> Formula:
    """Constructs an INDIRECT formula."""
    return Formula(f"INDIRECT({_val(ref_text)})")
