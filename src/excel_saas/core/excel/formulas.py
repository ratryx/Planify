from typing import Union, Any, List
from .references import Reference

class Expression:
    """Base class for any formula expression."""
    def __str__(self) -> str:
        raise NotImplementedError

class Formula:
    """
    Represents a safely constructed Excel formula.
    Should be passed to the engine to be rendered with a leading '='.
    """
    def __init__(self, expr: Union[Expression, str]):
        self.expr = expr

    def __str__(self) -> str:
        # Ensure it starts with '='
        s = str(self.expr)
        if not s.startswith("="):
            return f"={s}"
        return s

def _val(value: Any) -> str:
    """Converts a Python value/Reference/Expression into an Excel formula fragment."""
    if isinstance(value, (Reference, Expression)):
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

class FuncExpr(Expression):
    def __init__(self, name: str, *args):
        self.name = name.upper()
        self.args = args

    def __str__(self) -> str:
        rendered_args = ",".join(_val(arg) for arg in self.args)
        return f"{self.name}({rendered_args})"

class BinaryExpr(Expression):
    def __init__(self, left, op: str, right):
        self.left = left
        self.op = op
        self.right = right

    def __str__(self) -> str:
        return f"{_val(self.left)}{self.op}{_val(self.right)}"

def func(name: str, *args) -> Expression:
    """Constructs an Excel function call."""
    return FuncExpr(name, *args)

def sum_func(sum_range) -> Expression:
    return func("SUM", sum_range)

def sumifs(sum_range, *criteria_pairs) -> Expression:
    """
    Constructs a SUMIFS formula expression.
    """
    if len(criteria_pairs) % 2 != 0:
        raise ValueError("SUMIFS requires an even number of criteria arguments.")

    return func("SUMIFS", sum_range, *criteria_pairs)

def subtract(left, right) -> Expression:
    return BinaryExpr(left, "-", right)

def add(left, right) -> Expression:
    return BinaryExpr(left, "+", right)

def equals(left, right) -> Expression:
    return BinaryExpr(left, "=", right)

def divide(left, right) -> Expression:
    return BinaryExpr(left, "/", right)

def negate(value) -> Expression:
    class UnaryExpr(Expression):
        def __str__(self):
            return f"-{_val(value)}"
    return UnaryExpr()

def if_func(condition, true_val, false_val) -> Expression:
    return func("IF", condition, true_val, false_val)

def isnumber(value) -> Expression:
    return func("ISNUMBER", value)

def isblank(value) -> Expression:
    return func("ISBLANK", value)

def and_func(*conditions) -> Expression:
    return func("AND", *conditions)

def or_func(*conditions) -> Expression:
    return func("OR", *conditions)

def not_func(condition) -> Expression:
    return func("NOT", condition)

def greater_than(left, right) -> Expression:
    return BinaryExpr(left, ">", right)

def greater_or_equal(left, right) -> Expression:
    return BinaryExpr(left, ">=", right)

def not_equals(left, right) -> Expression:
    return BinaryExpr(left, "<>", right)

def literal(value: Any) -> Expression:
    """Forces a value to be treated as a literal (useful if you just want to quote a string in an expression tree)."""
    class LiteralExpr(Expression):
        def __str__(self):
            return _val(value)
    return LiteralExpr()
