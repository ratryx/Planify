import pytest
from excel_saas.core.excel.references import TableRef, StringRef, DefinedNameRef, CellRef
from excel_saas.core.excel.formulas import Formula, sumifs, subtract, add, literal, func, _val

def test_table_ref():
    assert str(TableRef("tblLancamentos")) == "tblLancamentos"
    assert str(TableRef("tblLancamentos", "Valor")) == "tblLancamentos[Valor]"

def test_defined_name_ref():
    assert str(DefinedNameRef("lista_categorias")) == "lista_categorias"

def test_cell_ref():
    assert str(CellRef("A1")) == "A1"
    assert str(CellRef("A1", "Comece Aqui")) == "'Comece Aqui'!A1"

def test_string_ref():
    assert str(StringRef("A1:B10")) == "A1:B10"

def test_formula_builder():
    f = Formula("SUM(A1:A10)")
    assert str(f) == "=SUM(A1:A10)"

    # Pre-pended equal sign shouldn't double up
    f2 = Formula("=SUM(A1:A10)")
    assert str(f2) == "=SUM(A1:A10)"

def test_val_coercion():
    assert _val("Receita") == '"Receita"'
    assert _val(True) == "TRUE"
    assert _val(False) == "FALSE"
    assert _val(10.5) == "10.5"
    assert _val(TableRef("tbl", "col")) == "tbl[col]"
    # Test string escaping
    assert _val('Despesa "Especial"') == '"Despesa ""Especial"""'

def test_ast_sumifs():
    val_ref = TableRef("tbl", "Valor")
    tipo_ref = TableRef("tbl", "Tipo")

    expr = sumifs(val_ref, tipo_ref, literal('Receita'))
    assert str(expr) == 'SUMIFS(tbl[Valor],tbl[Tipo],"Receita")'

    f = Formula(expr)
    assert str(f) == '=SUMIFS(tbl[Valor],tbl[Tipo],"Receita")'

def test_ast_arithmetic():
    expr = subtract(func("SUM", "A1:A10"), literal(10))
    assert str(expr) == 'SUM("A1:A10")-10'

    expr2 = add(literal(5), literal(20))
    assert str(expr2) == '5+20'

def test_formula_nesting():
    val_ref = TableRef("tbl", "Valor")
    tipo_ref = TableRef("tbl", "Tipo")
    expr = subtract(
        sumifs(val_ref, tipo_ref, literal('Receita')),
        sumifs(val_ref, tipo_ref, literal('Despesa'))
    )
    assert str(expr) == 'SUMIFS(tbl[Valor],tbl[Tipo],"Receita")-SUMIFS(tbl[Valor],tbl[Tipo],"Despesa")'
    assert str(Formula(expr)) == '=SUMIFS(tbl[Valor],tbl[Tipo],"Receita")-SUMIFS(tbl[Valor],tbl[Tipo],"Despesa")'

def test_ast_new_functions():
    from excel_saas.core.excel.formulas import index_func, aggregate_func, row_func, iferror_func
    
    assert str(index_func(literal("A1:A10"), literal(5))) == 'INDEX("A1:A10",5)'
    assert str(aggregate_func(literal(15), literal(6), literal("array"), literal(1))) == 'AGGREGATE(15,6,"array",1)'
    assert str(row_func(literal("A1"))) == 'ROW("A1")'
    assert str(row_func()) == 'ROW()'
    assert str(iferror_func(literal("value"), literal("fallback"))) == 'IFERROR("value","fallback")'
