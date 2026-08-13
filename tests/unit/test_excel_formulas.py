import pytest
from excel_saas.core.excel.references import TableRef, StringRef
from excel_saas.core.excel.formulas import Formula, sumifs, indirect, _val

def test_table_ref():
    assert str(TableRef("tblLancamentos")) == "tblLancamentos"
    assert str(TableRef("tblLancamentos", "Valor")) == "tblLancamentos[Valor]"

def test_string_ref():
    assert str(StringRef("A1:B10")) == "A1:B10"

def test_formula_builder():
    f = Formula("SUM(A1:A10)")
    assert str(f) == "=SUM(A1:A10)"
    
    # Pre-pended equal sign shouldn't double up
    f2 = Formula("=SUM(A1:A10)")
    assert str(f2) == "=SUM(A1:A10)"

def test_val_coercion():
    assert _val(10) == "10"
    assert _val("Test") == '"Test"'
    assert _val("Test\"Quote") == '"Test""Quote"'
    assert _val(True) == "TRUE"
    assert _val(None) == '""'
    assert _val(TableRef("tbl", "col")) == "tbl[col]"
    assert _val(Formula("=A1")) == "A1"

def test_sumifs_builder():
    val_ref = TableRef("tbl", "Valor")
    crit_ref = TableRef("tbl", "Tipo")
    
    f = sumifs(val_ref, crit_ref, "Receita")
    assert str(f) == '=SUMIFS(tbl[Valor],tbl[Tipo],"Receita")'
    
    with pytest.raises(ValueError):
        sumifs(val_ref, crit_ref) # Missing criteria

def test_indirect_builder():
    cat_ref = TableRef("tblCategorias", "Categoria")
    f = indirect(cat_ref)
    assert str(f) == '=INDIRECT(tblCategorias[Categoria])'
