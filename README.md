# Excel SaaS Architecture

A Python architecture for generating professional, declarative Excel workbooks using XlsxWriter.

## Architecture

This project strictly separates:
- **Domain**: Templates (e.g., `finance_personal`) dictate *what* exists in the workbook.
- **WorkbookPlan**: A declarative representation of the workbook, sheets, and tables. Includes `DefinedNamePlan` for robust validation handling.
- **WorkbookEngine**: Knows *how* to translate a `WorkbookPlan` into XlsxWriter operations.
- **Themes**: Dictate *how* cells and tables look based on semantic roles. Separated from templates; the exact same code generates Light and Dark files.
- **Formulas & References**: Python AST layer (`formulas.py`, `references.py`) ensures that formulas are generated safely and robustly without raw string concatenations in template code.

## UX & Architecture Decisions

- **Table UX Over Protection**: Excel does not natively expand structured Tables on protected worksheets (even when `insert_rows` is allowed). Our strategy is to leave input-heavy sheets like `Lançamentos` unprotected to prioritize natural table expansion over preventing accidental edits. Calculation sheets (like Dashboard) remain protected.
- **No Volatile Formulas**: Data validation uses stable Workbook Defined Names rather than volatile `INDIRECT()` calls.
- **Compatibility**: Targets Microsoft Excel 2021+ and Microsoft 365 natively.


## Running Tests

Ensure you have installed the development dependencies:
```bash
pip install -e .[dev]
```

Run tests:
```bash
pytest tests/
```
