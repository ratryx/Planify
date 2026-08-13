# Excel SaaS Architecture

A Python architecture for generating professional, declarative Excel workbooks using XlsxWriter.

## Architecture

This project strictly separates:
- **Domain**: Templates (e.g., `finance_personal`) dictate *what* exists in the workbook.
- **WorkbookPlan**: A declarative representation of the workbook, sheets, and tables.
- **WorkbookEngine**: Knows *how* to translate a `WorkbookPlan` into XlsxWriter operations.
- **Themes**: Dictate *how* cells and tables look based on semantic roles.

## Running Tests

Ensure you have installed the development dependencies:
```bash
pip install -e .[dev]
```

Run tests:
```bash
pytest tests/
```
