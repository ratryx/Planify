# Planify

A Python architecture for generating professional, declarative Excel workbooks using XlsxWriter. 
The current engine is a reusable declarative Excel-generation core intended to support additional templates/domains later.

Currently implemented template: **Financeiro Pessoal** (`finance_personal`)

## Feature Scope

The `finance_personal` workbook generates a comprehensive personal finance environment supporting 15 integrated domains:
- Contas
- Cartões
- Lançamentos
- Faturas
- Parcelamentos
- Orçamento
- Metas
- Reserva de Emergência
- Investimentos
- Dashboard de Investimentos
- Patrimônio
- Dashboard Patrimônio
- Dívidas
- Projeções
- Análises Financeiras

Features:
- Light and Dark themes natively generated from the same layout logic.
- Targets Microsoft Excel 2021 / Microsoft 365 natively.
- No VBA / macros used.
- Formulas generated declaratively through a Python AST layer.

## Architecture

This project strictly separates:
- **Domain**: Templates (e.g., `finance_personal`) dictate *what* exists in the workbook.
- **WorkbookPlan**: A declarative representation of the workbook, sheets, and tables. Includes `DefinedNamePlan` for robust validation handling.
- **WorkbookEngine**: Knows *how* to translate a `WorkbookPlan` into XlsxWriter operations.
- **Themes**: Dictate *how* cells and tables look based on semantic roles. Separated from templates; the exact same code generates Light and Dark files.
- **Formulas & References**: Python AST layer (`formulas.py`, `references.py`) ensures that formulas are generated safely and robustly without raw string concatenations in template code.

## Generation API

The workbook can be generated programmatically using the internal API:

```python
from excel_saas.application.generate_workbook import generate
from excel_saas.core.models.generation_request import GenerationRequest

request = GenerationRequest(
    template_id="finance_personal",
    year=2026,
    theme="light",
    with_sample_data=False,
    locale="pt_BR",
    currency="BRL",
    profile="default",
    reserve_months=6,
    projection_horizon=12
)

path = generate(request, "output")
```

Generation options supported:
- `template_id`: Which template to run (currently `finance_personal`)
- `year`: Base year for dates and projections
- `locale` / `currency`: Localization preferences
- `theme`: `"light"` or `"dark"`
- `with_sample_data`: Injects test data if True
- `profile`: User demographic profile 
- `reserve_months`: Target emergency reserve coverage in months
- `projection_horizon`: Effectively bounded by the `finance_personal` projection implementation to 1..60 months.

## Important Product Limitations

1. **Projeções** contain only KNOWN commitments: credit card schedules + structural debt payments.
2. **Projeções** do NOT predict future income, variable expenses, future cash balance, or future net worth.
3. The **Recorrente?** tag in `Lançamentos` does NOT currently extrapolate recurring transactions automatically.
4. **"Posição patrimonial registrada"** reflects only registered information explicitly tracked in the workbook.
5. Card debt existing before use of Planify is not automatically known because cards do not have an opening-balance field.
6. Microsoft Excel runtime recalculation validation is a release gate and is separate from openpyxl structural tests.

## UX & Technical Decisions

- **Table UX Over Protection**: Excel does not natively expand structured Tables on protected worksheets (even when `insert_rows` is allowed). Our strategy is to leave input-heavy sheets like `Lançamentos` unprotected to prioritize natural table expansion over preventing accidental edits. Calculation sheets (like Dashboards) remain protected.
- **Stable References**: Data validations use stable workbook defined names rather than volatile `INDIRECT()` references. The `Projeções` worksheet intentionally uses `TODAY()` to maintain a rolling forward horizon.
- **Compatibility**: Targets Microsoft Excel 2021+ and Microsoft 365 natively.

## Running Tests

Ensure you have installed the development dependencies:
```bash
pip install -e ".[dev]"
```

Run tests:
```bash
pytest
```
