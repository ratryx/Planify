import xlsxwriter
from typing import Dict, Any, Tuple, Optional
from dataclasses import dataclass
from excel_saas.core.models.workbook_plan import WorkbookPlan, WorksheetPlan, CellPlan, TablePlan, DataValidationPlan
from excel_saas.themes.base import Theme
from excel_saas.core.models.cell_roles import CellRole
from excel_saas.core.excel.naming import sanitize_worksheet_name, sanitize_table_name

@dataclass(frozen=True)
class ResolvedCellFormat:
    role: CellRole
    bold: bool
    size: int
    number_format: Optional[str]
    # In the future, this class will include borders, alignments, fill color, text_wrap, etc.
    # The immutability and hashability ensure perfect cache behavior without dictionary keys.

class FormatManager:
    """Manages creation and caching of XlsxWriter Format objects to avoid limits."""
    def __init__(self, workbook: xlsxwriter.Workbook, theme: Theme):
        self.workbook = workbook
        self.theme = theme
        self.cache: dict[ResolvedCellFormat, xlsxwriter.format.Format] = {}

    def get_format(self, role: CellRole, bold: bool = False, size: int = 12, number_format: str = None) -> xlsxwriter.format.Format:
        resolved = ResolvedCellFormat(role=role, bold=bold, size=size, number_format=number_format)

        if resolved in self.cache:
            return self.cache[resolved]

        fmt_kwargs = {
            'font_name': self.theme.font_family,
            'font_color': self.theme.colors.text_primary,
            'bg_color': self.theme.colors.background,
            'bold': bold,
            'locked': True # By default, cells are locked (protected)
        }

        if size is not None:
            fmt_kwargs['font_size'] = size

        if number_format:
            fmt_kwargs['num_format'] = number_format

        if role == CellRole.INPUT:
            fmt_kwargs['bg_color'] = self.theme.colors.input_background
            fmt_kwargs['font_color'] = self.theme.colors.text_primary
            fmt_kwargs['locked'] = False # Inputs can be edited
            fmt_kwargs['border'] = 1
            fmt_kwargs['border_color'] = self.theme.colors.input_border
        elif role == CellRole.FORMULA:
            fmt_kwargs['bg_color'] = self.theme.colors.surface_secondary
            fmt_kwargs['font_color'] = self.theme.colors.text_secondary
        elif role == CellRole.HEADER:
            fmt_kwargs['bg_color'] = self.theme.colors.surface
            fmt_kwargs['bold'] = True
        elif role == CellRole.TITLE:
            fmt_kwargs['bg_color'] = self.theme.colors.background
            fmt_kwargs['font_color'] = self.theme.colors.accent
            fmt_kwargs['bold'] = True

        fmt = self.workbook.add_format(fmt_kwargs)
        self.cache[resolved] = fmt
        return fmt


class WorkbookEngine:
    def __init__(self, plan: WorkbookPlan, theme: Theme):
        self.plan = plan
        self.theme = theme

    def render(self, filepath: str) -> None:
        self.plan.validate()

        workbook = xlsxwriter.Workbook(filepath)
        format_mgr = FormatManager(workbook, self.theme)

        for dn in self.plan.defined_names:
            workbook.define_name(dn.name, str(dn.refers_to))

        for sheet_plan in self.plan.worksheets:
            self._render_sheet(workbook, sheet_plan, format_mgr)

        workbook.close()

    def _render_sheet(self, workbook: xlsxwriter.Workbook, sheet_plan: WorksheetPlan, format_mgr: FormatManager) -> None:
        safe_name = sanitize_worksheet_name(sheet_plan.name)
        worksheet = workbook.add_worksheet(safe_name)

        if sheet_plan.show_gridlines is False:
            worksheet.hide_gridlines(2)

        if sheet_plan.freeze_panes:
            worksheet.freeze_panes(sheet_plan.freeze_panes)

        if sheet_plan.tab_color:
            worksheet.set_tab_color(sheet_plan.tab_color)

        if sheet_plan.is_protected:
            options = {
                'insert_rows': True,
                'sort': True,
                'autofilter': True,
                'format_cells': True,
                'format_columns': True,
                'format_rows': True
            }
            worksheet.protect('', options) # Protects locked cells but allows table expansion

        for col_idx, width in sheet_plan.column_widths.items():
            worksheet.set_column(col_idx, col_idx, width)

        self._render_cells(worksheet, sheet_plan.cells, format_mgr)
        self._render_tables(worksheet, sheet_plan.tables, format_mgr)

    def _render_cells(self, worksheet: xlsxwriter.worksheet.Worksheet, cells: list[CellPlan], format_mgr: FormatManager) -> None:
        for cell in cells:
            fmt = format_mgr.get_format(cell.role, cell.bold, cell.size, cell.number_format)

            from excel_saas.core.excel.formulas import Formula, Expression

            if cell.formula is not None:
                if isinstance(cell.formula, (Formula, Expression)):
                    f_str = str(cell.formula)
                    if not f_str.startswith('='):
                        f_str = f"={f_str}"
                    if cell.value is not None:
                        worksheet.write_formula(cell.row, cell.col, f_str, fmt, cell.value)
                    else:
                        worksheet.write_formula(cell.row, cell.col, f_str, fmt)

                else:
                    # Raw strings should NEVER become formulas
                    worksheet.write_string(cell.row, cell.col, str(cell.formula), fmt)
            else:
                if cell.value is not None:
                    if isinstance(cell.value, str) and cell.value.startswith('='):
                        # Use write_string to prevent XlsxWriter from interpreting it as formula
                        worksheet.write_string(cell.row, cell.col, cell.value, fmt)
                    else:
                        worksheet.write(cell.row, cell.col, cell.value, fmt)
                else:
                    worksheet.write_blank(cell.row, cell.col, "", fmt)

            if cell.validation:
                self._apply_data_validation(worksheet, cell.row, cell.col, cell.row, cell.col, cell.validation)

    def _render_tables(self, worksheet: xlsxwriter.worksheet.Worksheet, tables: list[TablePlan], format_mgr: FormatManager) -> None:
        for table in tables:
            safe_name = sanitize_table_name(table.name)

            # Prepare column definitions
            columns = []
            for col in table.columns:
                col_def = {'header': col.header}
                if col.formula is not None:
                    f_str = str(col.formula)
                    if not f_str.startswith('='):
                        f_str = f"={f_str}"
                    col_def['formula'] = f_str
                if col.number_format:
                    # Note: Table column formats in xlsxwriter must be applied to the column or cell level,
                    # but we can set it via format object.
                    fmt = format_mgr.get_format(col.role, number_format=col.number_format)
                    col_def['format'] = fmt
                columns.append(col_def)

            options = {
                'name': safe_name,
                'columns': columns,
                'style': self.theme.table_style,
                'total_row': table.show_total_row
            }
            if table.data:
                options['data'] = table.data

            # Calculate range: start_cell to end_cell
            start_row, start_col = xlsxwriter.utility.xl_cell_to_rowcol(table.start_cell)

            num_cols = len(table.columns)
            num_rows = len(table.data) if table.data else 1  # At least 1 empty row for the table

            # Header row + data rows. If total_row is shown, that adds another row.
            total_rows = 1 + num_rows + (1 if table.show_total_row else 0)

            end_row = start_row + total_rows - 1
            end_col = start_col + num_cols - 1

            ref = xlsxwriter.utility.xl_range(start_row, start_col, end_row, end_col)

            worksheet.add_table(ref, options)

            for i, col in enumerate(table.columns):
                c_idx = start_col + i
                options = {}
                if col.hidden:
                    options['hidden'] = True

                if col.width is not None or options:
                    worksheet.set_column(c_idx, c_idx, col.width, None, options)
                if col.validation:
                    # Apply validation from row after header down to end of table
                    self._apply_data_validation(worksheet, start_row + 1, c_idx, end_row, c_idx, col.validation)

    def _apply_data_validation(self, worksheet, first_row, first_col, last_row, last_col, validation: DataValidationPlan):
        options = {
            'validate': validation.validate,
            'input_title': validation.input_title,
            'input_message': validation.input_message,
            'error_title': validation.error_title,
            'error_message': validation.error_message,
            'ignore_blank': validation.ignore_blank
        }

        if validation.source is not None:
            source = validation.source
            if hasattr(source, '__str__') and not isinstance(source, list):
                source = str(source)
            options['source'] = source

        if validation.criteria is not None:
            options['criteria'] = validation.criteria

        if validation.minimum is not None:
            if options.get('criteria') in ('>', '<', '>=', '<=', '==', '!='):
                options['value'] = validation.minimum
            else:
                options['minimum'] = validation.minimum

        if validation.maximum is not None:
            options['maximum'] = validation.maximum

        worksheet.data_validation(first_row, first_col, last_row, last_col, options)
