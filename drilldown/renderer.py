# -*- coding: utf-8 -*-

import logging
from traceback import format_exc
import xlsxwriter
import pandas as pd
from pandas import isnull, notnull


_default_formats = {
    # Page header
    'title': {
        'font_size': 24,
        'bold': True,
        'align': 'left',
        'text_wrap': False,
    },
    'description': {
        'font_size': 9,
        'align': 'left',
        'text_wrap': False,
    },
    # Table header
    'table_index_column_names': {
        'font_size': 11,
        'align': 'left',
        'text_wrap': False,
        'bold': True,
    },
    'table_column_names': {
        'font_size': 11,
        'align': 'center',
        'text_wrap': True,
        'bold': True,
    },
    # Table body
    'table_index_cells': {
        'font_size': 11,
        'align': 'left',
        'text_wrap': False,
        'bold': True,
    },
    'table_cells': {
        'font_size': 9,
        'align': 'center',
    },
    # Mixins
    'link_mixin': {'font_color': 'blue', 'underline': True},
}

_default_config = {
    'group_border_style': 1,
}


class MergedArea:
    """
    Represents rectangular area of merged cells on a worksheet
    """
    def __init__(self, first_row, first_col, cell):
        self._first_row = first_row
        self._first_col = first_col
        self._last_row = first_row
        self._last_col = first_col
        self._cell = cell

    @property
    def cell(self):
        return self._cell

    def extend(self, num_rows, num_cols):
        self._last_row += num_rows
        self._last_col += num_cols

    @property
    def coords(self):
        return [self._first_row, self._first_col,
                self._last_row, self._last_col]

    def cell_is_equal_to(self, cell):
        return str(self._cell) == str(cell)


class Renderer:
    """
    Renderer doesn't calculate anything. It only renders, paints and does other formatting.
    """
    @property
    def pages(self):
        return self._pages

    def __init__(self, filename, formats={}, config={}):
        self._filename = filename
        self._pages = []

        # Take default formats, update them with user settings.
        self._formats = _default_formats
        for format_name, format_styles in formats.items():
            self._formats[format_name].update(format_styles)
        assert {'title', 'description',
                'table_index_column_names', 'table_column_names',
                'table_index_cells', 'table_cells',
                'link_mixin'} <= set(self._formats.keys())

        # Take default config, update it with user settings.
        self._config = _default_config
        for k, v in config:
            self._config[k] = v

    def add_page(self, page):
        self._pages.append(page)

    def render_pages(self, skip_errors=False):
        """
        skip_errors: if True, only report errors but continue execution.
        """
        # Open the workbook.
        book = xlsxwriter.Workbook(self._filename)
        # Initialize formats in this book.
        formats = {name: book.add_format(props)
                   for name, props in self._formats.items()}
        # Put pages on sheets of the book.
        for page in self.pages:
            try:
                self._render_page(page, book, formats)
            except:
                if skip_errors:
                    exc_str = format_exc()
                    logging.error(f'Failed to process page {page}:\n{exc_str}')
                    continue
                raise

        book.close()

    def _render_page(self, page, book, formats):
        # Create a worksheet.
        sheet = book.add_worksheet(page.name)

        # Put header and navigation on the page.
        frame = page.table.frame
        if not isinstance(frame.index, pd.MultiIndex):
            frame.index = pd.MultiIndex.from_arrays([frame.index.values],
                                                    names=[frame.index.name])
        index_shift = len(frame.index.names)
        header_shift = 4

        sheet.write(0, 0, page.header.title, formats['title'])
        sheet.write(1, 0, page.header.description, formats['description'])
        if page.parent is not None:
            sheet.write_url(2, 0, f"internal:'{page.parent.name}'!A1", string="Go back")

        sheet.freeze_panes(header_shift, index_shift)

        # Set column widths.
        if page.table.column_widths is not None:
            for column_num, column_width in enumerate(page.table.column_widths):
                sheet.set_column(column_num, column_num, column_width)
        # Hide hidden columns.
        for column_num in page.table.hidden_columns:
            sheet.set_column(column_num, column_num, options={'hidden': True})

        # Write table header:
        # index part
        for col_num, index_name in enumerate(frame.index.names):
            sheet.write(header_shift-1,
                        col_num,
                        str(index_name),
                        formats['table_index_column_names'])
        # columns part
        for col_num, column_name in enumerate(frame.columns):
            sheet.write(header_shift-1,
                        index_shift+col_num,
                        str(column_name),
                        formats['table_column_names'])

        # Initialize merged areas list for indices.
        merged_areas = [None] * len(frame.index.names)

        group_level = (page.table.group_level
                       if page.table.group_level is not None
                       else -1)

        # Put data on the page.
        # Write rows:
        for row_num, (row_index, row_values) in enumerate(frame.iterrows()):
            # Write indices:
            value_have_changed_on_level = None
            for col_num, index_cell in enumerate(row_index):
                # Initialize merged areas on first row.
                if row_num == 0:
                    merged_areas[col_num] = MergedArea(header_shift, col_num, index_cell)
                    continue
                # If cell value is equal to one stored in MergedArea for this column,
                # and no cells have changed on prevous columns,
                # then just extend MergedArea for this column by one row.
                if (value_have_changed_on_level is None
                        and merged_areas[col_num].cell_is_equal_to(index_cell)):
                    merged_areas[col_num].extend(1, 0)
                    continue
                # Otherwise, write the current MergedArea and create new one.
                if value_have_changed_on_level is None:
                    value_have_changed_on_level = col_num
                # if group have closed, apply a special style.
                cell_format = self._formats['table_index_cells'].copy()
                if (value_have_changed_on_level is not None
                        and value_have_changed_on_level <= group_level):
                    cell_format.update({'bottom': self._config['group_border_style']})
                sheet.merge_range(*(merged_areas[col_num].coords), "", book.add_format(cell_format))
                self._write_cell(*(merged_areas[col_num].coords[:2]),
                                 merged_areas[col_num].cell,
                                 sheet,
                                 book,
                                 cell_format)
                merged_areas[col_num] = MergedArea(header_shift+row_num, col_num, index_cell)

            # Write cells:
            for col_num, cell in enumerate(row_values):
                # if group on level 0 have closed, apply a special style.
                cell_format = self._formats['table_cells'].copy()
                if (value_have_changed_on_level is not None
                        and value_have_changed_on_level <= group_level):
                    cell_format.update({'top': self._config['group_border_style']})
                self._write_cell(header_shift+row_num,
                                 index_shift+col_num,
                                 cell,
                                 sheet,
                                 book,
                                 cell_format)

        # Flush final merged areas.
        for area in merged_areas:
            sheet.merge_range(*(area.coords), "")
            self._write_cell(*(area.coords[:2]),
                             area.cell,
                             sheet,
                             book,
                             self._formats['table_index_cells'])

        # Draw bottom line to close the table.
        for col_num in range(len(frame.index.names) + len(frame.columns)):
            sheet.write(header_shift + row_num + 1,
                        col_num,
                        None,
                        book.add_format({'top': self._config['group_border_style']}))

    def _write_cell(self, row_num, col_num, cell, sheet, book, default_format):
        # For empty cells only apply the format to the cell.
        if isnull(cell):
            sheet.write(row_num, col_num, None, book.add_format(default_format))
            return
        # Assemble cell format:
        cell_format = default_format.copy()
        # color:
        if hasattr(cell, 'color') and notnull(cell.color):
            cell_format.update({'bg_color': cell.color})
        # Write cell to the sheet.
        if hasattr(cell, 'link') and cell.link is not None:
            link = f"internal:'{cell.link}'!A1"
            cell_format.update(self._formats['link_mixin'])
            cell_format = book.add_format(cell_format)
            sheet.write_url(row_num, col_num, link, cell_format, str(cell))
        else:
            cell_format = book.add_format(cell_format)
            sheet.write(row_num, col_num, str(cell), cell_format)
