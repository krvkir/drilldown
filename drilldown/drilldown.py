# -*- coding: utf-8 -*-
import logging
from traceback import format_exc
import xlsxwriter
from pandas import isnull, notnull

"""
These classes represent structure of a book.

Presentation details should be avoided to allow writing
to different output formats (i.e. a spreadsheet or a website).
"""


class Header:
    # Title is short summary of what is on the page.
    @property
    def title(self):
        return self._title

    # Description explains details: how values are computed, etc.
    @property
    def description(self):
        return self._description

    def __init__(self, title, description):
        self._title = title
        self._description = description


class Navbar:
    # Top-level (horizontal) structure of the book.
    @property
    def links(self):
        return self._links

    # Top-down (vertical) levels representing drillable structure of the book.
    @property
    def breadcrumbs(self):
        return self._breadcrumbs

    def __init__(self, links, breadcrumbs):
        self._links = links
        self._breadcrumbs = breadcrumbs


class Table:
    # Dataframe with `columns`, `index`, `values`.
    @property
    def frame(self):
        return self._frame

    # Groups are folded using outlines (with total row).
    # Formatting visually separates groups between each other.
    @property
    def groups(self):
        return self._groups

    # Visual properties.
    @property
    def column_widths(self):
        return self._column_widths


    def __init__(self, frame, groups=None, column_widths=None):
        self._frame = frame
        self._groups = groups
        self._column_widths = column_widths


class Page:
    # Page name -- used for references.
    @property
    def name(self):
        return self._name

    # Page parent (if needed).
    @property
    def parent(self):
        return self._parent

    # Page structural elements: header, navbar, data table.
    @property
    def header(self):
        return self._header

    @property
    def navbar(self):
        return self._navbar

    @property
    def table(self):
        return self._table

    def __repr__(self):
        return f"Page(name={self.name})"

    def __init__(self, name, header, table, navbar=None, parent=None):
        self._name = name
        self._parent = parent
        self._header = header
        self._navbar = navbar
        self._table = table

    def set_navbar(self, navbar):
        """ Navbar can be set afterwards"""
        self._navbar = navbar


class Cell:
    """
    Cell has few abstract properties describing its content: string, background color.

    They do not depend on the exact backend. Renderer should take care
    of converting them to backend parlance.
    """

    # @property
    # def value(self):
    #     return self._get_value()

    @property
    def string(self):
        return self._get_string()

    def __str__(self):
        return self._get_string()

    @property
    def color(self):
        return self._get_color()

    @property
    def link(self):
        return self._get_link()

    # def _get_value(self):
    #     raise NotImplemented()

    def _get_string(self):
        # User must implement this method to use cell.
        raise NotImplemented()

    def _get_color(self):
        # By default cell has no color. User may override this.
        return None

    def _get_link(self):
        # By default cell has no link. User may override this.
        return None


class MergedArea:
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

    def __init__(self, filename, formats):
        self._filename = filename
        assert {'title', 'description',
                'table_index_column_names', 'table_column_names',
                'table_index_cells', 'table_cells',
                'link_mixin'} <= set(formats.keys())
        self._formats = formats
        self._pages = []

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

        # Put data on the page.
        # Write rows:
        for row_num, (row_index, row_values) in enumerate(frame.iterrows()):
            # Write indices:
            no_values_have_changed = True
            for col_num, index_cell in enumerate(row_index):
                # Initialize merged areas on first row.
                if row_num == 0:
                    merged_areas[col_num] = MergedArea(header_shift, col_num, index_cell)
                    continue
                # If cell value is equal to one stored in MergedArea for this column,
                # and no cells have changed on prevous columns,
                # then just extend MergedArea for this column by one row.
                if no_values_have_changed and merged_areas[col_num].cell_is_equal_to(index_cell):
                    merged_areas[col_num].extend(1, 0)
                    continue
                # Otherwise, write the current MergedArea and create new one.
                sheet.merge_range(*(merged_areas[col_num].coords), "")
                self._write_cell(*(merged_areas[col_num].coords[:2]),
                                 merged_areas[col_num].cell,
                                 sheet,
                                 book,
                                 self._formats['table_index_cells'])
                merged_areas[col_num] = MergedArea(header_shift+row_num, col_num, index_cell)
                no_values_have_changed = False
            # Write cells:
            for col_num, cell in enumerate(row_values):
                self._write_cell(header_shift+row_num,
                                 index_shift+col_num,
                                 cell,
                                 sheet,
                                 book,
                                 self._formats['table_cells'])

        # Flush final merged areas.
        for area in merged_areas:
            sheet.merge_range(*(area.coords), "")
            self._write_cell(*(area.coords[:2]),
                             area.cell,
                             sheet,
                             book,
                             self._formats['table_index_cells'])

    def _write_cell(self, row_num, col_num, cell, sheet, book, default_format):
        # Skip empty cells.
        if isnull(cell):
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
