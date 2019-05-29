# coding: utf-8

from libc.stdint cimport uint32_t, uint16_t, uint8_t
from impeller.c_worksheet cimport *
from impeller.c_workbook cimport *
from impeller.c_common cimport *
from impeller.c_format cimport *
from cython.view cimport array as cvarray

from .exceptions import ImpellerInvalidParameterError
from .compatibility import num_types, str_types

import warnings
import datetime
import math
import re


range_parts = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')

cpdef xl_cell_to_rowcol(cell_str):
    """
    Convert a cell reference in A1 notation to a zero indexed row and column.

    Args:
       cell_str:  A1 style string.

    Returns:
        row, col: Zero indexed cell row and column indices.

    """
    if not cell_str:
        return 0, 0

    match = range_parts.match(cell_str)
    col_str = match.group(2)
    row_str = match.group(4)

    # Convert base26 column string to number.
    expn = 0
    col = 0
    for char in reversed(col_str):
        col += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1

    # Convert 1-index to zero-index
    row = int(row_str) - 1
    col -= 1

    return row, col

def convert_column_args(method):
    """
    Decorator function to convert A1 notation in columns method calls
    to the default row/col notation.

    """
    def column_wrapper(self, *args, **kwargs):

        try:
            # First arg is an int, default to row/col notation.
            if len(args):
                int(args[0])
        except ValueError:
            # First arg isn't an int, convert to A1 notation.
            cell_1, cell_2 = [col + '1' for col in args[0].split(':')]
            _, col_1 = xl_cell_to_rowcol(cell_1)
            _, col_2 = xl_cell_to_rowcol(cell_2)
            new_args = [col_1, col_2]
            new_args.extend(args[1:])
            args = new_args

        return method(self, *args, **kwargs)

    return column_wrapper

def convert_range_args(method):
    """
    Decorator function to convert A1 notation in range method calls
    to the default row/col notation.

    """
    def cell_wrapper(self, *args, **kwargs):

        try:
            # First arg is an int, default to row/col notation.
            if len(args):
                int(args[0])
        except ValueError:
            # First arg isn't an int, convert to A1 notation.
            if ':' in args[0]:
                cell_1, cell_2 = args[0].split(':')
                row_1, col_1 = xl_cell_to_rowcol(cell_1)
                row_2, col_2 = xl_cell_to_rowcol(cell_2)
            else:
                row_1, col_1 = xl_cell_to_rowcol(args[0])
                row_2, col_2 = row_1, col_1

            new_args = [row_1, col_1, row_2, col_2]
            new_args.extend(args[1:])
            args = new_args

        return method(self, *args, **kwargs)

    return cell_wrapper

def convert_cell_args(method):
    """
    Decorator function to convert A1 notation in cell method calls
    to the default row/col notation.

    """
    def cell_wrapper(self, *args, **kwargs):

        try:
            # First arg is an int, default to row/col notation.
            if len(args):
                int(args[0])
        except ValueError:
            # First arg isn't an int, convert to A1 notation.
            new_args = list(xl_cell_to_rowcol(args[0]))
            new_args.extend(args[1:])
            args = new_args

        return method(self, *args, **kwargs)

    return cell_wrapper

# TODO: User must provide additional arguments as named arguments
# TODO: Function needs further attention
# set_column("A:D", 60, cell_format=..., options={...})
#             col   width  etc
def convert_column_args2(*args):
    """
    Used to be a decorator in Python version
    """
    try:
        # First arg is an int, default to row/col notation.
        if len(args):
            int(args[0])
        return args
    except ValueError:
        # First arg isn't an int, convert to A1 notation.
        cell_1, cell_2 = [col + '1' for col in args[0].split(':')]
        _, col_1 = xl_cell_to_rowcol(cell_1)
        _, col_2 = xl_cell_to_rowcol(cell_2)
        new_args = [col_1, col_2]
        new_args.extend(args[1:])
        if new_args[2] is None:
            # Deleting Nonned last_col
            # func("A:D", None, width=..., cell_format=...)
            # OR func2("A:D", None)
            del new_args[2]
        if len(new_args) > 3 and new_args[3] is None:
            # Deleting Nonned width
            # func("A:E", 60, ...)
            del new_args[3]
        args = new_args
        return args


# TODO: Some methods are decorated to access nice looking ranges
# TODO: Not compatible with Py version - Py ver does not accept name or anything in __init__
cdef class Worksheet:
    cdef void _init_defaults(self):
        self.strings_to_numbers = False
        self.strings_to_urls = True
        self.nan_inf_to_errors = False
        self.strings_to_formulas = True

    def __cinit__(self, name, *args, **kwargs):
        self.name = name
        name_bytes = pystring_to_c(self.name)
        self.c_name = name_bytes
        self._init_defaults()

    cdef void _set_ptr(self, lxw_worksheet* ptr):
        self.this_ptr = ptr

    cdef void _add_sheet(self, Workbook wb):
        cdef lxw_worksheet* ws_ptr = workbook_add_worksheet(wb.this_ptr, self.c_name)
        self._set_ptr(ws_ptr)

    # def __dealloc__(self):
    #     if self.this_ptr is not NULL:
    #         lxw_worksheet_free(self.this_ptr)

    cdef lxw_format* _c_format(self, Format cell_format):
        cdef lxw_format* frmt
        if cell_format is None:
            frmt = NULL
        else:
            frmt = cell_format.this_ptr
        return frmt

    cpdef void activate(self):
        worksheet_activate(self.this_ptr)

    @convert_range_args
    def autofilter(self, int first_row, int first_col, int last_row, int last_col):
        raise_on_error(worksheet_autofilter(self.this_ptr, first_row, first_col, last_row, last_col))

    cpdef void center_horizontally(self):
        worksheet_center_horizontally(self.this_ptr)

    cpdef void center_vertically(self):
        worksheet_center_vertically(self.this_ptr)

    cpdef void fit_to_pages(self, int width, int height):
        worksheet_fit_to_pages(self.this_ptr, width, height)

    @convert_cell_args
    def freeze_panes(self, int row, int col, top_row=None, left_col=None, int pane_type=0):
        if top_row is None:
            top_row = row

        if left_col is None:
            left_col = col

        worksheet_freeze_panes_opt(self.this_ptr, row, col, <int>top_row, <int>left_col, pane_type)

    # TODO: Absolutely no idea what top_row and left_col do. Example may be?
    # TODO: Since C API _opt function is undocumented should call simple if top_row and left_col both is None
    @convert_cell_args
    def split_panes(self, float x, float y, top_row=None, left_col=None):
        # In C API defaults is zeros
        if top_row is None:
            top_row = 0
        if left_col is None:
            left_col = 0
        # Beware, y and x swapped
        worksheet_split_panes_opt(self.this_ptr, y, x, top_row, left_col)

    cpdef void hide_gridlines(self, int option=1):
        worksheet_gridlines(self.this_ptr, py_gridlines_to_c(option))

    cpdef void hide(self):
        worksheet_hide(self.this_ptr)

    cpdef void hide_zero(self):
        worksheet_hide_zero(self.this_ptr)

    cpdef void print_across(self):
        worksheet_print_across(self.this_ptr)

    @convert_range_args
    def print_area(self, int first_row, int first_col, int last_row, int last_col):
        raise_on_error(worksheet_print_area(self.this_ptr, first_row, first_col, last_row, last_col))

    cpdef void print_row_col_headers(self):
        worksheet_print_row_col_headers(self.this_ptr)

    cpdef void right_to_left(self):
        worksheet_right_to_left(self.this_ptr)

    cpdef void select(self):
        worksheet_select(self.this_ptr)

    cpdef void set_default_row(self, float height=LXW_DEF_ROW_HEIGHT, bint hide_unused_rows=False):
        worksheet_set_default_row(self.this_ptr, height, hide_unused_rows)

    cpdef void set_first_sheet(self):
        worksheet_set_first_sheet(self.this_ptr)

    cpdef void set_landscape(self):
        worksheet_set_landscape(self.this_ptr)

    cpdef void set_portrait(self):
        worksheet_set_portrait(self.this_ptr)

    cpdef void set_print_scale(self, int scale):
        if scale < 10 or scale > 400:
            raise ImpellerInvalidParameterError("Scale in set_print_scale must be greater than 10 and "
                                                "less than 400. Your value is %d" % scale)
        worksheet_set_print_scale(self.this_ptr, scale)

    cpdef void set_page_view(self):
        worksheet_set_page_view(self.this_ptr)

    # TODO: Python version not documented
    cpdef void set_paper(self, int paper_size):
        if not paper_size < 0:
            raise ImpellerInvalidParameterError("Invalid paper size index in set_paper")
        worksheet_set_paper(self.this_ptr, paper_size)

    cpdef void set_start_page(self, int start_page):
        worksheet_set_start_page(self.this_ptr, start_page)

    cpdef void set_tab_color(self, color):
        worksheet_set_tab_color(self.this_ptr, pystring_to_c(color))

    cpdef void set_zoom(self, int zoom=100):
        if zoom < 10 or zoom > 400:
            raise ImpellerInvalidParameterError("Zoom in set_zoom must be greater than 10 and "
                                                "less than 400. Your value is %d" % zoom)
        worksheet_set_zoom(self.this_ptr, zoom)


    # margin marked as 'for backward compatibility'
    cpdef void set_footer(self, footer='', dict options={}, margin=None):
        # TODO: Options are completely not supported
        cdef lxw_header_footer_options opts
        if options:
            warnings.warn("No options supported in set_footer, except 'margin'")
        if margin is None:
            # default Excel
            margin = options.get('margin', 0.3)
        opts.margin = margin
        raise_on_error(worksheet_set_footer_opt(self.this_ptr, pystring_to_c(footer), &opts))

    # TODO: Refactor a little
    cpdef void set_header(self, header='', dict options={}, margin=None):
        # TODO: options are completely unsupported
        cdef lxw_header_footer_options opts
        if options:
            warnings.warn("No options supported in set_footer, except 'margin'")
        if margin is None:
            # default Excel
            margin = options.get('margin', 0.3)
        opts.margin = margin
        raise_on_error(worksheet_set_header_opt(self.this_ptr, pystring_to_c(header), &opts))

    # TODO: Not very useful
    cdef uint32_t* _get_c_array(self, py_list):
        # +1 for last zero. See http://libxlsxwriter.github.io/worksheet_8h.html#a9601745a2e9e7b1e194b7f5283f197f0
        # TODO: Format I4 is not reliable way to define. Is it?
        cdef lxw_row_t[::1] cyarray = cvarray(shape=(len(py_list) + 1,), itemsize=sizeof(lxw_row_t), format='I4')
        for i in range(len(py_list)):
            cyarray[i] = <lxw_row_t>py_list[i]
        cyarray[len(py_list)] = 0
        return &cyarray[0]

    # TODO: TEST TEST TEST Does it really do not need memory allocation?
    # breaks can be destroyed after method call, cause C API memcpy it anyway
    cpdef void set_h_pagebreaks(self, breaks):
        raise_on_error(worksheet_set_h_pagebreaks(self.this_ptr, <lxw_row_t*>self._get_c_array(breaks)))

    cpdef void set_v_pagebreaks(self, breaks):
        raise_on_error(worksheet_set_v_pagebreaks(self.this_ptr, <lxw_row_t*>self._get_c_array(breaks)))

    cpdef void set_margins(self, float left=0.7, float right=0.7, float top=0.75, float bottom=0.75):
        worksheet_set_margins(self.this_ptr, left, right, top, bottom)

    @convert_column_args
    def set_column(self, firstcol, lastcol, width=None, Format cell_format=None, dict options={}):
        cdef lxw_row_col_options opts
        # TODO: Only hidden option is supported
        if 'collapsed' in options or 'level' in options:
            warnings.warn("Only hidden option supported in set_column")
        opts.hidden = options.get('hidden', False)
        opts.collapsed = options.get('collapsed', False)
        opts.level = options.get('level', 0)
        if width is None:
            if options.get('hidden', False):
                width = 0
            else:
                width = LXW_DEF_COL_WIDTH
        if width == 0:
            opts.hidden = True
        raise_on_error(worksheet_set_column_opt(self.this_ptr, firstcol, lastcol, width,
                                                self._c_format(cell_format), &opts))

    # TODO: Refactor a little
    cpdef void set_row(self, int row, height=None, Format cell_format=None, dict options={}):
        cdef lxw_row_col_options opts
        # TODO: Only hidden option is supported
        if 'collapsed' in options or 'level' in options:
            warnings.warn("Only hidden option supported in set_row")
        opts.hidden = options.get('hidden', False)
        opts.collapsed = options.get('collapsed', False)
        opts.level = options.get('level', 0)
        if height is None:
            if options.get('hidden', False):
                height = 0
            else:
                height = LXW_DEF_ROW_HEIGHT
        if height == 0:
            opts.hidden = True
        raise_on_error(worksheet_set_row_opt(self.this_ptr, row, height, self._c_format(cell_format),
                                             &opts))


    cdef void _check_protection_options(self, dict opts):
        supported_protection_options = {'sheet', 'content', 'options', 'scenarios', 'format_cells', 'format_columns',
                                        'format_rows', 'insert_columns', 'insert_rows', 'insert_hyperlinks',
                                        'delete_columns', 'delete_rows', 'select_locked_cells', 'sort', 'autofilter',
                                        'pivot_tables', 'select_unlocked_cells'}
        for key in opts.keys():
            if key not in supported_protection_options:
                raise ImpellerInvalidParameterError("Invalid protection option %s is not supported" % key)

    cpdef void protect(self, password='', dict options={}):
        cdef lxw_protection protection
        self._check_protection_options(options)
        protection.no_sheet = not options.get('sheet', True)
        protection.content = options.get('content', False)
        protection.objects = options.get('options', False)
        protection.scenarios = options.get('scenarios', False)
        protection.format_cells = options.get('format_cells', False)
        protection.format_columns = options.get('format_columns', False)
        protection.format_rows = options.get('format_rows', False)
        protection.insert_columns = options.get('insert_columns', False)
        protection.insert_rows = options.get('insert_rows', False)
        protection.insert_hyperlinks = options.get('insert_hyperlinks', False)
        protection.delete_columns = options.get('delete_columns', False)
        protection.delete_rows = options.get('delete_rows', False)
        protection.no_select_locked_cells = not options.get('select_locked_cells', True)
        protection.sort = options.get('sort', False)
        protection.autofilter = options.get('autofilter', False)
        protection.pivot_tables = options.get('pivot_tables', False)
        protection.no_select_unlocked_cells = not options.get('select_unlocked_cells', True)

        worksheet_protect(self.this_ptr, pystring_to_c(password), &protection)

    @convert_column_args
    def repeat_columns(self, first_col, last_col=None):
        # TODO: Check for zero too?
        if last_col is None:
            last_col = first_col
        raise_on_error(worksheet_repeat_columns(self.this_ptr, first_col, last_col))

    cpdef void repeat_rows(self, int first_row, last_row=None):
        # TODO: Check for zero too?
        if last_row is None:
            last_row = first_row
        raise_on_error(worksheet_repeat_rows(self.this_ptr, first_row, last_row))


    @convert_cell_args
    def insert_chart(self, int row, int col, Chart chart, dict options={}):
        cdef lxw_image_options opts
        # TODO: There is todo in .c file about chart defaults
        x_offset = options.get("x_offset", 0)
        y_offset = options.get("y_offset", 0)
        x_scale = options.get("x_scale", 1)
        y_scale = options.get("y_scale", 1)

        # Allow Chart to override the scale and offset.
        # TODO: There is not such fields in lxw_chart structure
        # if chart.this_ptr.x_scale != 1:
        #     x_scale = chart.this_ptr.x_scale
        #
        # if chart.this_ptr.y_scale != 1:
        #     y_scale = chart.this_ptr.y_scale
        #
        # if chart.this_ptr.x_offset:
        #     x_offset = chart.this_ptr.x_offset
        #
        # if chart.this_ptr.y_offset:
        #     y_offset = chart.this_ptr.y_offset

        opts.x_offset = x_offset
        opts.y_offset = y_offset
        opts.x_scale = x_scale
        opts.y_scale = y_scale

        raise_on_error(worksheet_insert_chart_opt(self.this_ptr, row, col, chart.this_ptr, &opts))

    @convert_cell_args
    def insert_image(self, int row, int col, filename, dict options={}):
        cdef lxw_image_options opts

        x_offset = options.get("x_offset", 0)
        y_offset = options.get("y_offset", 0)
        x_scale = options.get("x_scale", 1)
        y_scale = options.get("y_scale", 1)
        opts.x_offset = x_offset
        opts.y_offset = y_offset
        opts.x_scale = x_scale
        opts.y_scale = y_scale

        url = options.get('url', None)
        tip = options.get('tip', None)
        # TODO: Two options are not used int .c version
        anchor = options.get('positioning', None)
        image_data = options.get('image_data', None)
        if anchor or image_data:
            raise ImpellerInvalidParameterError("positioning and image_data options in insert_image "
                                                "are not supported yet")
        if url:
            url = pystring_to_c(url)
            opts.url = url
        else:
            opts.url = NULL
        if tip:
            tip = pystring_to_c(tip)
            opts.tip = tip
        else:
            opts.tip = NULL

        raise_on_error(worksheet_insert_image_opt(self.this_ptr, row, col, pystring_to_c(filename), &opts))


    @convert_range_args
    def merge_range(self, int first_row, int first_col, int last_row, int last_col,
                           data, Format cell_format=None):
        raise_on_error(worksheet_merge_range(self.this_ptr, first_row, first_col, last_row, last_col,
                                             b'', self._c_format(cell_format)))
        self.write(first_row, first_col, data, cell_format)

    @convert_range_args
    def set_selection(self, int first_row, int first_col, int last_row, int last_col):
        worksheet_set_selection(self.this_ptr, first_row, first_col, last_row, last_col)

    # TODO: C function is huge and needs some attention
    @convert_cell_args
    def write_url(self, row, col, url, Format cell_format=None, string=None, tip=None):
        cdef char* c_string
        cdef char* c_tip
        # Allow empty string?
        if string is None:
            bytes_string = pystring_to_c(url)
            c_string = bytes_string
        else:
            bytes_string = pystring_to_c(string)
            c_string = bytes_string
        if tip is None:
            c_tip = NULL
        else:
            tip_bytes = pystring_to_c(tip)
            c_tip = tip_bytes
        raise_on_error(worksheet_write_url_opt(self.this_ptr, row, col, pystring_to_c(url),
                                               self._c_format(cell_format), c_string, c_tip))

    cpdef void write_array_formula_strict(self, int first_row, int first_col, int last_row, int last_col, formula,
                                   Format cell_format=None, float value=0):
        raise_on_error(worksheet_write_array_formula_num(self.this_ptr, first_row, first_col,
                                                         last_row, last_col, pystring_to_c(formula),
                                                         self._c_format(cell_format), value))

    # TODO: Consider seriously to remove convert_args decorator. Critical
    # It seems not very Pythonic, mess things up, slow down, mess type hints

    @convert_range_args
    def write_array_formula(self, first_row, first_col, last_row, last_col, formula,
                                   Format cell_format=None, float value=0):
        raise_on_error(worksheet_write_array_formula_num(self.this_ptr, first_row, first_col,
                                                         last_row, last_col, pystring_to_c(formula),
                                                         self._c_format(cell_format), value))

    cpdef void write_formula_strict(self, int row, int col, formula, Format cell_format=None, float value=0):
        raise_on_error(worksheet_write_formula_num(self.this_ptr, row, col, pystring_to_c(formula),
                                                   self._c_format(cell_format), value))

    @convert_cell_args
    def write_formula(self, row, col, formula, Format cell_format=None, float value=0):
        raise_on_error(worksheet_write_formula_num(self.this_ptr, row, col, pystring_to_c(formula),
                                                   self._c_format(cell_format), value))

    cpdef void write_datetime_strict(self, int row, int col, dtm date, Format cell_format=None):
        raise_on_error(worksheet_write_datetime(self.this_ptr, row, col, convert_datetime(date),
                                                self._c_format(cell_format)))

    @convert_cell_args
    def write_datetime(self, row, col, dtm date, Format cell_format=None):
        raise_on_error(worksheet_write_datetime(self.this_ptr, row, col, convert_datetime(date),
                                                self._c_format(cell_format)))

    cpdef void write_blank_strict(self, int row, int col, blank, Format cell_format=None):
        if cell_format is None:
            raise ImpellerInvalidParameterError("Blank cells without format ignored by Excel")
        raise_on_error(worksheet_write_blank(self.this_ptr, row, col, self._c_format(cell_format)))

    @convert_cell_args
    def write_blank(self, row, col, blank, Format cell_format=None):
        if cell_format is None:
            raise ImpellerInvalidParameterError("Blank cells without format ignored by Excel")
        raise_on_error(worksheet_write_blank(self.this_ptr, row, col, self._c_format(cell_format)))

    cpdef void write_boolean_strict(self, int row, int col, bint boolean, Format cell_format=None):
        raise_on_error(worksheet_write_boolean(self.this_ptr, row, col, boolean, self._c_format(cell_format)))

    @convert_cell_args
    def write_boolean(self, row, col, boolean, Format cell_format=None):
        raise_on_error(worksheet_write_boolean(self.this_ptr, row, col, boolean, self._c_format(cell_format)))

    cpdef void write_number_strict(self, int row, int col, double data, Format cell_format=None):
        raise_on_error(worksheet_write_number(self.this_ptr, row, col, data, self._c_format(cell_format)))

    @convert_cell_args
    def write_number(self, row, col, data, Format cell_format=None):
        self.write_number_strict(row, col, data, cell_format)

    cpdef void write_string_strict(self, int row, int col, data, Format cell_format=None):
        data_bytes = _ustring(data).encode("utf8")
        raise_on_error(worksheet_write_string(self.this_ptr, row, col, data_bytes, self._c_format(cell_format)))

    @convert_cell_args
    def write_string(self, row, col, data, Format cell_format=None):
        data_bytes = _ustring(data).encode("utf8")
        raise_on_error(worksheet_write_string(self.this_ptr, row, col, data_bytes, self._c_format(cell_format)))

    @convert_cell_args
    def write(self, row, col, *args):
        # Check the number of args passed.
        if not len(args):
            raise TypeError("write() takes at least 4 arguments (3 given)")

        # The first arg should be the token for all write calls.
        token = args[0]
        try:
            fmt = args[1]
        except IndexError:
            fmt = None

        # Write string types.
        if isinstance(token, str_types):
            # Map the data to the appropriate write_*() method.
            if token == '':
                return self.write_blank_strict(row, col, None, fmt)

            elif self.strings_to_formulas and token.startswith('='):
                # TODO: Missed value
                return self.write_formula_strict(row, col, token, fmt)

            elif self.strings_to_urls and re.match('(ftp|http)s?://', token):
                # TODO: Missing strict version
                return self.write_url(row, col, *args)

            elif self.strings_to_urls and re.match('mailto:', token):
                return self.write_url(row, col, *args)

            elif self.strings_to_urls and re.match('(in|ex)ternal:', token):
                return self.write_url(row, col, *args)

            # TODO: Dropped self._isnan, self._isinf - dropped 2.5 and Jython support
            elif self.strings_to_numbers:
                try:
                    f = float(token)
                    if (self.nan_inf_to_errors or
                            (not math.isnan(f) and not math.isinf(f))):
                        return self.write_number_strict(row, col, f, fmt)
                except ValueError:
                    # Not a number, write as a string.
                    pass

                return self.write_string_strict(row, col, token, fmt)

            else:
                # We have a plain string.
                return self.write_string_strict(row, col, token, fmt)

        # Write number types.
        if isinstance(token, num_types):
            return self.write_number_strict(row, col, token, fmt)

        # Write None as a blank cell.
        if token is None:
            return self.write_blank_strict(row, col, None, fmt)

        # Write boolean types.
        if isinstance(token, bool):
            return self.write_boolean_strict(row, col, token, fmt)

        # TODO: Dropped support of time, timedelta, date
        # Write datetime objects.
        if isinstance(token, datetime.datetime):
            return self.write_datetime_strict(row, col, token, fmt)

        # We haven't matched a supported type. Try float.
        try:
            f = float(token)
            return self.write_number_strict(row, col, f, fmt)
        except ValueError:
            pass
        except TypeError:
            raise TypeError("Unsupported type %s in write()" % type(token))

        # Finally try string.
        try:
            str(token)
            return self.write_string_strict(row, col, token, fmt)
        except ValueError:
            raise TypeError("Unsupported type %s in write()" % type(token))
