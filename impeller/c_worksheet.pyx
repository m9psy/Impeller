# coding: utf8

from libc.stdint cimport uint32_t, uint16_t, uint8_t
from impeller.c_worksheet cimport *
from impeller.c_workbook cimport *
from impeller.c_common cimport lxw_row_t, lxw_col_t, raise_on_error, pystring_to_c, py_gridlines_to_c
from impeller.c_format cimport *

from .exceptions import ImpellerInvalidParameterError

# TODO: Not compatible with Py version - Py ver does not accept name or anything in __init__
cdef class WorkSheet:
    def __cinit__(self, name, *args, **kwargs):
        self.name = name
        name_bytes = pystring_to_c(self.name)
        self.c_name = name_bytes

    cdef void _set_ptr(self, lxw_worksheet* ptr):
        self.this_ptr = ptr

    cdef void _add_sheet(self, WorkBook wb):
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

    cpdef void autofilter(self, int first_row, int first_col, int last_row, int last_col):
        raise_on_error(worksheet_autofilter(self.this_ptr, first_row, first_col, last_row, last_col))

    cpdef void center_horizontally(self):
        worksheet_center_horizontally(self.this_ptr)

    cpdef void center_vertically(self):
        worksheet_center_vertically(self.this_ptr)

    cpdef void fit_to_pages(self, int width, int height):
        worksheet_fit_to_pages(self.this_ptr, width, height)

    cpdef void freeze_panes(self, int row, int col, int top_row=None, int left_col=None, int pane_type=0):
        if top_row is None:
            top_row = row

        if left_col is None:
            left_col = col

        worksheet_freeze_panes_opt(self.this_ptr, row, col, top_row, left_col, pane_type)

    cpdef void hide_gridlines(self, int option=1):
        worksheet_gridlines(self.this_ptr, py_gridlines_to_c(option))

    cpdef void hide(self):
        worksheet_hide(self.this_ptr)

    cpdef void hide_zero(self):
        worksheet_hide_zero(self.this_ptr)

    cpdef void print_across(self):
        worksheet_print_across(self.this_ptr)

    cpdef void print_area(self, int first_row, int first_col, int last_row, int last_col):
        raise_on_error(worksheet_print_area(self.this_ptr, first_row, first_col, last_row, last_col))

    cpdef void print_row_col_headers(self):
        worksheet_print_row_col_headers(self.this_ptr)


    cpdef void insert_chart(self, int row, int col, Chart chart, dict options={}):
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

    cpdef void insert_image(self, int row, int col, filename, dict options={}):
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


    cpdef void merge_range(self, int first_row, int first_col, int last_row, int last_col,
                           data, Format cell_format=None):
        raise_on_error(worksheet_merge_range(self.this_ptr, first_row, first_col, last_row, last_col,
                                             b'', self._c_format(cell_format)))
        self.write(first_row, first_col, data, cell_format)


    cpdef void write_number(self, lxw_row_t row, lxw_col_t col, double data, Format cell_format=None):
        raise_on_error(worksheet_write_number(self.this_ptr, row, col, data, self._c_format(cell_format)))

    cpdef void write_string(self, lxw_row_t row, lxw_col_t col, data, Format cell_format=None):
        data_bytes = _ustring(data).encode("utf8")
        raise_on_error(worksheet_write_string(self.this_ptr, row, col, data_bytes, self._c_format(cell_format)))

    # TODO: Non typed writing
    cpdef void write(self, int row, int col, data, Format cell_format):
        pass