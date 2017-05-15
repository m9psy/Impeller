# coding: utf8

from libc.stdint cimport uint32_t, uint16_t, uint8_t
from impeller.c_worksheet cimport *
from impeller.c_workbook cimport *
from impeller.c_common cimport lxw_row_t, lxw_col_t, raise_on_error, pystring_to_c
from impeller.c_format cimport *

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

    cpdef void write_number(self, lxw_row_t row, lxw_col_t col, double data, Format cell_format=None):
        raise_on_error(worksheet_write_number(self.this_ptr, row, col, data, self._c_format(cell_format)))

    cpdef void write_string(self, lxw_row_t row, lxw_col_t col, data, Format cell_format=None):
        data_bytes = _ustring(data).encode("utf8")
        raise_on_error(worksheet_write_string(self.this_ptr, row, col, data_bytes, self._c_format(cell_format)))