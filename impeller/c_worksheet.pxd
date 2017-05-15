#coding: utf8

from impeller.c_workbook cimport WorkBook
from impeller.c_common cimport *
from impeller.c_format cimport *

cdef extern from "xlsxwriter.h":
    ctypedef struct lxw_worksheet:
        pass

    lxw_error worksheet_write_number(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                     double number, lxw_format *format);
    lxw_error worksheet_write_string(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                     const char *string, lxw_format *format);

    void lxw_worksheet_free(lxw_worksheet* worksheet);

cdef class WorkSheet:
    cdef lxw_worksheet* this_ptr;
    cdef const char* c_name
    cpdef name;


    cdef lxw_format* _c_format(self, Format cell_format);
    cdef void _set_ptr(self, lxw_worksheet* ptr);
    cdef void _add_sheet(self, WorkBook wb);

    cpdef void write_number(self, lxw_row_t row, lxw_col_t col, double data, Format cell_format=*);
    cpdef void write_string(self, lxw_row_t row, lxw_col_t col, data, Format cell_format=*);