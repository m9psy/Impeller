# coding: utf-8

from libc.stdint cimport uint32_t, uint16_t, uint8_t
from libc.time cimport time_t

from cpython.datetime cimport datetime as dtm

from impeller.c_chart cimport *
from impeller.c_format cimport *
from impeller.c_worksheet cimport *
from impeller.c_common cimport *
from impeller.c_chart cimport *

cdef extern from "xlsxwriter.h":
    cdef struct lxw_workbook_options:
        uint8_t constant_memory;
        char *tmpdir;

    ctypedef struct lxw_workbook:
        pass

    cdef struct lxw_doc_properties:
        # The title of the Excel Document.
        char *title;

        # The subject of the Excel Document.
        char *subject;

        # The author of the Excel Document.
        char *author;

        # The manager field of the Excel Document.
        char *manager;

        # The company field of the Excel Document.
        char *company;

        # The category of the Excel Document.
        char *category;

        # The keywords of the Excel Document.
        char *keywords;

        # The comment field of the Excel Document.
        char *comments;

        # The status of the Excel Document.
        char *status;

        # The hyperlink base url of the Excel Document.
        char *hyperlink_base;

        time_t created;


    lxw_workbook *workbook_new(const char *filename);
    lxw_workbook *workbook_new_opt(const char *filename, lxw_workbook_options *options);

    lxw_error workbook_define_name(lxw_workbook *workbook, const char *name, const char *formula);
    lxw_worksheet *workbook_get_worksheet_by_name(lxw_workbook *workbook, const char *name);

    lxw_error workbook_set_custom_property_boolean(lxw_workbook *workbook, const char *name, uint8_t value);
    lxw_error workbook_set_custom_property_datetime(lxw_workbook *workbook, const char *name, lxw_datetime *datetime);
    # Use set_number instead
    lxw_error workbook_set_custom_property_integer(lxw_workbook *workbook, const char *name, uint32_t value);
    lxw_error workbook_set_custom_property_number(lxw_workbook *workbook, const char *name, double value);
    lxw_error workbook_set_custom_property_string(lxw_workbook *workbook, const char *name, const char *value);
    lxw_error workbook_set_properties(lxw_workbook *workbook, lxw_doc_properties *properties);

    lxw_chart *workbook_add_chart(lxw_workbook *workbook, uint8_t chart_type);
    lxw_format *workbook_add_format(lxw_workbook *workbook);
    lxw_worksheet *workbook_add_worksheet(lxw_workbook *workbook, const char *sheetname);

    lxw_error workbook_close(lxw_workbook *workbook);

    void lxw_workbook_free(lxw_workbook *workbook);


cdef class WorkBook:
    cdef lxw_workbook *this_ptr;
    cpdef filename;

    cdef void _check_options(self, dict options);

    cpdef void define_name(self, name, formula);

    cpdef Worksheet get_worksheet_by_name(self, name);
    cpdef Worksheet add_worksheet(self, name);
    cpdef Format add_format(self, dict properties=*);
    cpdef Chart add_chart(self, dict options=*);

    # It have been made as cpdef so user can directly avoiding unnecessary type checks
    cpdef void set_custom_property_boolean(self, name, bint value);
    cpdef void set_custom_property_datetime(self, name, dtm value);
    # Use set_number instead
    cpdef void set_custom_property_integer(self, name, uint32_t value);
    cpdef void set_custom_property_number(self, name, double value);
    cpdef void set_custom_property_string(self, name, value);
    cpdef void _set_prop(self, name, object value, property_name);
    cpdef int set_custom_property(self, name, object value, property_type=*);
    cpdef void set_properties(self, dict properties=*);

    cpdef void close(self);
