# coding: utf-8

from impeller.c_workbook cimport workbook_add_chart

cdef class Chart:
    def __cinit__(self, dict options={}, *args, **kwargs):
        pass

    cdef void _set_ptr(self, lxw_chart* ptr):
        self.this_ptr = ptr

    cdef void _add_chart(self, Workbook wb, uint8_t chart_type):
        cdef lxw_chart* new_chart = workbook_add_chart(wb.this_ptr, chart_type);
        self._set_ptr(new_chart)

    # TODO: Working with charts
    cpdef void add_series(self, dict options={}):
        pass
