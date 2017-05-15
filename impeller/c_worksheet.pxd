#coding: utf8

from libc.stdint cimport uint32_t, uint16_t, uint8_t, int32_t

from impeller.c_workbook cimport WorkBook
from impeller.c_common cimport *
from impeller.c_format cimport *
from impeller.c_chart cimport lxw_chart, Chart

cdef extern from "xlsxwriter.h":
    ctypedef struct lxw_worksheet:
        pass

    ctypedef struct lxw_image_options:

        # Offset from the left of the cell in pixels.
        int32_t x_offset;
        # Offset from the top of the cell in pixels.
        int32_t y_offset;
        # X scale of the image as a decimal.
        double x_scale;
        # Y scale of the image as a decimal.
        double y_scale;

        char *url;
        char *tip;

    enum lxw_gridlines:
        # Hide screen and print gridlines.
        LXW_HIDE_ALL_GRIDLINES = 0,
        # Show screen gridlines.
        LXW_SHOW_SCREEN_GRIDLINES,
        # Show print gridlines.
        LXW_SHOW_PRINT_GRIDLINES,
        # Show screen and print gridlines.
        LXW_SHOW_ALL_GRIDLINES

    lxw_error worksheet_write_number(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                     double number, lxw_format *format);
    lxw_error worksheet_write_string(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                     const char *string, lxw_format *format);

    void worksheet_activate(lxw_worksheet *worksheet);
    lxw_error worksheet_autofilter(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col,
                                   lxw_row_t last_row, lxw_col_t last_col);

    void worksheet_center_horizontally(lxw_worksheet *worksheet);
    void worksheet_center_vertically(lxw_worksheet *worksheet);

    void worksheet_fit_to_pages(lxw_worksheet *worksheet, uint16_t width, uint16_t height);
    void worksheet_freeze_panes_opt(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col,
                                    lxw_row_t top_row, lxw_col_t left_col, uint8_t type);

    # TODO: Option type not uint8, but lxw_gridlines (enum in worksheet.h)?
    void worksheet_gridlines(lxw_worksheet *worksheet, uint8_t option);
    void worksheet_hide(lxw_worksheet *worksheet);
    void worksheet_hide_zero(lxw_worksheet *worksheet);
    void worksheet_print_across(lxw_worksheet *worksheet);
    void worksheet_print_row_col_headers(lxw_worksheet *worksheet);

    lxw_error worksheet_print_area(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col,
                                   lxw_row_t last_row, lxw_col_t last_col);

    lxw_error worksheet_insert_chart_opt(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                         lxw_chart *chart, lxw_image_options *user_options);
    lxw_error worksheet_insert_image_opt(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                         const char *filename, lxw_image_options *options);

    lxw_error worksheet_merge_range(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col,
                                    lxw_row_t last_row, lxw_col_t last_col, const char *string,
                                    lxw_format *format);

    void lxw_worksheet_free(lxw_worksheet* worksheet);

cdef class WorkSheet:
    cdef lxw_worksheet* this_ptr;
    cdef const char* c_name
    cpdef name;


    cdef lxw_format* _c_format(self, Format cell_format);
    cdef void _set_ptr(self, lxw_worksheet* ptr);
    cdef void _add_sheet(self, WorkBook wb);

    cpdef void activate(self);
    cpdef void autofilter(self, int first_row, int first_col, int last_row, int last_col);

    cpdef void center_horizontally(self);
    cpdef void center_vertically(self);

    cpdef void fit_to_pages(self, int width, int height);
    cpdef void freeze_panes(self, int row, int col, int top_row=*, int left_col=*, int pane_type=*);
    cpdef void hide_gridlines(self, int option=*);
    cpdef void hide(self);
    cpdef void hide_zero(self);
    cpdef void print_across(self);
    cpdef void print_area(self, int first_row, int first_col, int last_row, int last_col);
    cpdef void print_row_col_headers(self);

    cpdef void insert_chart(self, int row, int col, Chart chart, dict options=*);
    cpdef void insert_image(self, int row, int col, filename, dict options=*);

    # TODO: Return void?
    cpdef void merge_range(self, int first_row, int first_col, int last_row, int last_col, data, Format cell_format=*);

    cpdef void write_number(self, lxw_row_t row, lxw_col_t col, double data, Format cell_format=*);
    cpdef void write_string(self, lxw_row_t row, lxw_col_t col, data, Format cell_format=*);
    cpdef void write(self, int row, int col, data, Format cell_format);