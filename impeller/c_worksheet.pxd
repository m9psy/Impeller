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
        
    ctypedef struct lxw_protection:
        # Turn off selection of locked cells. This in on in Excel by default.
        uint8_t no_select_locked_cells;
    
        # Turn off selection of unlocked cells. This in on in Excel by default.
        uint8_t no_select_unlocked_cells;
    
        # Prevent formatting of cells.
        uint8_t format_cells;
    
        # Prevent formatting of columns.
        uint8_t format_columns;
    
        # Prevent formatting of rows.
        uint8_t format_rows;
    
        # Prevent insertion of columns.
        uint8_t insert_columns;
    
        # Prevent insertion of rows.
        uint8_t insert_rows;
    
        # Prevent insertion of hyperlinks.
        uint8_t insert_hyperlinks;
    
        # Prevent deletion of columns.
        uint8_t delete_columns;
    
        # Prevent deletion of rows.
        uint8_t delete_rows;
    
        # Prevent sorting data.
        uint8_t sort;
    
        # Prevent filtering data.
        uint8_t autofilter;
    
        # Prevent insertion of pivot tables.
        uint8_t pivot_tables;
    
        # Protect scenarios.
        uint8_t scenarios;
    
        # Protect drawing objects.
        uint8_t objects;
    
        uint8_t no_sheet;
        uint8_t content;

    # TODO: Only hidden supported
    ctypedef struct lxw_row_col_options:
    # Hide the row/column
        uint8_t hidden;
        uint8_t level;
        uint8_t collapsed;

    ctypedef struct lxw_header_footer_options:
        # Header or footer margin in inches. Excel default is 0.3.
        double margin;

    enum pane_types:
        NO_PANES = 0,
        FREEZE_PANES,
        SPLIT_PANES,
        FREEZE_SPLIT_PANES

    # TODO: Double identifier. Already another one in c_format.pxd
    ctypedef int32_t lxw_color_t;

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
    void worksheet_right_to_left(lxw_worksheet *worksheet);
    void worksheet_select(lxw_worksheet *worksheet);
    void worksheet_set_default_row(lxw_worksheet *worksheet, double height, uint8_t hide_unused_rows);
    void worksheet_set_first_sheet(lxw_worksheet *worksheet);
    void worksheet_set_landscape(lxw_worksheet *worksheet);
    void worksheet_set_portrait(lxw_worksheet *worksheet);
    void worksheet_set_page_view(lxw_worksheet *worksheet);
    void worksheet_set_paper(lxw_worksheet *worksheet, uint8_t paper_type);
    void worksheet_set_print_scale(lxw_worksheet *worksheet, uint16_t scale);
    void worksheet_set_start_page(lxw_worksheet *worksheet, uint16_t start_page);
    lxw_error worksheet_set_h_pagebreaks(lxw_worksheet *worksheet, lxw_row_t breaks[]);
    lxw_error worksheet_set_v_pagebreaks(lxw_worksheet *worksheet, lxw_col_t breaks[]);

    void worksheet_set_margins(lxw_worksheet *worksheet, double left, double right, double top, double bottom);

    lxw_error worksheet_set_footer_opt(lxw_worksheet *worksheet, const char *string,
                                       lxw_header_footer_options *options);
    lxw_error worksheet_set_header_opt(lxw_worksheet *worksheet, const char *string,
                                       lxw_header_footer_options *options);

    void worksheet_protect(lxw_worksheet *worksheet, const char *password, lxw_protection *options);

    lxw_error worksheet_set_column_opt(lxw_worksheet *worksheet, lxw_col_t first_col, lxw_col_t last_col,
                                       double width, lxw_format *format, lxw_row_col_options *options);
    lxw_error worksheet_set_row_opt(lxw_worksheet *worksheet, lxw_row_t row, double height,
                                    lxw_format *format, lxw_row_col_options *options);

    lxw_error worksheet_repeat_columns(lxw_worksheet *worksheet, lxw_col_t first_col, lxw_col_t last_col);
    lxw_error worksheet_repeat_rows(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_row_t last_row);

    lxw_error worksheet_print_area(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col,
                                   lxw_row_t last_row, lxw_col_t last_col);

    lxw_error worksheet_insert_chart_opt(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                         lxw_chart *chart, lxw_image_options *user_options);
    lxw_error worksheet_insert_image_opt(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                         const char *filename, lxw_image_options *options);

    lxw_error worksheet_merge_range(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col,
                                    lxw_row_t last_row, lxw_col_t last_col, const char *string,
                                    lxw_format *format);

    # TODO: Undocumented on C API. Possibly use worksheet_write_url function
    lxw_error worksheet_write_url_opt(lxw_worksheet *worksheet, lxw_row_t row_num, lxw_col_t col_num,
                                      const char *url, lxw_format *format, const char *string,
                                      const char *tooltip);
    lxw_error worksheet_write_array_formula_num(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col,
                                                lxw_row_t last_row, lxw_col_t last_col, const char *formula,
                                                lxw_format *format, double result);
    lxw_error worksheet_write_blank(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, lxw_format *format);
    lxw_error worksheet_write_boolean(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                      int value, lxw_format *format);
    lxw_error worksheet_write_formula_num(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                          const char *formula, lxw_format *format, double result);
    lxw_error worksheet_write_datetime(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                       lxw_datetime *datetime, lxw_format *format);
    lxw_error worksheet_write_number(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                     double number, lxw_format *format);
    lxw_error worksheet_write_string(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col,
                                     const char *string, lxw_format *format);

    void worksheet_set_selection(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col,
                                 lxw_row_t last_row, lxw_col_t last_col);
    void worksheet_set_tab_color(lxw_worksheet *worksheet, lxw_color_t color);
    void worksheet_set_zoom(lxw_worksheet *worksheet, uint16_t scale);
    # TODO: Undocumented on C API. Possibly use simple version
    void worksheet_split_panes_opt(lxw_worksheet *worksheet, double vertical, double horizontal,
                                   lxw_row_t top_row, lxw_col_t left_col);

    lxw_error worksheet_write_array_formula(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col,
                                            lxw_row_t last_row, lxw_col_t last_col, const char *formula,
                                            lxw_format *format);



    void lxw_worksheet_free(lxw_worksheet* worksheet);

    # TODO: How to export define?
    cdef double LXW_DEF_COL_WIDTH = 8.43
    cdef double LXW_DEF_ROW_HEIGHT = 15.0



cdef class WorkSheet:
    cdef lxw_worksheet* this_ptr;
    cdef const char* c_name
    cpdef name;


    cdef lxw_format* _c_format(self, Format cell_format);
    cdef void _set_ptr(self, lxw_worksheet* ptr);
    cdef void _add_sheet(self, WorkBook wb);
    cdef uint32_t* _get_c_array(self, py_list);

    cpdef void activate(self);
    # def autofilter(self, int first_row, int first_col, int last_row, int last_col);

    cpdef void center_horizontally(self);
    cpdef void center_vertically(self);

    cpdef void fit_to_pages(self, int width, int height);

    # No type for top_row, left_col because it can be None
    # def freeze_panes(self, int row, int col, top_row=*, left_col=*, int pane_type=*);
    # def split_panes(self, float x, float y, top_row=*, left_col=*);

    cpdef void hide_gridlines(self, int option=*);
    cpdef void hide(self);
    cpdef void hide_zero(self);
    cpdef void print_across(self);
    # def print_area(self, int first_row, int first_col, int last_row, int last_col);
    cpdef void print_row_col_headers(self);
    cpdef void right_to_left(self);
    cpdef void select(self);
    cpdef void set_default_row(self, float height=*, bint hide_unused_rows=*);
    cpdef void set_landscape(self);
    cpdef void set_portrait(self);
    cpdef void set_page_view(self);
    cpdef void set_first_sheet(self);
    cpdef void set_paper(self, int paper_size);
    cpdef void set_print_scale(self, int scale);
    cpdef void set_start_page(self, int start_page);
    cpdef void set_zoom(self, int zoom=*);
    cpdef void set_tab_color(self, color);

    cpdef void set_h_pagebreaks(self, breaks);
    cpdef void set_v_pagebreaks(self, breaks);

    cpdef void set_margins(self, float left=*, float right=*, float top=*, float bottom=*);

    # No type for margin because it can be None
    # TODO: Inspect C API some more, cause result may change
    cpdef void set_footer(self, footer=*, dict options=*, margin=*);
    cpdef void set_header(self, header=*, dict options=*, margin=*);

    # No type for width/height because it can be None. firstcol may be string - see working with ranges
    # TODO: first_col, last_col, but api has no underscore
    # def set_column(self, firstcol, lastcol, width=*, Format cell_format=*, dict options=*);
    cpdef void set_row(self, int row, height=*, Format cell_format=*, dict options=*);

    cdef void _check_protection_options(self, dict opts);
    cpdef void protect(self, password=*, dict options=*);

    # No type for last_col, because it can be None. first_col can be a range string "A:E"
    # def method type, because it is decorated with helper to resolve range strings
    # def void repeat_columns(self, first_col, last_col=*);
    cpdef void repeat_rows(self, int first_row, last_row=*);

    # def insert_chart(self, int row, int col, Chart chart, dict options=*);
    # def insert_image(self, int row, int col, filename, dict options=*);

    # def  merge_range(self, int first_row, int first_col, int last_row, int last_col, data, Format cell_format=*);

    # def set_selection(self, int first_row, int first_col, int last_row, int last_col);

    # TODO: Optimization Decrease calling overhead, declaring methods as cpdef
    # convert_cell_args decorator must be somewhat rewritten in this case
    # Python API actually returns error codes

    # def write_url(self, int row, int col, url, Format cell_format=*, string=*, tip=*);
    # def write_array_formula(self, int first_row, int first_col, int last_row, int last_col, formula,
    #                                Format cell_format=*, float value=*);
    # def write_blank(self, int row, int col, blank, Format cell_format=*);
    # def write_boolean(self, int row, int col, bint boolean, Format cell_format=*);
    # def write_formula(self, int row, int col, formula, Format cell_format=*, float value=*);
    # def write_datetime(self, int row, int col, dtm date, Format cell_format=*);
    # def write_number(self, int row, int col, float data, Format cell_format=*);
    # def write_string(self, int row, int col, data, Format cell_format=*);
    cpdef void write(self, int row, int col, data, Format cell_format);


    # Missing in C API
    # TODO: write_rich_string, write_comment
    # TODO: show_comments, show_comments_author
    # TODO: write_row, write_column
    # TODO: insert_textbox
    # TODO: add_sparkline -> HUGE
    # TODO: insert_button
    # TODO: data_validation -> HUGE
    # TODO: conditional_format -> HUGE
    # TODO: add_table -> HUGE method
    # TODO: outline_settings
    # TODO: set_vba_name

    # TODO: get_name
    # TODO: filter_column, filter_column_list <- addition to autofilter
