# coding: utf-8

from libc.stdint cimport uint32_t, int32_t, uint16_t, int16_t, uint8_t
from impeller.c_workbook cimport WorkBook

cdef extern from "xlsxwriter.h":
    ctypedef struct lxw_format:
        uint8_t bold
        uint8_t italic;
        uint8_t font_strikeout;
        uint8_t hidden;
        uint8_t locked;
        uint8_t text_wrap;
        uint8_t shrink;

    enum lxw_defined_colors:
        LXW_COLOR_BLACK = 0x1000000,
        LXW_COLOR_BLUE = 0x0000FF,
        LXW_COLOR_BROWN = 0x800000,
        LXW_COLOR_CYAN = 0x00FFFF,
        LXW_COLOR_GRAY = 0x808080,
        LXW_COLOR_GREEN = 0x008000,
        LXW_COLOR_LIME = 0x00FF00,
        LXW_COLOR_MAGENTA = 0xFF00FF,
        LXW_COLOR_NAVY = 0x000080,
        LXW_COLOR_ORANGE = 0xFF6600,
        LXW_COLOR_PINK = 0xFF00FF,
        LXW_COLOR_PURPLE = 0x800080,
        LXW_COLOR_RED = 0xFF0000,
        LXW_COLOR_SILVER = 0xC0C0C0,
        LXW_COLOR_WHITE = 0xFFFFFF,
        LXW_COLOR_YELLOW = 0xFFFF00

    enum lxw_format_underlines:
        LXW_UNDERLINE_SINGLE = 1,
        LXW_UNDERLINE_DOUBLE,
        LXW_UNDERLINE_SINGLE_ACCOUNTING,
        LXW_UNDERLINE_DOUBLE_ACCOUNTING
        
    enum lxw_format_alignments:
        # No alignment. Cell will use Excel's default for the data type
        LXW_ALIGN_NONE = 0,
        # Left horizontal alignment
        LXW_ALIGN_LEFT,
        # Center horizontal alignment
        LXW_ALIGN_CENTER,
        # Right horizontal alignment
        LXW_ALIGN_RIGHT,
        # Cell fill horizontal alignment
        LXW_ALIGN_FILL,
        # Justify horizontal alignment
        LXW_ALIGN_JUSTIFY,
        # Center Across horizontal alignment
        LXW_ALIGN_CENTER_ACROSS,
        # Left horizontal alignment
        LXW_ALIGN_DISTRIBUTED,
        # Top vertical alignment
        LXW_ALIGN_VERTICAL_TOP,
        # Bottom vertical alignment
        LXW_ALIGN_VERTICAL_BOTTOM,
        # Center vertical alignment
        LXW_ALIGN_VERTICAL_CENTER,
        # Justify vertical alignment
        LXW_ALIGN_VERTICAL_JUSTIFY,
        # Distributed vertical alignment
        LXW_ALIGN_VERTICAL_DISTRIBUTED

    ctypedef int32_t lxw_color_t;

    void format_set_font_name(lxw_format *format_ptr, const char *font_name);
    void format_set_font_size(lxw_format *format_ptr, uint16_t size);
    void format_set_font_color(lxw_format *format_ptr, lxw_color_t color);
    void format_set_bold(lxw_format *format_ptr);
    void format_set_italic(lxw_format *format_ptr);
    void format_set_underline(lxw_format *format_ptr, uint8_t style);
    void format_set_font_strikeout(lxw_format *format_ptr);
    void format_set_font_script(lxw_format *format_ptr, uint8_t style);

    void format_set_num_format(lxw_format *format_ptr, const char *num_format);

    void format_set_unlocked(lxw_format *format_ptr);
    void format_set_hidden(lxw_format *format_ptr);

    void format_set_align(lxw_format *format_ptr, uint8_t alignment);
    void format_set_text_wrap(lxw_format *format_ptr);
    void format_set_rotation(lxw_format *format_ptr, int16_t angle);
    void format_set_indent(lxw_format *format_ptr, uint8_t level);
    void format_set_text_wrap(lxw_format *format_ptr);
    void format_set_shrink(lxw_format *format_ptr);
    void format_set_pattern(lxw_format *format_ptr, uint8_t index);

    void format_set_bg_color(lxw_format *format_ptr, lxw_color_t color);
    void format_set_fg_color(lxw_format *format_ptr, lxw_color_t color);

    void format_set_border(lxw_format *format_ptr, uint8_t style);
    void format_set_bottom(lxw_format *format_ptr, uint8_t style);
    void format_set_top(lxw_format *format_ptr, uint8_t style);
    void format_set_left(lxw_format *format_ptr, uint8_t style);
    void format_set_right(lxw_format *format_ptr, uint8_t style);

    void format_set_border_color(lxw_format *format_ptr, lxw_color_t color);
    void format_set_bottom_color(lxw_format *format_ptr, lxw_color_t color);
    void format_set_top_color(lxw_format *format_ptr, lxw_color_t color);
    void format_set_left_color(lxw_format *format_ptr, lxw_color_t color);
    void format_set_right_color(lxw_format *format_ptr, lxw_color_t color);

    # TODO: This methods does not present in Format class
    # Undocumented
    void format_set_diag_border(lxw_format *format_ptr, uint8_t value);
    void format_set_diag_color(lxw_format *format_ptr, lxw_color_t color);
    void format_set_diag_type(lxw_format *format_ptr, uint8_t value);

    void format_set_font_charset(lxw_format *format_ptr, uint8_t value);
    void format_set_font_condense(lxw_format *format_ptr);
    void format_set_font_extend(lxw_format *format_ptr);
    void format_set_font_family(lxw_format *format_ptr, uint8_t value);
    void format_set_font_outline(lxw_format *format_ptr);
    void format_set_font_scheme(lxw_format *format_ptr, const char *font_scheme);
    void format_set_font_shadow(lxw_format *format_ptr);

    # void format_set_num_format_index(lxw_format *format_ptr
    void format_set_reading_order(lxw_format *format_ptr, uint8_t value);
    void format_set_text_justlast(lxw_format *format_ptr);
    void format_set_theme(lxw_format *format_ptr, uint8_t value);
    void format_set_valign(lxw_format *format_ptr, uint8_t alignment);

# TODO: Can not import #defined value for some reason
cdef int LXW_COLOR_UNSET = -1;

cdef class Format:
    cdef lxw_format* this_ptr;

    cdef void _set_ptr(self, lxw_format* ptr);

    cdef void _add_format(self, WorkBook wb);

    cpdef void set_font_name(self, font_name);
    cpdef void set_font_size(self, int font_size=*);
    # String with color
    cpdef void set_font_color(self, font_color);
    cpdef void set_bold(self, bint bold=*);
    cpdef void set_italic(self, bint italic=*);
    cpdef void set_underline(self, int underline=*);
    cpdef void set_font_strikeout(self, bint font_strikeout=*);
    cpdef void set_font_script(self, int font_script);
    # UNDOCUMENTED AND UNSUPPORTED BY C BACKEND
    # cpdef void set_font_outline(self, bint font_outline=True);
    # cpdef void set_font_shadow(self, bint font_shadow=True);
    # Number
    cpdef void set_num_format(self, num_format);
    # Protection
    cpdef void set_locked(self, bint locked=*);
    cpdef void set_hidden(self, bint hidden=*);
    # Alignment
    cpdef void set_align(self, alignment);
    cpdef void set_center_across(self, align_type=*);
    cpdef void set_rotation(self, int rotation);
    cpdef void set_indent(self, int indent=*);
    cpdef void set_text_wrap(self, bint text_wrap=*);
    cpdef void set_shrink(self, bint shrink=*);
    # TODO: format.h missing set_justlast, but format.c has it. How to proceed?
  	# cpdef void set_text_justlast(self, bint text_justlast=True);
    # Pattern
    # Py values match exactly C enum
    cpdef void set_pattern(self, int pattern=*);
    cpdef void set_bg_color(self, bg_color);
    cpdef void set_fg_color(self, fg_color);
    # Border
    cdef void _check_border(self, int style=*);
    cpdef void set_border(self, style=*);
    cpdef void set_bottom(self, bottom=*);
    cpdef void set_top(self, top=*);
    cpdef void set_left(self, left=*);
    cpdef void set_right(self, right=*);
    cpdef void set_border_color(self, color);
    cpdef void set_bottom_color(self, bottom_color);
    cpdef void set_top_color(self, top_color);
    cpdef void set_left_color(self, left_color);
    cpdef void set_right_color(self, right_color);
