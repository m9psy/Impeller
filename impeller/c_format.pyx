# coding: utf8

from impeller.c_format cimport *
from impeller.c_common cimport pystring_to_c, raise_on_error, pystring_to_color, py_underline_to_c, py_align_to_c
from impeller.c_workbook cimport workbook_add_format, WorkBook
from impeller.exceptions import ImpellerInvalidParameterError


# TODO: How to set UNSET_COLOR?


cdef class Format:
    def __cinit__(self, dict properties={}, *args, **kwargs):
        pass

    cdef void _set_ptr(self, lxw_format* ptr):
        self.this_ptr = ptr

    cdef void _add_format(self, WorkBook wb):
        cdef lxw_format* new_format = workbook_add_format(wb.this_ptr)
        self._set_ptr(new_format)

    cpdef void set_font_name(self, font_name):
        cdef bytes b_name = pystring_to_c(font_name)
        format_set_font_name(self.this_ptr, b_name)

    cpdef void set_font_size(self, int font_size=11):
        format_set_font_size(self.this_ptr, font_size)

    cpdef void set_font_color(self, color):
        cdef lxw_color_t clr = pystring_to_color(color)
        format_set_font_color(self.this_ptr, clr)

    cpdef void set_bold(self, bint bold=True):
        # C function can set only LXW_TRUE, there is no `unbold` function
        # This method imitate set_bold + set_unbold
        self.this_ptr.bold = bold

    cpdef void set_italic(self, bint italic=True):
        # C function can set only LXW_TRUE, there is no `unstrike` function
        # This method imitate set_strike + set_unstrike
        self.this_ptr.italic = italic

    # TODO: Match Python integer values and C value's
    cpdef void set_underline(self, int underline=1):
        format_set_underline(self.this_ptr, py_underline_to_c(underline))

    cpdef void set_font_strikeout(self, bint font_strikeout=True):
        # C function can set only LXW_TRUE, there is no `unstrikeout` function
        # This method imitate set_strikeout + set_unstrikeout
        self.this_ptr.font_strikeout = font_strikeout

    cpdef void set_font_script(self, int font_script):
        format_set_font_script(self.this_ptr, font_script)

    # TODO:
    # UNDOCUMENTED AND UNSUPPORTED BY C BACKEND
    # cpdef void set_font_outline(self, bint font_outline=True);
    # cpdef void set_font_shadow(self, bint font_shadow=True);

    # Number
    cpdef void set_num_format(self, num_format):
        format_set_num_format(self.this_ptr, pystring_to_c(num_format))

    # Protection
    cpdef void set_locked(self, bint locked=True):
        # C function can set only LXW_FALSE, there is no `set_locked` function
        # This method imitate set_locked + set_unlocked
        self.this_ptr.locked = locked

    cpdef void set_hidden(self, bint hidden=True):
        # C function can set only LXW_TRUE, there is no `set_hidden` function
        # This method imitate set_hidden + set_unhidden
        self.this_ptr.hidden = hidden

    # Alignment
    cpdef void set_align(self, alignment):
        format_set_align(self.this_ptr, py_align_to_c(alignment))

    cpdef void set_center_across(self, align_type=None):
        # TODO: align_type is not used
        format_set_align(self.this_ptr, LXW_ALIGN_CENTER_ACROSS)

    cpdef void set_rotation(self, int rotation):
        rotation = int(rotation)

        # Map user angle to Excel angle.
        if rotation == 270:
            rotation = 255
        elif -90 <= rotation <= 90:
            if rotation < 0:
                rotation = -rotation + 90
        else:
            raise ImpellerInvalidParameterError(
                "Rotation rotation outside range: -90 <= angle <= 90")
        format_set_rotation(self.this_ptr, rotation)

    cpdef void set_indent(self, int indent=1):
        if indent > 255:
            raise ImpellerInvalidParameterError("Wrong indent level. Indent must be >0 and <256")
        format_set_indent(self.this_ptr, indent)

    cpdef void set_text_wrap(self, bint text_wrap=True):
        # C function can set only LXW_TRUE, there is no `set_unwrap` function
        # This method imitate set_wrap + set_unwrap
        self.this_ptr.text_wrap = text_wrap

    cpdef void set_shrink(self, bint shrink=True):
        # C function can set only LXW_TRUE, there is no `set_unshrink` function
        # This method imitate set_shrink + set_unshrink
        self.this_ptr.shrink = shrink

    # TODO: format.h missing set_justlast, but format.c has it. How to proceed?
  	# cpdef void set_text_justlast(self, bint text_justlast=True);
    # Pattern
    # Py values match exactly C enum
    cpdef void set_pattern(self, int pattern=1):
        if pattern < 0 or pattern > 18:
            raise ImpellerInvalidParameterError("Invalid pattern value. Pattern must be between 0 and 18")
        format_set_pattern(self.this_ptr, pattern)

    cpdef void set_bg_color(self, bg_color):
        format_set_bg_color(self.this_ptr, pystring_to_color(bg_color))

    cpdef void set_fg_color(self, fg_color):
        format_set_fg_color(self.this_ptr, pystring_to_color(fg_color))

    cdef void _check_border(self, int style=0):
        if style < 0 or style > 13:
            raise ImpellerInvalidParameterError("Invalid border style. Border must be between 0 and 13")

    # Border
    cpdef void set_border(self, style=1):
        self._check_border(style)
        format_set_border(self.this_ptr, style)

    cpdef void set_bottom(self, bottom=1):
        self._check_border(bottom)
        format_set_bottom(self.this_ptr, bottom)

    cpdef void set_top(self, top=1):
        self._check_border(top)
        format_set_top(self.this_ptr, top)

    cpdef void set_left(self, left=1):
        self._check_border(left)
        format_set_left(self.this_ptr, left)

    cpdef void set_right(self, right=1):
        self._check_border(right)
        format_set_right(self.this_ptr, right)

    cpdef void set_border_color(self, color):
        format_set_border_color(self.this_ptr, pystring_to_color(color))

    cpdef void set_bottom_color(self, bottom_color):
        format_set_bottom_color(self.this_ptr, pystring_to_color(bottom_color))

    cpdef void set_top_color(self, top_color):
        format_set_top_color(self.this_ptr, pystring_to_color(top_color))

    cpdef void set_left_color(self, left_color):
        format_set_left_color(self.this_ptr, pystring_to_color(left_color))

    cpdef void set_right_color(self, right_color):
        format_set_right_color(self.this_ptr, pystring_to_color(right_color))
