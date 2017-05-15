# coding: utf8

"""
common.h from libxlsxwriter + various utilities
"""
import sys
pyver = sys.version_info[0]

from cpython.mem cimport PyMem_Malloc, PyMem_Realloc, PyMem_Free
from cpython.datetime cimport datetime as dtm
from impeller.c_common cimport lxw_error, lxw_datetime
from impeller.c_format cimport *
from impeller.exceptions import *
from impeller.c_chart cimport *
from impeller.c_worksheet cimport *

cdef lxw_datetime* convert_datetime(dtm value):
    """
    Converts Python datetime to lxw_datetime struct
    :param value: datetime to convert
    :type value: datetime.datetime
    :return:
    """
    cdef lxw_datetime* c_date = <lxw_datetime*>PyMem_Malloc(sizeof(lxw_datetime))
    if c_date is NULL:
        raise MemoryError("Failed to malloc temporary lxw_datetime")
    c_date.year = value.year
    c_date.month = value.month
    c_date.day = value.day
    c_date.hour = value.hour
    c_date.min = value.minute
    cdef double micro_fractions = 0.001 * 0.001 * value.microsecond
    # For some reason does not work with double value
    c_date.sec = value.second + round(micro_fractions, ndigits=3)
    return c_date

cdef void free_datetime(lxw_datetime* c_date):
    """
    Free object, created in convert_datetime.
    We can free it after pass to C, because workbook_set_custom_property_datetime uses memcpy inside
    :param c_date:
    :return:
    """
    PyMem_Free(c_date)


default_exception_params = {"ex_class": ImpellerOtherError,
                            "lxw_message": "Other error"}
# For performance considerations create plain vector?
# On the other hand additional overhead is tiny compared to other things
exception_table = {
    LXW_ERROR_MEMORY_MALLOC_FAILED: {"ex_class": ImpellerMemoryError,
                                     "lxw_message": "Memory error, failed to malloc() required memory."},
    LXW_ERROR_CREATING_XLSX_FILE: {"ex_class": ImpellerFileError,
                                   "lxw_message": "Error creating output xlsx file. Usually a permissions error."},
    LXW_ERROR_CREATING_TMPFILE: {"ex_class": ImpellerFileError,
                                 "lxw_message": "Error encountered when creating a tmpfile during file assembly."},
    LXW_ERROR_ZIP_FILE_OPERATION: {"ex_class": ImpellerZipError,
                                   "lxw_message": "Zlib error with a file operation while creating xlsx file."},
    LXW_ERROR_ZIP_FILE_ADD: {"ex_class": ImpellerZipError,
                             "lxw_message": "Zlib error when adding sub file to xlsx file."},
    LXW_ERROR_ZIP_CLOSE: {"ex_class": ImpellerZipError,
                          "lxw_message": "Zlib error when closing xlsx file."},
    LXW_ERROR_NULL_PARAMETER_IGNORED: {"ex_class": ImpellerParameterError,
                                       "lxw_message": "NULL function parameter ignored."},
    LXW_ERROR_PARAMETER_VALIDATION: {"ex_class": ImpellerParameterError,
                                     "lxw_message": "Function parameter validation error."},
    LXW_ERROR_SHEETNAME_LENGTH_EXCEEDED: {"ex_class": ImpellerLongWorksheetNameError,
                                          "lxw_message": "Worksheet name exceeds Excel's limit of 31 characters."},
    LXW_ERROR_INVALID_SHEETNAME_CHARACTER: {"ex_class": ImpellerInvalidWorksheetNameError,
                                            "lxw_message": "Worksheet name contains invalid Excel character: '[]:*?/\\'"},
    LXW_ERROR_SHEETNAME_ALREADY_USED: {"ex_class": ImpellerWorksheetNameUsedError,
                                       "lxw_message": "Worksheet name is already in use."},
    LXW_ERROR_128_STRING_LENGTH_EXCEEDED: {"ex_class": ImpellerLongParameterError,
                                           "lxw_message": "Parameter exceeds Excel's limit of 128 characters."},
    LXW_ERROR_255_STRING_LENGTH_EXCEEDED: {"ex_class": ImpellerLongParameterError,
                                           "lxw_message": "Parameter exceeds Excel's limit of 255 characters."},
    LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED: {"ex_class": ImpellerLongParameterError,
                                           "lxw_message": "String exceeds Excel's limit of 32,767 characters."},
    LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND: {"ex_class": ImpellerSharedStringIndexError,
                                              "lxw_message": "Error finding internal string index."},
    LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE: {"ex_class": ImpellerIndexOutOfRangeError,
                                             "lxw_message": "Worksheet row or column index out of range."},
    LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED: {"ex_class": ImpellerMaxURLsExceededError,
                                                   "lxw_message": "Maximum number of worksheet URLs (65530) exceeded."},
    LXW_ERROR_IMAGE_DIMENSIONS: {"ex_class": ImpellerImageError,
                                 "lxw_message": "Couldn't read image dimensions or DPI."},
    LXW_MAX_ERRNO: {"ex_class": ImpellerOtherError,
                    "lxw_message": "Other error"}
}

# TODO: Make decorator?
cpdef void raise_on_error(lxw_error result):
    """
    Check C function return code and raise appropriate exception if needed
    :param result:
    :return:
    """
    if result != LXW_NO_ERROR:
        if result in exception_table:
            exc_params = exception_table[result]
        else:
            exc_params = default_exception_params
        exc_class = exc_params["ex_class"]
        exc_message = exc_params["lxw_message"]
        raise exc_class(message=exc_message, code=result)

# Example from http://docs.cython.org/en/latest/src/tutorial/strings.html
# TODO: Should be modified?
cdef unicode _ustring(s):
    if type(s) is unicode:
        # fast path for most common case(s)
        return <unicode>s
    elif isinstance(s, bytes) and pyver < 3:
        # only accept byte strings in Python 2.x, not in Py3
        return (<bytes>s).decode('utf8')
    elif isinstance(s, unicode):
        # an evil cast to <unicode> might work here in some(!) cases,
        # depending on what the further processing does.  to be safe,
        # we can always create a copy instead
        return unicode(s)
    else:
        raise TypeError("Seems like wrong string type")

cdef bytes pystring_to_c(s):
    s_bytes = _ustring(s).encode("utf8")
    return s_bytes

cdef unicode c_to_u(const char* s):
    return s.decode('utf8')


available_colors = {
    'UNSET': LXW_COLOR_UNSET,
    'black': LXW_COLOR_BLACK,
    'blue': LXW_COLOR_BLUE,
    'brown': LXW_COLOR_BROWN,
    'cyan': LXW_COLOR_CYAN,
    'gray': LXW_COLOR_GRAY,
    'green': LXW_COLOR_GREEN,
    'lime': LXW_COLOR_LIME,
    'magenta': LXW_COLOR_MAGENTA,
    'navy': LXW_COLOR_NAVY,
    'orange': LXW_COLOR_ORANGE,
    'pink': LXW_COLOR_PINK,
    'purple': LXW_COLOR_PURPLE,
    'red': LXW_COLOR_RED,
    'silver': LXW_COLOR_SILVER,
    'white': LXW_COLOR_WHITE,
    'yellow': LXW_COLOR_YELLOW
}


cdef int pystring_to_color(color_string):
    if color_string in available_colors:
        return available_colors[color_string]
    else:
        # string as HTML #RRGGBB
        return int(color_string[1:], base=16)


available_underlines = {
    1: LXW_UNDERLINE_SINGLE,
    2: LXW_UNDERLINE_DOUBLE,
    33: LXW_UNDERLINE_SINGLE_ACCOUNTING,
    34: LXW_UNDERLINE_DOUBLE_ACCOUNTING
}


cdef int py_underline_to_c(int underline):
    if underline in available_underlines:
        return available_underlines[underline]
    raise ImpellerInvalidParameterError("Invalid underline param. {1, 2, 33, 34} available only")


available_alignments = {
    'left': LXW_ALIGN_LEFT,
    'center': LXW_ALIGN_CENTER,
    'centre': LXW_ALIGN_CENTER,
    'right': LXW_ALIGN_RIGHT,
    'fill': LXW_ALIGN_FILL,
    'justify': LXW_ALIGN_JUSTIFY,
    'center_across': LXW_ALIGN_CENTER_ACROSS,
    'centre_across': LXW_ALIGN_CENTER_ACROSS,
    'distributed': LXW_ALIGN_DISTRIBUTED,
    'justify_distributed': LXW_ALIGN_DISTRIBUTED,
    'top': LXW_ALIGN_VERTICAL_TOP,
    'vcenter': LXW_ALIGN_VERTICAL_CENTER,
    'vcentre': LXW_ALIGN_VERTICAL_CENTER,
    'bottom': LXW_ALIGN_VERTICAL_BOTTOM,
    'vjustify': LXW_ALIGN_VERTICAL_JUSTIFY,
    'vdistributed': LXW_ALIGN_VERTICAL_DISTRIBUTED
}

cdef int py_align_to_c(alignment):
    if alignment in available_alignments:
        return available_alignments[alignment]
    raise ImpellerInvalidParameterError("Invalid %s alignment" % alignment)


# TODO: Missing stock chart type. Copy from ChartStock?
available_chart_types = {
    'doughnut_': LXW_CHART_DOUGHNUT,
    'line_': LXW_CHART_LINE,
    'pie_': LXW_CHART_PIE,

    'area_': LXW_CHART_AREA,
    'area_stacked': LXW_CHART_AREA_STACKED,
    'area_percent_stacked': LXW_CHART_AREA_STACKED_PERCENT,

    'bar_': LXW_CHART_BAR,
    'bar_stacked': LXW_CHART_BAR_STACKED,
    'bar_percent_stacked': LXW_CHART_BAR_STACKED_PERCENT,

    'column_': LXW_CHART_COLUMN,
    'column_stacked': LXW_CHART_COLUMN_STACKED,
    'column_percent_stacked': LXW_CHART_COLUMN_STACKED_PERCENT,

    'scatter_': LXW_CHART_SCATTER,
    'scatter_straight_with_markers': LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS,
    'scatter_straight': LXW_CHART_SCATTER_STRAIGHT,
    'scatter_smooth_with_markers': LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS,
    'scatter_smooth': LXW_CHART_SCATTER_SMOOTH,

    'radar_': LXW_CHART_RADAR,
    'radar_with_markers': LXW_CHART_RADAR_WITH_MARKERS,
    'radar_filled': LXW_CHART_RADAR_FILLED
}

cdef int py_chart_options_to_c(chart_type, chart_subtype):
    separator = '_'
    chart_type_full = chart_type + separator + chart_subtype
    if chart_type_full in available_chart_types:
        return available_chart_types[chart_type_full]
    raise ImpellerInvalidParameterError("Invalid chart type %s with subtype %s" % chart_type, chart_subtype)


available_gridlines = {
    2: LXW_HIDE_ALL_GRIDLINES,
    1: LXW_SHOW_SCREEN_GRIDLINES,
    # TODO: Not used in Python ver?
    3: LXW_SHOW_PRINT_GRIDLINES,
    0: LXW_SHOW_ALL_GRIDLINES
}

cdef int py_gridlines_to_c(int gridlines_type):
    if gridlines_type in available_gridlines:
        return available_gridlines[gridlines_type]
    raise ImpellerInvalidParameterError("Invalid grinlines param %d" % gridlines_type)
