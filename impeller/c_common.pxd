# coding: utf8
from libc.stdint cimport uint32_t, uint16_t, uint8_t

from cpython.datetime cimport datetime as dtm

cdef extern from "xlsxwriter.h":
    ctypedef enum lxw_error:
        # No error. 
        LXW_NO_ERROR = 0,

        # Memory error, failed to malloc() required memory.
        LXW_ERROR_MEMORY_MALLOC_FAILED,

        # Error creating output xlsx file. Usually a permissions error. 
        LXW_ERROR_CREATING_XLSX_FILE,
    
        # Error encountered when creating a tmpfile during file assembly. 
        LXW_ERROR_CREATING_TMPFILE,
    
        # Zlib error with a file operation while creating xlsx file.
        LXW_ERROR_ZIP_FILE_OPERATION,
    
        # Zlib error when adding sub file to xlsx file.
        LXW_ERROR_ZIP_FILE_ADD,
    
        # Zlib error when closing xlsx file.
        LXW_ERROR_ZIP_CLOSE,
    
        # NULL function parameter ignored.
        LXW_ERROR_NULL_PARAMETER_IGNORED,
    
        # Function parameter validation error. 
        LXW_ERROR_PARAMETER_VALIDATION,
    
        # Worksheet name exceeds Excel's limit of 31 characters. 
        LXW_ERROR_SHEETNAME_LENGTH_EXCEEDED,
    
        # Worksheet name contains invalid Excel character: '[]:*?/\\' 
        LXW_ERROR_INVALID_SHEETNAME_CHARACTER,
    
        # Worksheet name is already in use.
        LXW_ERROR_SHEETNAME_ALREADY_USED,
    
        # Parameter exceeds Excel's limit of 128 characters. 
        LXW_ERROR_128_STRING_LENGTH_EXCEEDED,
    
        # Parameter exceeds Excel's limit of 255 characters. 
        LXW_ERROR_255_STRING_LENGTH_EXCEEDED,
    
        # String exceeds Excel's limit of 32,767 characters. 
        LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED,
    
        # Error finding internal string index.
        LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND,
    
        # Worksheet row or column index out of range. 
        LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE,
    
        # Maximum number of worksheet URLs (65530) exceeded. 
        LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED,
    
        # Couldn't read image dimensions or DPI.
        LXW_ERROR_IMAGE_DIMENSIONS,
    
        LXW_MAX_ERRNO

    # TODO: Конвертация из Py в эту структуру
    ctypedef struct lxw_datetime:
        # Year     : 1900 - 9999
        int year;
        # Month    : 1 - 12
        int month;
        # Day      : 1 - 31
        int day;
        # Hour     : 0 - 23
        int hour;
        # Minute   : 0 - 59
        int min;
        # Seconds  : 0 - 59.999
        double sec;

    ctypedef uint32_t lxw_row_t;

    ctypedef uint32_t lxw_col_t;


# TODO: Switch type to cpython.datetime?
cdef lxw_datetime* convert_datetime(dtm value);
cdef void free_datetime(lxw_datetime* c_date);

cpdef void raise_on_error(lxw_error result);

cdef int pystring_to_color(color_string);
cdef int py_underline_to_c(int underline);
cdef int py_align_to_c(alignment);
cdef int py_chart_options_to_c(chart_type, chart_subtype);

cdef bytes pystring_to_c(s);
cdef unicode _ustring(s);
cdef unicode c_to_u(const char* s);
