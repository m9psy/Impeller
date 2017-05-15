# coding: utf8

# Python xlsxwriter using plain Exceptions, so this will work
from impeller.c_common import *


class AbstractImpellerError(Exception):
    """abc"""
    pass


# ----------------------- Custom Python exceptions -----------------------

class ImpellerInvalidParameterError(AbstractImpellerError, ValueError):
    pass


# ----------------------- Exceptions caused by C code -----------------------

class ImpellerCError(AbstractImpellerError):
    """
    C error - exception caused by bad return code in C function
    """
    lxw_code = None
    message = ""

    def __init__(self, message="", code=0):
        self.message = message
        self.lxw_code = code

    def __repr__(self, *args, **kwargs):
        parent_repr = super(ImpellerCError, self).__repr__(*args, **kwargs)
        parent_repr += "FastXLSXError " + str(self.lxw_code) + " " + self.message
        return parent_repr

    def __str__(self, *args, **kwargs):
        parent_str = super(ImpellerCError, self).__str__(*args, **kwargs)
        parent_str += "Message: " + self.message
        parent_str += "LXW CODE: " + str(self.lxw_code)
        return parent_str


class ImpellerMemoryError(ImpellerCError, MemoryError):
    """
    Memory error, failed to malloc() required memory.
    LXW_ERROR_MEMORY_MALLOC_FAILED code
    """
    pass


class ImpellerFileError(ImpellerCError, IOError):
    """
    Error creating output xlsx file. Usually a permissions error.
    Error encountered when creating a tmpfile during file assembly.
    LXW_ERROR_CREATING_XLSX_FILE, LXW_ERROR_CREATING_TMPFILE codes
    """
    pass


class ImpellerZipError(ImpellerCError, IOError):
    """
    Zlib error with a file operation while creating xlsx file.
    Zlib error when adding sub file to xlsx file.
    Zlib error when closing xlsx file.
    LXW_ERROR_ZIP_FILE_OPERATION, LXW_ERROR_ZIP_FILE_ADD, LXW_ERROR_ZIP_CLOSE codes
    """
    pass


class ImpellerParameterError(ImpellerCError):
    """
    NULL function parameter ignored.
    Function parameter validation error.
    LXW_ERROR_NULL_PARAMETER_IGNORED, LXW_ERROR_PARAMETER_VALIDATION codes
    """
    pass


class ImpellerLongWorksheetNameError(ImpellerCError):
    """
    Worksheet name exceeds Excel's limit of 31 characters.
    LXW_ERROR_SHEETNAME_LENGTH_EXCEEDED code
    """
    pass


class ImpellerInvalidWorksheetNameError(ImpellerCError):
    """
    Worksheet name contains invalid Excel character: '[]:*?/\\'
    LXW_ERROR_INVALID_SHEETNAME_CHARACTER code
    """
    pass


class ImpellerWorksheetNameUsedError(ImpellerCError):
    """
    Worksheet name is already in use.
    LXW_ERROR_SHEETNAME_ALREADY_USED code
    """
    pass


class ImpellerLongParameterError(ImpellerCError):
    """
    Parameter exceeds Excel's limit of 128 characters.
    Parameter exceeds Excel's limit of 255 characters.
    String exceeds Excel's limit of 32,767 characters.
    LXW_ERROR_128_STRING_LENGTH_EXCEEDED, LXW_ERROR_255_STRING_LENGTH_EXCEEDED, LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED
    """
    pass


class ImpellerSharedStringIndexError(ImpellerCError, IndexError):
    """
    Error finding internal string index.
    LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND
    """
    pass


class ImpellerIndexOutOfRangeError(ImpellerCError, IndexError):
    """
    Worksheet row or column index out of range.
    LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE
    """
    pass


class ImpellerMaxURLsExceededError(ImpellerCError):
    """
    Maximum number of worksheet URLs (65530) exceeded.
    LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED
    """
    pass


class ImpellerImageError(ImpellerCError):
    """
    Couldn't read image dimensions or DPI.
    LXW_ERROR_IMAGE_DIMENSIONS
    """
    pass


class ImpellerOtherError(ImpellerCError):
    """
    LXW_MAX_ERRNO
    """
    pass
