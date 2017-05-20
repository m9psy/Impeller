# coding: utf-8

from libc.stdlib cimport malloc, free
from libc.stdint cimport uint32_t, uint16_t, uint8_t

from cpython.datetime cimport datetime as dtm

from impeller.c_workbook cimport workbook_new_opt, lxw_workbook_free, lxw_workbook_options, lxw_workbook, \
                                 lxw_doc_properties, py_chart_options_to_c
from impeller.c_worksheet cimport Worksheet
from impeller.c_common cimport raise_on_error, convert_datetime, free_datetime, _ustring, pystring_to_c
from .exceptions import ImpellerInvalidParameterError

from .compatibility import *

import warnings

import datetime

cdef class Workbook:
    cdef void _check_options(self, dict options):
        """
        Only "constant_memory", "tmpdir" supported options, say something if user provide other ones
        :param options:
        :type options: dict
        :return:
        """
        known_options = {"constant_memory", "tmpdir"}
        opt_keys = set(options.keys())
        if len(opt_keys - known_options) > 0:
            warnings.warn('Only "constant_memory" and "tmpdir" supported as options')

    def __cinit__(self, filename, dict options={}, *args, **kwargs):
        cdef lxw_workbook_options opts
        filename_bytes = _ustring(filename).encode("utf8")
        if options is not None or options != {}:
            self._check_options(options)
            opts.constant_memory = options.get("constant_memory", False)
            tdir = options.get("tmpdir", None)
            if tdir is None:
                opts.tmpdir = NULL
            else:
                tdir = _ustring(tdir).encode("utf8")
                opts.tmpdir = tdir
            self.this_ptr = workbook_new_opt(filename_bytes, &opts)
        else:
            self.this_ptr = workbook_new_opt(filename_bytes, NULL)
        if self.this_ptr is NULL:
            raise MemoryError("Unable to create workbook on C backend, no error codes was provided")

    def __init__(self, filename, options={}):
        """
        Creates Workbook object (orly?). Only supported options are constant_memory and tmpdir
        :param filename:
        :type filename:
        :param options:
        :type options: dict
        :return:
        """
        self.filename = filename

    # dealloc take place in close method automatically
    # def __dealloc__(self):
    #     if self.this_ptr is not NULL:
    #         lxw_workbook_free(self.this_ptr)

    cpdef Worksheet add_worksheet(self, name):
        ws = Worksheet(_ustring(name))
        ws._add_sheet(self)
        return ws

    cpdef Worksheet get_worksheet_by_name(self, name):
        cdef bytes name_bytes = _ustring(name).encode("utf8")
        ws_ptr = workbook_get_worksheet_by_name(self.this_ptr, name_bytes)
        ws = Worksheet(_ustring(name))
        ws._set_ptr(ws_ptr)
        return ws

    cpdef void define_name(self, name, formula):
        raise_on_error(workbook_define_name(self.this_ptr, pystring_to_c(name), pystring_to_c(formula)))

    # TODO: Refactor or hide this set_... methods?
    cpdef void set_custom_property_boolean(self, name, bint value):
        raise_on_error(workbook_set_custom_property_boolean(self.this_ptr, pystring_to_c(name), value))

    cpdef void set_custom_property_datetime(self, name, dtm value):
        """
        Set custom datetime property
        :param name: Parameter name
        :param value: Parameter value
        :type value: datetime.datetime
        :return: void
        """
        cdef lxw_datetime* tmp_ptr = convert_datetime(value)
        raise_on_error(workbook_set_custom_property_datetime(self.this_ptr, pystring_to_c(name), tmp_ptr))
        # We can free it after pass to C, because workbook_set_custom_property_datetime uses memcpy
        free_datetime(tmp_ptr)

    # Use set_number instead
    cpdef void set_custom_property_integer(self, name, uint32_t value):
        raise_on_error(workbook_set_custom_property_integer(self.this_ptr, pystring_to_c(name), value))

    cpdef void set_custom_property_number(self, name, double value):
        raise_on_error(workbook_set_custom_property_number(self.this_ptr, pystring_to_c(name), value))

    cpdef void set_custom_property_string(self, name, value):
        cdef bytes value_bytes = pystring_to_c(value)
        raise_on_error(workbook_set_custom_property_string(self.this_ptr, pystring_to_c(name), value_bytes))

    cpdef void _set_prop(self, name, object value, property_name):
        prop_setters = {
            "bool": self.set_custom_property_boolean,
            "date": self.set_custom_property_datetime,
            "number_int": self.set_custom_property_integer,
            "number": self.set_custom_property_number,
            "text": self.set_custom_property_string
        }
        if property_name not in prop_setters:
            raise ImpellerInvalidParameterError("Invalid property_name. Only " + str(prop_setters.keys()) + " supported")
        prop_setters[property_name](name, value)


    cpdef int set_custom_property(self, name, object value, property_type=None):
        """
        Set a custom document property.
        :param name: The name of the custom property.
        :param value: The value of the custom property.
        :param property_type: The type of the custom property. Optional.
        :type name:
        :type value: varying
        :type property_type:
        :return:
        """
        # Almost fully copied from xlsxwriter.workbook.py
        if name is None or value is None:
            warnings.warn("The name and value parameters must be non-None in "
                          "set_custom_property()")
            return -1

        if not property_type:
            # Determine the property type from the Python type.
            if isinstance(value, bool):
                property_type = 'bool'
            elif isinstance(value, datetime.datetime):
                property_type = 'date'
            elif isinstance(value, int_types):
                property_type = 'number_int'
            elif isinstance(value, num_types):
                property_type = 'number'
            else:
                property_type = 'text'

        if property_type == 'text' and len(value) > 255:
            warnings.warn("Length of 'value' parameter exceeds Excel's limit of 255 "
                 "characters in set_custom_property(): '%s'. Value will be will be cut off" %
                 force_unicode(value))
            # It will anyway raise an exception later
            value = value[:251] + "..."

        if len(name) > 255:
            warnings.warn("Length of 'name' parameter exceeds Excel's limit of 255 "
                          "characters in set_custom_property(): '%s'. Value will be will be cut off" %
                 force_unicode(name))
            # It will anyway raise an exception later
            name = name[:251] + "..."

        self._set_prop(name, value, property_type)

    cpdef void set_properties(self, dict properties={}):
        cdef lxw_doc_properties props
        prop_list = ("title", "subject", "author", "manager", "company", "category",
                     "keywords", "comments", "status", "hyperlink_base")
        if properties:
            # TODO: So how do I use setattr here?
            title = properties.get("title", None)
            subject = properties.get("subject", None)
            author = properties.get("author", None)
            manager = properties.get("manager", None)
            company = properties.get("company", None)
            category = properties.get("category", None)
            keywords = properties.get("keywords", None)
            comments = properties.get("comments", None)
            status = properties.get("status", None)
            hyperlink_base = properties.get("hyperlink_base", None)
            if title:
                title = pystring_to_c(title)
                props.title = title
            else:
                props.title = NULL
            if subject:
                subject = pystring_to_c(subject)
                props.subject = subject
            else:
                props.subject = NULL
            if author:
                author_bytes = _ustring(author).encode("utf8")[:]
                props.author = author_bytes
            else:
                props.author = NULL
            if manager:
                manager = pystring_to_c(manager)
                props.manager = manager
            else:
                props.manager = NULL
            if company:
                company = pystring_to_c(company)
                props.company = company
            else:
                props.company = NULL
            if category:
                category = pystring_to_c(category)
                props.category = category
            else:
                props.category = NULL
            if keywords:
                keywords = pystring_to_c(keywords)
                props.keywords = keywords
            else:
                props.keywords = NULL
            if comments:
                comments = pystring_to_c(comments)
                props.comments = comments
            else:
                props.comments = NULL
            if status:
                status = pystring_to_c(status)
                props.status = status
            else:
                props.status = NULL
            if hyperlink_base:
                hyperlink_base = pystring_to_c(hyperlink_base)
                props.hyperlink_base = hyperlink_base
            else:
                props.hyperlink_base = NULL
            raise_on_error(workbook_set_properties(self.this_ptr, &props))
        else:
            raise_on_error(workbook_set_properties(self.this_ptr, NULL))

    cpdef Format add_format(self, dict properties={}):
        fmt = Format()
        fmt._add_format(self)
        return fmt

    cpdef Chart add_chart(self, dict options={}):
        chrt = Chart(options)
        chart_type = options.get("type", "")
        chart_subtype = options.get("subtype", "")
        chrt._add_chart(self, py_chart_options_to_c(chart_type, chart_subtype))
        return chrt

    cpdef void close(self):
        raise_on_error(workbook_close(self.this_ptr))