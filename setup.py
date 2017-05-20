# coding: utf-8

from setuptools import setup
from setuptools.extension import Extension
from Cython.Build import cythonize

import os
import sys
import glob


language_level = sys.version_info[0]
_THIS_ = os.path.dirname(os.path.realpath(__file__))

libxlsx_include = os.path.join(_THIS_, "libs", "XLSX", "include")
libxlsx_lib = os.path.join(_THIS_, "libs", "XLSX", "lib")
libxlsx_compiled = os.path.join(_THIS_, "libs", "XLSX", "bin")
zlib_compiled = os.path.join(_THIS_, "libs", "zlib", "bin")

module_name = "impeller"
MODULE_PATH = os.path.join(_THIS_, module_name)

workbook = os.path.join(MODULE_PATH, "c_workbook.pyx")
worksheet = os.path.join(MODULE_PATH, "c_worksheet.pyx")
common = os.path.join(MODULE_PATH, "c_common.pyx")
formatting = os.path.join(MODULE_PATH, "c_format.pyx")
chart = os.path.join(MODULE_PATH, "c_chart.pyx")

extensions = [
    Extension(name="impeller.c_workbook",
              sources=[workbook],
              include_dirs=[_THIS_, libxlsx_include, MODULE_PATH],
              library_dirs=[libxlsx_lib],
              libraries=['xlsxwriter'],
              language="c"),

    Extension(name="impeller.c_worksheet",
              sources=[worksheet],
              include_dirs=[_THIS_, libxlsx_include, MODULE_PATH],
              library_dirs=[libxlsx_lib],
              libraries=['xlsxwriter'],
              language="c"),

    Extension(name="impeller.c_format",
              sources=[formatting],
              include_dirs=[_THIS_, libxlsx_include, MODULE_PATH],
              library_dirs=[libxlsx_lib],
              libraries=['xlsxwriter'],
              language="c"),

    Extension(name="impeller.c_common",
              sources=[common],
              include_dirs=[_THIS_, libxlsx_include, MODULE_PATH],
              library_dirs=[libxlsx_lib],
              libraries=['xlsxwriter'],
              language="c"),

    Extension(name="impeller.c_chart",
              sources=[chart],
              include_dirs=[_THIS_, libxlsx_include, MODULE_PATH],
              library_dirs=[libxlsx_lib],
              libraries=['xlsxwriter'],
              language="c"),
]

setup(
    name="Impeller",
    version='0.1.dev0',
    packages=['impeller'],
    ext_modules=cythonize(extensions, compiler_directives={'language_level': language_level})
)
