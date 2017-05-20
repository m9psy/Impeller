# coding: utf-8

from setuptools import setup
from setuptools.extension import Extension
from Cython.Build import cythonize

import os
import sys
import shutil
import warnings


def is_windows():
    return os.name == 'nt'


def check_win_binaries(*binaries):
    for binary in binaries:
        if not os.path.exists(binary):
            warnings.warn("binary is missing in %s directory" % binary)
            return False
    return True


def move_binaries():
    libxlsx_compiled_dir = os.path.join(_THIS_, "libs", "XLSX", "bin")
    zlib_compiled_dir = os.path.join(_THIS_, "libs", "zlib", "bin")
    zlib_compiled = os.path.join(zlib_compiled_dir, "zlib.dll")
    libxlsx_compiled = os.path.join(libxlsx_compiled_dir, "xlsxwriter.dll")

    if check_win_binaries(zlib_compiled, libxlsx_compiled):
        shutil.copyfile(zlib_compiled, os.path.join(MODULE_PATH, "zlib.dll"))
        shutil.copyfile(libxlsx_compiled, os.path.join(MODULE_PATH, "xlsxwriter.dll"))
    else:
        warnings.warn("Some sufficient binaries are missing, try to build them with build_c_libs.cmd")


language_level = sys.version_info[0]
_THIS_ = os.path.dirname(os.path.realpath(__file__))

module_name = "impeller"
MODULE_PATH = os.path.join(_THIS_, module_name)

if is_windows():
    libxlsx_include = os.path.join(_THIS_, "libs", "XLSX", "include")
    libxlsx_lib = os.path.join(_THIS_, "libs", "XLSX", "lib")
    move_binaries()

workbook = [os.path.join(MODULE_PATH, "c_workbook.pyx")]
worksheet = [os.path.join(MODULE_PATH, "c_worksheet.pyx")]
common = [os.path.join(MODULE_PATH, "c_common.pyx")]
formatting = [os.path.join(MODULE_PATH, "c_format.pyx")]
chart = [os.path.join(MODULE_PATH, "c_chart.pyx")]

common_extension_args = {
    "libraries": ['xlsxwriter'],
    "language": "c"
}

OS_specific_extension_args = {}

if is_windows():
    OS_specific_extension_args = {
        "include_dirs": [libxlsx_include],
        "library_dirs": [libxlsx_lib],
    }

# dlls for windows, linux have them in default location
package_data = {}

if is_windows():
    # Instead of providing them in PATH - ship with the lib
    package_data = {'impeller': ["zlib.dll", "xlsxwriter.dll"]}

extensions_args = {}
extensions_args.update(common_extension_args)
extensions_args.update(OS_specific_extension_args)

extensions_sources = [{"name": "impeller.c_workbook", "sources": workbook},
                      {"name": "impeller.c_worksheet", "sources": worksheet},
                      {"name": "impeller.c_format", "sources": formatting},
                      {"name": "impeller.c_common", "sources": common},
                      {"name": "impeller.c_chart", "sources": chart}]

extensions = [Extension(**extensions_args, **source) for source in extensions_sources]

setup(
    name="Impeller",
    description='Thin Cython wrapper around modified fork of libxlsxwriter, '
                'partially compatible with Python XlsxWriter. The purpose of the project is fast .xlsx writing.',
    author='Dmitriy Emelianov',
    author_email='emelianovds@yandex.ru',
    url="https://github.com/m9psy/Impeller",
    version='0.1.dev0',
    packages=['impeller'],
    package_data=package_data,
    ext_modules=cythonize(extensions, compiler_directives={'language_level': language_level}),
    classifiers=[
        "License :: OSI Approved :: BSD License",
        "Development Status :: 2 - Pre-Alpha",
        "Intended Audience :: Developers",
        "Operating System :: POSIX",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: C",
        "Programming Language :: Cython",
        "Programming Language :: Python",
        "Topic :: Software Development :: Libraries",
        "Topic :: Office/Business :: Financial :: Spreadsheet"
    ]
)
