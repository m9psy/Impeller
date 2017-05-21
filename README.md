# Impeller

Impeller is a rotor used to increase (or decrease in case of turbines) the pressure and flow of a fluid.

It is also a thin Cython wrapper around modified [fork](https://github.com/m9psy/libxlsxwriter) of [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter), partially compatible with [Python XlsxWriter](https://github.com/jmcnamara/XlsxWriter).

<table>
  <tr>
    <td rowspan="2">
    Build status
    </td>
    <td>
      <a title="Travis CI build status" href="https://travis-ci.org/m9psy/Impeller.svg?branch=master">
        <img src="https://travis-ci.org/m9psy/Impeller.svg?branch=master">
      </a>
    </td>
  </tr>
  <tr>
    <td>
      <a title="Appveyor build status" href="https://ci.appveyor.com/project/m9psy/impeller">
        <img src="https://ci.appveyor.com/api/projects/status/7eclq5f5qbxpsqyv?svg=true">
      </a>
    </td>
  </tr>
</table>

Since the only reason to create this package is _speed obsession_, here are some crude results (work in progress and current problem is API compatibility). Results was received by running `misc/simple_bench.py` on Windows machine. Benchmark saving 10000 * 10 cells with extensive styling (every even row styled).

<table>
<head>
  <tr>
    <th>Lib name</th>
    <th title="Close time - time to create archive">Timing in seconds</th>
    <th>Times slower than Impeller</th>
  </tr>
</head>
<tbody>
<tr>
  <td>Impeller</td>
  <td>Close time: 0.29 Total: 0.43</td>
  <td>1x</td>
</tr>
<tr>
  <td>Python XlsxWriter</td>
  <td>Close time: 1.32 Total: 2.00</td>
  <td>4.7x</td>
</tr>
<tr>
  <td>xlwt</td>
  <td>Close time: 0.34 Total: 1.30</td>
  <td>3.0x</td>
</tr>
<tr>
  <td>pyexcelerate</td>
  <td>Close time: 1.88 Total: 2.17</td>
  <td>5.0x</td>
</tr>
<tr>
  <td>openpyxl</td>
  <td>Close time: 2.21 Total: 7.56</td>
  <td>17.6x</td>
</tr>
</tbody>
</table>

Some comments: `xlwt` produces `.xls` files, not `.xlsx` and unable to handle more than 65536 cells (`.xls` format limitation and xlwt is not participating in second competition). `openpyxl` can actually not only write, but read and edit files (the only one who builds the elements tree, I suppose). openpyxl however was tested in `write_only` mode. XlsxWriter and Impeller are sharing the same API.

It is possible to run Impeller faster: instead of all-in-one `write` method you can use type-stricted methods like:
`write_number_strict`
`write_string_strict`
`write_blank_strict`
etc.
This methods are much faster, than simple `write` (so you can dump an array faster), but requires correct third parameter type. Also you can not use Excel notation to provide row and col ("A2", "E15" not working) in fast methods - you can convert this notation to indexes via `xl_cell_to_rowcol` function.

Another table for 100000 * 10 cells:

<table>
<head>
  <tr>
    <th>Lib name</th>
    <th title="Close time - time to create archive">Timing in seconds</th>
    <th>Times slower than Impeller</th>
  </tr>
</head>
<tbody>
<tr>
  <td>Impeller</td>
  <td>Close time: 3.23 Total: 4.69</td>
  <td>1x</td>
</tr>
<tr>
  <td>Python XlsxWriter</td>
  <td>Close time: 10.61 Total: 18.17</td>
  <td>3.9x</td>
</tr>
<tr>
  <td>pyexcelerate</td>
  <td>Close time: 17.48 Total: 20.74</td>
  <td>4.4x</td>
</tr>
<tr>
  <td>openpyxl</td>
  <td>Close time: 21.22 Total: 69.00</td>
  <td>14.7x</td>
</tr>
</tbody>
</table>


TODO: benchmark.c, comparing Cython, C and modified C versions

## Current status:
  Unstable, not usable at all, work in progress.

TODO: Move to docs.

TODO: Other compilers

2.7, 3.5 Python supported.

How to build from source:

Since there are submodules, you will need to clone them too:
`git clone --recursive https://github.com/m9psy/Impeller`

  ## Win cmake:

  1. You will need to build zlib first. The build process is usual, CmakeLists for zlib was not changed. The only difference is different install path:
  ```
  cd Impeller\zlib
  mkdir msvc_solution64
  cd msvc_solution64
  cmake .. -G"Visual Studio 14 2015 Win64" -DCMAKE_INSTALL_PREFIX="..\..\libs\zlib" -DFAST_LMATCH=ON
  cmake --build . --config Release --target INSTALL
  ```
  Alternatively it is possible to find precompiled (or use modified one) zlib. By default zlib will be compiled with `__cdecl` convention. For `__stdcall` you can find ready-to-use solutions in contrib/vc folder.
  
  2. You will need to build libxlsxwirter and tell where zlib should be found.
  ```
  cd Impeller\libxlsxwriter
  mkdir msvc_solution64
  cd msvc_solution64
  cmake .. -G"Visual Studio 14 2015 Win64" -DBUILD_STATIC=OFF -DZLIB_ROOT=..\..\libs\zlib -DZLIB_INCLUDE_DIR=..\..\libs\zlib\include -DZLIB_LIBRARY=..\..\libs\zlib\lib\zlib.lib -DCMAKE_INSTALL_PREFIX=..\..\libs\XLSX
  cmake --build . --config Release --target INSTALL
  ```
  On Windows pointing to correct zlib location is somewhat tricky - you will see `Performing Test ZLIB_COMPILES - Success` message if everything is OK.
  Both commands can be called via `build_c_libs.cmd`
  
  3. Cython is required for the next step. wheel is required - (`pip install -r dev_requirements.txt`). You will need to build the extension:
  ```
  cd Impeller
  python setup.py build_ext bdist_wheel
  ```
  As a result there will be .whl file in the dist directory - you can install it with `pip install <wheel_name>.whl` command.
  
## Linux:
  1. zlib is required. Use your favorite package manager like `sudo apt-get install zlib1g-dev`.
  2. build-essential is requried.
  3. checkinstall is required.
  ```
  cd libxlsxwriter
  make
  sudo checkinstall make install
  cd ..
  python setup.py build_ext bdist_wheel
  ```
  As a result there will be .whl file in the dist directory - you can install it with `pip install <wheel_name>.whl` command.
  
  
