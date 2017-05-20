# Impeller

Impeller is a rotor used to increase (or decrease in case of turbines) the pressure and flow of a fluid.

It is also a thin Cython wrapper around modified [fork](https://github.com/m9psy/libxlsxwriter) of [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter), partially compatible with [Python XlsxWriter](https://github.com/jmcnamara/XlsxWriter).

<table>
  <tr>
    <td>
    Build status
    </td>
    <td>
      <a title="Appveyor build status" href="https://ci.appveyor.com/project/m9psy/impeller">
        <img src="https://ci.appveyor.com/api/projects/status/7eclq5f5qbxpsqyv?svg=true">
      </a>
    </td>
  </tr>
</table>

TODO: Link to benchmark.py, benchmark.c, comparing Cython, C and modified C versions
The only difference is speed:
TODO: Benchmarks here

## Current status:
  Unstable, not usable at all, work in progress.

TODO: Move to docs.

TODO: Other compilers

TODO: build.bat


How to build from source:

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
  
  3. Cython is required for the next step. wheell is required (`pip install wheel`). You will need to build the extension:
  ```
  cd Impeller
  python setup.py build_ext bdist_wheel
  ```
  As a result there will be .whl file in the dist directory
  
