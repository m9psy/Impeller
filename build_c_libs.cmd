cd zlib
mkdir msvc_solution64
cd msvc_solution64
cmake .. -G"Visual Studio 14 2015 Win64" -DCMAKE_INSTALL_PREFIX="..\..\libs\zlib" -DFAST_LMATCH=ON
cmake --build . --config Release --target INSTALL

cd ..\..\libxlsxwriter
mkdir msvc_solution64
cd msvc_solution64
cmake .. -G"Visual Studio 14 2015 Win64" -DBUILD_STATIC=OFF -DZLIB_ROOT=..\..\libs\zlib -DZLIB_INCLUDE_DIR=..\..\libs\zlib\include -DZLIB_LIBRARY=..\..\libs\zlib\lib\zlib.lib -DCMAKE_INSTALL_PREFIX=..\..\libs\XLSX
cmake --build . --config Release --target INSTALL

cd ..\..