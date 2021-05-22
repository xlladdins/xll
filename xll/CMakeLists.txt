add_library(xll STATIC 
  auto.cpp
  debug.cpp
  defines.c
  dllmain.cpp
  error.cpp
  XLCALL.CPP
  macrofun/xll_addin.cpp
  macrofun/xll_addin_manager.h
  macrofun/xll_document.cpp
  macrofun/xll_function.cpp
  #macrofun/xll_install.cpp
  macrofun/xll_macrofun.cpp
  macrofun/xll_macrofun.h
  macrofun/xll_paste.cpp
  macrofun/xll_paste.h
  macrofun/xll_spreadsheet.cpp
  macrofun/xll_trace.cpp
  macrofun/xll_utility.cpp)

target_compile_definitions (xll PUBLIC UNICODE)
target_compile_definitions (xll PUBLIC _UNICODE)
