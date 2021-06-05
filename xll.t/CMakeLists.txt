add_executable(xll.t 
  ../xll/debug.cpp
  ../xll/defines.c
  oper.t.cpp
  registry.t.cpp
  sref.t.cpp
  utf8.t.cpp
  xll.t.cpp)

target_compile_definitions (xll.t PUBLIC UNICODE)
target_compile_definitions (xll.t PUBLIC _UNICODE)
target_include_directories(xll.t PRIVATE xll)
if (CMAKE_SIZEOF_VOID_P EQUAL 8)
	target_link_directories(xll.t PRIVATE ${CMAKE_SOURCE_DIR}/xll/x64)	
else()
	target_link_directories(xll.t PRIVATE ${CMAKE_SOURCE_DIR}/xll)		
endif()
target_link_libraries(xll.t xll XLCALL32)
