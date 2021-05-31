cmake_minimum_required(VERSION 3.10)
project(xll CXX C)

option(BUILD_EXAMPLES "Build examples" OFF)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_FLAGS_RELEASE "/MT /EHsc")
set(CMAKE_CXX_FLAGS_DEBUG "/MTd /EHsc")
set(CMAKE_CXX_FLAGS "/J /utf-8 /Wv:18")

add_subdirectory(xll)
if(BUILD_EXAMPLES)
  add_subdirectory(xll.t)
  add_subdirectory(test)
endif()
