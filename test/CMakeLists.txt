add_library(test SHARED 
  handle.cpp
  #lambda.cpp
  test.cpp)

target_compile_definitions (test PUBLIC UNICODE)
target_compile_definitions (test PUBLIC _UNICODE)
target_include_directories(test PRIVATE xll)
target_link_libraries(test xll XLCALL32)
if (CMAKE_SIZEOF_VOID_P EQUAL 8)
	target_link_directories(test PRIVATE ${CMAKE_SOURCE_DIR}/xll/x64)	
else()
	target_link_directories(test PRIVATE ${CMAKE_SOURCE_DIR}/xll)		
endif()
set_target_properties(test PROPERTIES SUFFIX ".xll")
set_target_properties(test PROPERTIES VS_DEBUGGER_COMMAND "$(registry:HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe)")
set_target_properties(test PROPERTIES VS_DEBUGGER_COMMAND_ARGUMENTS "\"$(TargetPath)\" /p \"$(ProjectDir)\"")
