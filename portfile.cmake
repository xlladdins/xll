set(XLL_VERSION 1.26.5)
set(XLL_HASH 0)

#architecture detection
if(VCPKG_TARGET_ARCHITECTURE STREQUAL "x86")
   set(XLL_ARCH Win32)
   set(XLL_CONFIGURATION _x86)
elseif(VCPKG_TARGET_ARCHITECTURE STREQUAL "x64")
   set(XLL_ARCH x64)
   set(XLL_CONFIGURATION _x86)
elseif(VCPKG_TARGET_ARCHITECTURE STREQUAL "arm")
   set(XLL_ARCH ARM)
   set(XLL_CONFIGURATION _Generic)
elseif(VCPKG_TARGET_ARCHITECTURE STREQUAL "arm64")
   set(XLL_ARCH ARM64)
   set(XLL_CONFIGURATION _Generic)
else()
   message(FATAL_ERROR "unsupported architecture")
endif()

vcpkg_from_github(
    OUT_SOURCE_PATH SOURCE_PATH
    REPO xlladdins/xll
#    REF v1.3.3 - use tag to pin down
    SHA512 0
    HEAD_REF master
)

vcpkg_install_msbuild(
    SOURCE_PATH ${SOURCE_PATH}
    PROJECT_SUBPATH xll/xll.vcxproj
    OPTIONS /p:UseEnv=True
    RELEASE_CONFIGURATION Release${XLL_CONFIGURATION}${XLL_CONFIGURATION_SUFFIX}
    DEBUG_CONFIGURATION Debug${XLL_CONFIGURATION}${XLL_CONFIGURATION_SUFFIX}
)

file(INSTALL
    ${SOURCE_PATH}/xll/addin.h
    ${SOURCE_PATH}/xll/args.h
    ${SOURCE_PATH}/xll/auto.h
    ${SOURCE_PATH}/xll/codec.h
    ${SOURCE_PATH}/xll/concepts.h
    ${SOURCE_PATH}/xll/defines.h
    ${SOURCE_PATH}/xll/ensure.h
    ${SOURCE_PATH}/xll/error.h
    ${SOURCE_PATH}/xll/excel.h
    ${SOURCE_PATH}/xll/exports.h
    ${SOURCE_PATH}/xll/fp.h
    ${SOURCE_PATH}/xll/handle.h
    ${SOURCE_PATH}/xll/on.h
    ${SOURCE_PATH}/xll/oper.h
    ${SOURCE_PATH}/xll/platform.h
    ${SOURCE_PATH}/xll/register.h
    ${SOURCE_PATH}/xll/registry.h
    ${SOURCE_PATH}/xll/splitpath.h
    ${SOURCE_PATH}/xll/sref.h
    ${SOURCE_PATH}/xll/traits.h
    ${SOURCE_PATH}/xll/utf8.h
    ${SOURCE_PATH}/xll/view.h
    ${SOURCE_PATH}/xll/win.h
    ${SOURCE_PATH}/xll/xll.h
    ${SOURCE_PATH}/xll/xllio.h
    ${SOURCE_PATH}/xll/xloper.h
    ${SOURCE_PATH}/xll/macrofun/xll_addin_manager.h
    ${SOURCE_PATH}/xll/macrofun/xll_macrofun.h
    ${SOURCE_PATH}/xll/macrofun/xll_paste.h
    DESTINATION ${CURRENT_PACKAGES_DIR}/include
)

vcpkg_configure_make(
    SOURCE_PATH ${SOURCE_PATH}
    OPTIONS ${XLL_OPTIONS}
    PREFER_NINJA
)
vcpkg_install_make()
