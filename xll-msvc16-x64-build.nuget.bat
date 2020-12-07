REM @echo off

Echo XLL Build NuGet
set INCREMENTDISABLE=true

REM # XEON x64 Build Vars #
set _SCRIPT_DRIVE=%~d0
set _SCRIPT_FOLDER=%~dp0
set SRC=%CD%
set BUILDTREE=%SRC%\build\
SET tbs_arch=x64
SET vcvar_arg=x86_amd64
SET ms_build_suffix=Bin\amd64
SET VS16="C:\Program Files (x86)\Microsoft Visual Studio\2019\Preview\"
SET MSB16="C:\Program Files (x86)\Microsoft Visual Studio\2019\Preview\MSBuild\Current\"
SET MSBPath=%MSB16%%ms_build_suffix%
set PATH=%MSBPath%;%SRC%;%PATH%

REM # VC Vars # !!! syntax incorrect !!!
call %VS16%VC\Auxiliary\Build\vcvarsall %vcvar_arg%
@echo on

REM # Clean Build Tree #
if defined INCREMENTDISABLE ( 
	echo "Incremental Build disabled"
    rd /s /q %BUILDTREE%
) else (
	echo "Incremental Build enabled"
)
mkdir %BUILDTREE%
cd %BUILDTREE%

msbuild ..\xll.sln /p:Configuration=Release /m

:copy_files
set BINDIR=%SRC%\build-nuget\
rd /s /q %BINDIR%
mkdir %BINDIR%
echo %BINDIR%
xcopy %BUILDTREE%Release\xll* %BINDIR%
xcopy %BUILDTREE%bin\Release\xll.lib %BINDIR%
del %BINDIR%xll
xcopy /I %BUILDTREE%xll %BINDIR%xll
copy %SRC%\xll-msvc16.targets %BINDIR%\xll-msvc16-x64.targets

mkdir %BUILDTREE%
cd %BUILDTREE%

:static_XLL
REM # XLL STATIC #
if exist %BUILDTREE%\Release\xll.lib (
	ECHO XLL Libs Found
	GOTO:copy_static_files
)

msbuild xll.sln /p:Configuration=Release /m

:copy_static_files
set BINDIR=%SRC%build-nuget\
rd /s /q %BINDIR%
mkdir %BINDIR%
echo %BINDIR%
xcopy %BUILDTREE%Release\xll* %BINDIR%

:nuget_req
REM # make nuget packages from binaries #
cd %BUILDTREE%
nuget pack %SRC%\xll-msvc16-x64.nuspec
cd %SRC%
REM --- exit ----
GOTO:eof
