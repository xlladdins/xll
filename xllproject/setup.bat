REM setup.bat - set up git and user settings
@echo off
git init
git submodule add https://github.com/xlladdins/xll.git
copy xll\test\test.vcxproj.user $projectname$.vcxproj.user
echo Add xll/xll.vcxproj to your solution.
echo Build and debug by pressing F5.
set /p dummy=Hit Enter to continue