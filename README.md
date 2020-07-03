# xll - a library for creating Excel add-ins

This library is a C++ wrapper for the Microsoft
[Excel Software Development Kit](https://docs.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit)
that makes it simple to create add-ins that register C/C++ functions 
that can be called from Excel.

## Prerequisites

Windows 10  
  The Excel SDK is not supported on MacOS.

[Visual Studio 2019](https://visualstudio.microsoft.com/)  
  Use the Community Edition and choose the `Desktop development with C++` workload.</dd>

[Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel)  
  Install the 64-bit version of Office 365 for the best experience.

[Sandcastle Help File Builder](https://github.com/EWSoftware/SHFB)  
  This is optional if you want to create help files integrated into Excel.

## Get Started

Fork the [xll](https://github.com/xlladdins/xll) repository and
[clone](https://docs.microsoft.com/en-us/visualstudio/get-started/tutorial-open-project-from-repo)
it in Visual Studio.
Run the `setup.bat` script in the top level `xll` folder.

Create a new project using `File ► New ► Project...` (`Ctrl-Shift-N`) and
select `XLL Project`. At this point you can compile and run the add-in
using `Debug ► Start Debugging` (`F5`). This compiles the dll, (with
file extension `.xll`), and starts Excel with the add-in loaded.

This is controlled by the _project properties_. Right click on a project
and select `Properties` (`Alt-Enter`) at the bottom of the popup menu.
Navigate to `Debugging` in `Configuration Properties`.

![debug properties](images/debug.png)

The `Command` `$(registry:HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe)`
looks up the registry entry for the full path of the Excel executable.
The `Command Arguments` `"$(TargetPath)" /p "$(ProjectDir)"` are passed to
Excel when the debugger starts. The variable `$(TargetPath)` is the
full path to the xll that was built and is opened by Excel. 
The `/p` flag to Excel sets the
default directory so `Ctrl-O` opens to the project directory.

## Add-in Functions

To register a C/C++ function that can be called from Excel create
an `AddIn` object that has information Excel requires.

```C++
#include <cmath>
AddIn xai_tgamma(
	Function(XLL_DOUBLE, :"xll_tgamma", "TGAMMA")
	.Args({
		Arg(XLL_DOUBLE, L"x", L"is the value for which you want to calculate Gamma.")
	})
	.FunctionHelp(L"Returns the Gamma function value.")
	.Category(L"Cmath")
);
double WINAPI xll_tgamma(double x)
{
XLLEXPORT
	return tgamma(x);
}
```

This add-in adds the function `TGAMMA` to Excel that calls the
`tgamma` function declared in the `<cmath>` library. THe
Gamma function is <math>&Gamma;(x) = &int;<sub>0</sub><sup>&infin;</sup> dx</math>