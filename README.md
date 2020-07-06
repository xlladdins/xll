# xll - a library for creating Excel add-ins

This library makes it simple call C and C++ functions from Excel.
It is much easier to use than the Microsoft
[Excel Software Development Kit](https://docs.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit).

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

The program to start and the arguments to use are specified in _project properties_.
Right click on a project and select `Properties` (`Alt-Enter`) at the bottom of the 
popup menu.
Navigate to `Debugging` in `Configuration Properties`.

![debug properties](images/debug.png)

The `Command` `$(registry:HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe)`
looks in the registry for the full path of the Excel executable.
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
	Function(XLL_DOUBLE, L"xll_tgamma", L"TGAMMA")
	.Args({
		Arg(XLL_DOUBLE, L"x", L"is the value for which you want to calculate Gamma.")
	})
	.FunctionHelp(L"Return the Gamma function value.")
	.Category(L"Cmath")
);
double WINAPI xll_tgamma(double x)
{
XLLEXPORT
	return tgamma(x);
}
```

This add-in registers the function `TGAMMA` with Excel to call the C++ function
`xll_tgamma` that returns a floating point `double`. It has one argument that
is also a `double` and is added to the Excel function wizard under the
`Cmath` category with the specified function help. Compare this to
the built-in Excel functon [`GAMMA`](https://support.microsoft.com/en-us/office/gamma-function-ce1702b1-cf55-471d-8307-f83be0fc5297).

The function `xll_gamma` calls the `tgamma` function` declared in 
the `<cmath>` library. All functions called from Excel must be declared
with [WINAPI](https://docs.microsoft.com/en-us/cpp/cpp/stdcall?view=vs-2019).
The macro `XLLEXPORT` causes the function to be exported
from the dll so it will be visible to Excel.

Recall the Gamma function is 
<math>&Gamma;(x) = &int;<sub>0</sub><sup>&infin;</sup> 
t<sup>x - 1</sup> e<sup>-t</sup>&nbsp;dt</math>. 
It satisfies <math>&Gamma;(x + 1) = x &Gamma;(x)</math>
for <math>x &ge; 0</math>. Since <math>&Gamma;(1) = 1</math> we have
<math>&Gamma(x + 1) = x!</math>
if <math>x</math> is a non-negative integer.