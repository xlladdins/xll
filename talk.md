% Embed C++ in Excel
% Keith A. Lewis

# Rationale

C++ is an algorithmic language. You can see the code but not the data.

Excel shows the data but not the code. Think of it as a debugger on steroids.

# Embed C++ in Excel

- Call C++/C/Fortran from Excel
- Use single inheritance
- Plug in 3rd party libraries
- Integrate with Excel's help documentation
- UTF-8 all the things

# AddIn

Specify the information Excel needs to call your function.

``` C++
AddIn xai_tgamma(
    Function(XLL_DOUBLE, "xll_tgamma", "TGAMMA")
    .Arguments({
	Arg(XLL_DOUBLE, "x", "is the value for which you want to calculate Gamma.")
    })
    .FunctionHelp("Return the Gamma function value.")
    .Category("CMATH")
    .HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal")
);
```

# Function Wizard

This shows up in the function wizard like so:

<img src="images/tgamma.png" alt="Function Wizard dialog">

# Implement the function

Use `WINAPI` for functions called from Excel.

Export functions from the dll with `#pragma XLLEXPORT`


```C++
#include <cmath> // for double tgamma(double)

double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT
	return tgamma(x);
}
```

# Macros

Macros take no arguments, they only produce side effects.

```C++
AddIn xai_macro(Macro("xll_macro", "XLL.MACRO"));
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
    Excel(xlcAlert, 
        Excel(xlfConcatenate,
            OPER("XLL.MACRO 召唤 with активный  cell: "), // use utf-8!
            Excel(xlfReftext, 
                Excel(xlfActiveCell), 
                OPER(true) // A1 style instead of R1C1
            )
        ),
        OPER(2), // general information
        OPER("https://xlladdins.github.io/Excel4Macros/alert.html")
    );
	
    return TRUE;
}
```
