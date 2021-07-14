% Embed C++ in Excel
% Keith A. Lewis

# Rationale

C++ is an algorithmic language. You can see the code but not the data.

Excel is purely functional and shows the data but not the code.

C++ and Excel are complementary.

Think of it as debugging on steroids.

# Embed C++ in Excel

- Call C++/C/Fortran from Excel
- UTF-8 all the things
- Integrate with Excel's help documentation
- Plug in 3rd party libraries
- Embed objects and use single inheritance

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

```C++
#include <cmath> // for double tgamma(double)

double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT
	return tgamma(x);
}
```

Use `WINAPI` for functions called from Excel.

Export functions from the dll with `#pragma XLLEXPORT`


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
`Excel(xlfFun, args...)` is equivalent to `=FUN(args, ...)` in Excel.

`Excel(xlcMacro, args, ...)` calls `MACRO(args, ...)` They can only
be called from macros. Calling them from functions is forbidden.
You will need to consult the [Excel4Macros](https://xlladdins.github.io/Excel4Macros/)
documentation to discern the appropriate arguments.

# `OPER`

An `OPER` is an Excel cell. It can be a number, string, boolean, ...,
or a 2-dimensional range of cells.

It is a variant type that acts like a built-in type. 
`OPER o(1.23)` is the number `1.23`. 
Assigning a string, `o = "foo"`, or boolean, `o = true`, turns it into a string or boolean. 
The `OPER` `xltype` member indicates the type.
These are defined in the
[Microsoft Excel SDK header file](https://github.com/xlladdins/xll/blob/master/xll/XLCALL.H)
as `xltypeNum`, `xltypeStr`, `xltypeBool`, ..., `xltypeMulti`.

# `xll::handle`

A `xll::handle<T>` is a pointer to an object of type `T` and behaves like a `std::unique_ptr<T>`.
It refers to exactly one object and calls `delete` on the object when it goes out of scope.
Its `ptr()` member function returns the pointer to the object.
The arrow `operator->()` also returns the pointer.
Use the `get()` member function to return a `HANDLEX` value that can be used in Excel.

# `HANDLEX`

A `HANDLEX` is a double. Its bits are the same bits as the pointer used
in the constructor `xll::handle<T>(T*)`.
Like `std::unique_ptr<T>`, this is called using `xll::handle<T> h(new T(args...))`.
Converting a `HANDLEX` to a pointer, and back, is just a cast. 
It only takes a few machine instructions instead of a lookup in an associative array.

You may, and should, be worried that the 64-bits of a pointer correspond to
a NaN or denormal IEEE 64-bit floating point number. Those do not survive a
round trip to Excel and back. Excel converts those to `#NUM!`.
On 64-bit Windows 10 the first 16-bits of a pointer are zero so we only need the remaining 48-bits.
Doubles can exactly represent integers up to 2<sup>53</sup> so we have plenty of room to spare.

# Calling member functions

The constructor `xll::handle<T>(HANDLEX)` converts a `HANDLEX` to a pointer.
It also checks if the `HANDLEX` was created by a prior call to `xll::handle<T>(T*)`.
If not, the pointer is set to `nullptr` and `explicit xll::handle<T>::operator bool() const`
will return `false`. If so, use `.ptr()` or `operator->()` to call member functions.

# Single inheritence

When creating a handle to an object `U` that is derived from `T`
use `xll::handle<T> h(new U(args...))` and return the `HANDLEX` with `h.get()`.
The handle can be used to call any member function of `T`.
To call member functions of `U` use `dynamic_cast`.
The convenience function `xll::handle<T>.as<U>()` also does this.

# Calling `delete`

Excel add-ins dealing with C++ objects typically have an object manager of some sort.
Excel knows nothing about when `new` is called to create an object so the
manager tries to keep track of objects that are being used and do garbage collection.

The simplest C++ garbage collector is `}`.  The `xll::handle` class will call
`delete` on objects created by `xll::handle<T>(T*)` when called from a cell
containing a `HANDLEX` from a previous call. The constructor is being provided
with new arguments so the old object is no longer required.

This can leak if the cell containing the handle is deleted. Moving and copying...

