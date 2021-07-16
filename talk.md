% Embed C++ in Excel
% Keith A. Lewis

# Rationale

C++ is an algorithmic language. You see the code but not the data.

Excel is purely functional. You see the data but not the code.

C++ and Excel are complementary.

Embedding C++ in Excel gives you a debugger on steroids.

# [xlladdins.com](https://xlladdins.com)

- Call C++/C/Fortran from Excel
- Use UTF-8 strings
- Integrate with Excel help documentation
- Embed objects and use single inheritance
- Plug in 3rd party libraries

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

This shows up in the `CMATH` category of the function wizard as:

<img src="images/tgamma.png" alt="Function Wizard dialog">

Click [Help on this function](https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/tgamma-tgammaf-tgammal)
to open the URL in a browser.

Implement `xll_tgamma`

```C++
// xll_tgamma.cpp - call <cmath> tgamma
#include <cmath>

double WINAPI xll_tgamma(double x)
{
#pragma XLLEXPORT
	return tgamma(x);
}
```

Use `WINAPI` for functions called by Excel if you don't like debugging mysterious corrupt call stack crashes.

Export functions from the dll with `#pragma XLLEXPORT`. If you forget you will get a warning when the add-in is opened.

# Macros

Macros take no arguments, they only produce side effects.

```C++
AddIn xai_macro(Macro("xll_macro", "XLL.MACRO"));
int WINAPI xll_macro(void)
{
#pragma XLLEXPORT
    Excel(xlcAlert, 
        Excel(xlfConcatenate,
            OPER("XLL.MACRO 召唤 with активный  cell: "), // use utf-8
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

`Excel(xlcMacro, args, ...)` calls `MACRO(args, ...)`. This can only
be called from macros, not functions.
Consult the [Excel4Macros](https://xlladdins.github.io/Excel4Macros/)
documentation to discern the appropriate arguments.

# `OPER`

An `OPER` is an Excel cell. It can be a number, string, boolean, ...,
or a 2-dimensional range of cells.

It is a variant that acts like a built-in type. 
`OPER o(1.23)` is the number `1.23`. 
Assigning a string, `o = "foo"`, or boolean, `o = true`, turns it into a string or boolean. 
The `xltype` member of `OPER` indicates the type.
These are defined in the
[Microsoft Excel SDK header file](https://github.com/xlladdins/xll/blob/master/xll/XLCALL.H)
as `xltypeNum`, `xltypeStr`, `xltypeBool`, ..., `xltypeMulti`.

# `xll::handle`

A `xll::handle<T>` has a pointer to an object of type `T` and behaves like `std::unique_ptr<T>`.
The constructor `xll::handle<T> h(new T(args...))` stores the pointer returned by `new`.
It refers to exactly one object and calls `delete` on the object when it goes out of scope.
Its `ptr()` and arrow `operator->()` member function return the pointer to the object.
Use the `get()` member function to return a `HANDLEX` value that can be used in Excel.

# `HANDLEX`

A `HANDLEX` is a double. Its bits are the same bits as the pointer.
Converting a `HANDLEX` to a pointer and back is a cast
that only takes a few machine instructions instead of a lookup in an associative array.
On 64-bit Windows 10 the first 16-bits of a pointer are zero so we only need the remaining 48-bits.
Doubles can exactly represent integers less than 2<sup>53</sup> so there is plenty of room to spare.

The constructor `xll::handle<T>(HANDLEX)` converts a `HANDLEX` to a pointer.
It also checks if the `HANDLEX` was created by a prior call to `xll::handle<T>(T*)`.
Use `explicit xll::handle<T>::operator bool() const` to detect that.

# Example

```C++
// base.h
#include <concepts>

template<std::semiregular T>
class base {
	T t;
public:
	base(const T& t) 
		: t(t) 
	{ }
	virtual ~base() 
	{ }
	T get() const
	{
		return t; 
	}
};
```
```C++
// xll_base.cpp
#include "base.h"
#include "xll/xll.h"

using namespace xll;

AddIn xai_base(
	Function(XLL_HANDLEX, "xll_base", "\\XLL.BASE")
	.Arguments({
		Arg(XLL_LPOPER, "cell", "is a cell."))
	})
	.Uncalced() // \XLL.BASE has a side effect
);
HANDLEX WINAPI xll_base(LPOPER po)
{
#pragma XLLEXPORT
	handle<base<OPER>> h(new base(*po));
	
	return h.get();
}

AddIn xai_base_get(
	Function(XLL_LPOPER, "xll_base_get", "XLL.BASE.GET")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle to a base<OPER> object."))
	})
);
LPOPER WINAPI xll_base_get(HANDLEX h)
{
#pragma XLLEXPORT
	handle<base<OPER>> h_(h);
	
	return h_ ? h_->get() : (LPOPER)ErrValue;
}
```

Note `h.get()` returns the `HANDLEX` that `h_->get()` uses to call the member function.
If the handle did not come from a previous call to `\XLL.BASE` then `#VALUE!`
is returned.

# Single inheritence

When creating a handle to an object `U` that is derived from `T`
use `xll::handle<T> h(new U(args...))` and return the `HANDLEX` with `h.get()`.
The handle can be used to call any member function of `T`.
To call member functions of `U` use `dynamic_cast`.
The convenience function `xll::handle<T>::as<U>()` does the dynamic cast for you.

```C++
// derived.h
#include "base.h"

template<std::semiregular T>
class derived : public base<T> {
	T t2;
public:
	derived(const T& t, const T& t2) 
		: base<T>(t), t2(t2)
	{ }
	~derived() 
	{ }
	T get2() const
	{
		return t2; 
	}
};
```
```C++
// xll_derived.cpp
#include "derived.h"
#include "xll/xll.h"

using namespace xll;

AddIn xai_derived(
	Function(XLL_LPOPER, "xll_derived", "\\XLL.DERIVED")
	.Arguments({
		Arg(XLL_LPOPER, "cell", "is a cell passed to base<OPER>. ))
		Arg(XLL_LPOPER, "cell2", "is a cell passed to derived<OPER>."))
	})
	.Uncalced() // \XLL.DERIVED has a side effect
);
LPOPER WINAPI xll_derived(LPOPER po, LPOPER po2)
{
#pragma XLLEXPORT
	handle<base<OPER>> h(new derived<OPER>(*po, *po2));
	
	return h.get();
}

AddIn xai_derived_get2(
	Function(XLL_LPOPER, "xll_derived_get2, "XLL.DERIVED.GET2")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle to a derived<OPER> object."))
	})
);
LPOPER WINAPI xll_derived_get2(HANDLEX h)
{
#pragma XLLEXPORT
	handle<base<OPER>> h_(h);
	auto h2 = h_.as<derived<OPER>>();
	
	return h2 ? h2->get2() : (LPOPER)ErrValue;
}
```

# Calling `delete`

Excel add-ins dealing with C++ objects typically have an "object manager."
Excel knows nothing about when `new` is called to create an object so the
manager tries to keep track of objects that are being used and do garbage collection.

The simplest C++ garbage collector is `}`.  The `xll::handle` class will call
`delete` on objects created by `xll::handle<T>(T*)` when called from a cell
containing a `HANDLEX` from a previous call. The constructor is being provided
with new arguments so the old object is no longer required.

This can leak if the cell containing the handle is deleted but not
when the cell is copied or moved.

## Remarks

It only takes a couple of lines of code to call or embed C++ in Excel.

Once you do this all of the functionality of Excel is available to you.
You can copy and paste to your heart's content and insert graphs
to get a better picture of data being returned by your code.
