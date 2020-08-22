// handle.cpp - Illustrate the use of handles
//#define XLOPERX XLOPER
#include "../xll/xll.h"

using namespace xll;

class base {
	OPER x;
public:
	base() { }
	base(const OPER& x) 
		: x(x) 
	{ }
	virtual ~base()
	{ }
	OPER& get() 
	{ 
		return x; 
	}
	base& set(const OPER& _x)
	{
		x = _x;

		return *this;
	}
};

AddIn xai_base(
	Function(XLL_HANDLEX, "?xll_base", "XLL.BASE")
	.Args({
		Arg(XLL_LPOPER, "cell", "is a cell or range of cells")
	})
	.FunctionHelp("Return a handle to a base object.")
	.Uncalced() // required for functions creating handles
);
HANDLEX WINAPI xll_base(LPOPER px)
{
#pragma XLLEXPORT
	xll::handle<base> h(new base(*px));

	return h.get();
}

AddIn xai_base_get(
	Function(XLL_LPOPER, "?xll_base_get", "XLL.BASE.GET")
	.Args({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by XLL.BASE")
	})
	.FunctionHelp("Return the value stored in base.")
);
LPOPER WINAPI xll_base_get(HANDLEX _h)
{
#pragma XLLEXPORT
	xll::handle<base> h(_h);

	return h ? &h->get() : (LPOPER)&ErrNA;
}

AddIn xai_base_set(
	Function(XLL_HANDLEX, "?xll_base_set", "XLL.BASE.SET")
	.Args({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by XLL.BASE"),
		Arg(XLL_LPOPER, "cell", "is a cell or range of cells")
	})
	.FunctionHelp("Set the value of base to cell.")
);
HANDLEX WINAPI xll_base_set(HANDLEX _h, LPOPER px)
{
#pragma XLLEXPORT
	xll::handle<base> h(_h);

	if (h) {
		h->set(*px);
	}

	return _h;
}

class derived : public base {
	OPER x2;
public:
	derived(const OPER& x, const OPER& x2)
		: base(x), x2(x2)
	{ }
	OPER& get2()
	{
		return x2;
	}
};

AddIn xai_derived(
	Function(XLL_HANDLEX, "?xll_derived", "XLL.DERIVED")
	.Args({
		Arg(XLL_LPOPER, "cell", "is a cell or range of cells"),
		Arg(XLL_LPOPER, "cell2", "is a cell or range of cells")
	})
	.FunctionHelp("Return a handle to a derived object.")
	.Uncalced() // required for functions creating handles
);
HANDLEX WINAPI xll_derived(LPOPER px, LPOPER px2)
{
#pragma XLLEXPORT
	// derived isa base
	xll::handle<base> h(new derived(*px, *px2));

	return h.get();
}

// XLL.BASE.GET calls base::get for handles returned by XLL.DERIVED.
AddIn xai_derived_get(
	Function(XLL_LPOPER, "?xll_derived_get", "XLL.DERIVED.GET")
	.Args({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by XLL.DERIVED")
	})
	.FunctionHelp("Return the second value stored in derived.")
);
LPOPER WINAPI xll_derived_get(HANDLEX _h)
{
#pragma XLLEXPORT
	// get handle to base class
	xll::handle<base> h(_h);
	// cast to derived
	derived* pd = dynamic_cast<derived*>(h.ptr());

	return pd ? &pd->get2() : (LPOPER)&ErrNA;
}
