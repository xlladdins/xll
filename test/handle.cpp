// handle.cpp - Illustrate the use of handles.
//#define XLOPERX XLOPER
#include "../xll/xll.h"

using namespace xll;

template<class X = OPER>
class base {
	X x;
public:
	base() { }
	base(const X& x) 
		: x(x) 
	{ }
	base(const base&) = default;
	base& operator=(const base&) = default;
	virtual ~base()
	{ }
	X& get() 
	{ 
		return x; 
	}
	const X& get() const
	{
		return x;
	}
	base& set(const X& _x)
	{
		x = _x;

		return *this;
	}
};

template<class X = OPER>
class derived : public base<X> {
	X x2;
public:
	derived()
	{ }
	derived(const X& x, const X& x2)
		: base<X>(x), x2(x2)
	{ }
	derived(const derived&) = default;
	derived& operator=(const derived&) = default;
	~derived()
	{ }

	X& get2()
	{
		return x2;
	}
	const X& get2() const
	{
		return x2;
	}
};

AddIn xai_base(
	Function(XLL_HANDLEX, "xll_base", "\\XLL.BASE")
	.Arguments({
		Arg(XLL_LPOPER, "cell", "is a cell or range of cells", "{\"a\",\"range\";\"of\",\"cells\"}")
	})
	.FunctionHelp("Return a handle to a base object.")
	.Uncalced() // required for functions creating handles
);
HANDLEX WINAPI xll_base(const LPOPER px)
{
#pragma XLLEXPORT
	xll::handle<base<>> h(new base<>(*px));

	return h.get();
}

// also works for any derived class
AddIn xai_base_get(
	Function(XLL_LPOPER, "xll_base_get", "XLL.BASE.GET")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by \\XLL.BASE", "\\XLL.BASE(TRUE)")
	})
	.FunctionHelp("Return the value stored in base.")
);
LPOPER WINAPI xll_base_get(HANDLEX _h)
{
#pragma XLLEXPORT
	static OPER result;
	xll::handle<base<>> h(_h);

	if (h) {
		result = h->get();
	}
	else {
		result = ErrNA;
	}

	return &result;
}

AddIn xai_base_set(
	Function(XLL_HANDLEX, "xll_base_set", "XLL.BASE.SET")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by \\XLL.BASE"),
		Arg(XLL_LPOPER, "cell", "is a cell or range of cells")
	})
	.FunctionHelp("Set the value of base to cell.")
);
HANDLEX WINAPI xll_base_set(HANDLEX _h, LPOPER px)
{
#pragma XLLEXPORT
	xll::handle<base<>> h(_h);

	if (h) {
		h->set(*px);
	}

	return _h;
}

AddIn xai_derived(
	Function(XLL_HANDLEX, "xll_derived", "\\XLL.DERIVED")
	.Arguments({
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
	xll::handle<base<>> h(new derived<>(*px, *px2));

	return h.get();
}

// XLL.BASE.GET calls base::get for handles returned by XLL.DERIVED.
AddIn xai_derived_get(
	Function(XLL_LPOPER, "xll_derived_get", "XLL.DERIVED.GET")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by XLL.DERIVED")
	})
	.FunctionHelp("Return the second value stored in derived.")
);
LPOPER WINAPI xll_derived_get(HANDLEX _h)
{
#pragma XLLEXPORT
	static OPER result;
	// get handle to base class
	xll::handle<base<>> h(_h);
	// downcast to derived
	derived<>* pd = h.as<derived<>>();

	if (pd) {
		result = pd->get();
	}
	else {
		result = ErrNA;
	}

	return &result;
}

// convert pointers to "base[0x<hexdigits>]"
static handle<base<>>::codec<XLOPERX> base_codec(OPER("base[0x"), OPER("]"));

AddIn xai_ebase(
	Function(XLL_LPOPER, "xll_ebase", "XLL.EBASE")
	.Arguments({
		Arg(XLL_LPOPER, "cell", "is a cell or range of cells")
	})
	.FunctionHelp("Return a handle to a base object.")
	.Uncalced() // required for functions creating handles
);
LPOPER WINAPI xll_ebase(const LPOPER px)
{
#pragma XLLEXPORT
	xll::handle<base<>> h(new base<>(*px));

	return (LPOPER)&base_codec.encode(h.get());
}

AddIn xai_ebase_get(
	Function(XLL_LPOPER, "xll_ebase_get", "XLL.EBASE.GET")
	.Arguments({
		Arg(XLL_LPOPER, "handle", "is a handle returned by XLL.BASE")
	})
	.FunctionHelp("Return the value stored in base.")
);
LPOPER WINAPI xll_ebase_get(const LPOPER ph)
{
#pragma XLLEXPORT
	static OPER result;
	xll::handle<base<>> h(base_codec.decode(*ph));

	if (h) {
		result = h->get();
	}
	else {
		result = ErrNA;
	}

	return &result;
}

