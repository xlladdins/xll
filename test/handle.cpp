// handle.cpp - Embed C++ in Excel using xll::handle<T>
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
	// put x in base and x2 n derived
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
};

AddIn xai_base(
	Function(XLL_HANDLEX, "xll_base", "\\XLL.BASE")
	.Arguments({
		Arg(XLL_LPOPER, "cell", "is a cell or range of cells")
	})
	.Uncalced() // required for functions creating handles
	.FunctionHelp("Return a handle to a base object.")
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
		Arg(XLL_HANDLEX, "handle", "is a handle returned by \\XLL.BASE")
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
		Arg(XLL_HANDLEX, "handle", "is a handle returned by \\XLL.BASE."),
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
	.Uncalced() // required for functions creating handles
	.FunctionHelp("Return a handle to a derived object.")
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
	Function(XLL_LPOPER, "xll_derived_get", "XLL.DERIVED.GET2")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by \\XLL.DERIVED.")
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
	auto pd = h.as<derived<>>();

	if (pd) {
		result = pd->get2();
	}
	else {
		result = ErrNA;
	}

	return &result;
}

// convert pointers to "\\base[0x<hexdigits>]"
static handle<base<>>::codec<XLOPERX> base_codec(OPER("\\base[0x"), OPER("]"));

AddIn xai_ebase(
	Function(XLL_LPOPER, "xll_ebase", "\\XLL.EBASE")
	.Arguments({
		Arg(XLL_LPOPER, "cell", "is a cell or range of cells")
	})
	.Uncalced()
	.FunctionHelp("Return a handle to a base object.")
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

#ifdef _DEBUG

AddIn xai_handle_test(Macro("xll_handle_test", "HANDLE.TEST"));
int WINAPI xll_handle_test()
{
#pragma XLLEXPORT
	try {

		// enter data in cell RiCj
		// https://xlladdins.github.io/Excel4Macros/formula.html
		auto Formula = [](unsigned i, unsigned j, auto data) {
			ensure (Excel(xlcFormula, OPER(data), OPER(REF(i,j))));
		};

		// get contents of cell RiCj;
		// https://xlladdins.github.io/Excel4Macros/get.cell.html
		auto Contents = [](unsigned i, unsigned j) {
			return Excel(xlfGetCell, OPER(5), OPER(REF(i, j)));
		};

		// https://xlladdins.github.io/Excel4Macros/new.html
		Excel(xlcNew, OPER(1), Missing, Missing); // worksheet

		// test base
		Formula(0, 0, 1.23);
		Formula(1, 0, "=\\XLL.BASE(R[-1]C[0])");
		Formula(2, 0, "=XLL.BASE.GET(R[-1]C[0])");
		
		ensure(Contents(2, 0) == 1.23);
		
		Formula(0, 0, "foo");
		ensure(Contents(2, 0) == "foo");

		// test derived
		Formula(4, 0, 4.56);
		Formula(5, 0, "bar");
		Formula(6, 0, "=\\XLL.DERIVED(R[-2]C[0], R[-1]C[0])");
		Formula(7, 0, "=XLL.BASE.GET(R[-1]C[0])");
		Formula(8, 0, "=XLL.DERIVED.GET2(R[-2]C[0])");

		ensure(Contents(7, 0) == 4.56); // derived isa base
		ensure(Contents(8, 0) == "bar");

		Formula(4, 0, "=NA()");
		ensure(Contents(7, 0) == ErrNA);
		ensure(Contents(8, 0) == "bar");
		Formula(5, 0, "baz");
		ensure(Contents(8, 0) == "baz");

		// use pretty handles
		auto R1C1 = Contents(0, 0);
		Formula(10, 0, "=\\XLL.EBASE(R1C1)");
		auto base = Contents(10, 0); // pretty name
		Formula(11, 0, "=XLL.EBASE.GET(R[-1]C[0])");
		ensure(Contents(11, 0) == R1C1);
		ensure(Contents(2, 0) = R1C1); // XLL.BASE.GET also called

		ensure(Excel(xlfLeft, base, OPER(8)) == "\\base[0x"); // check prefix
		Formula(0, 0, true);
		ensure(Contents(11, 0) == true);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}
Auto<OpenAfter> xaoa_handle_test(xll_handle_test);

#endif // _DEBUG
