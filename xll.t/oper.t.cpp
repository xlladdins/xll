// oper.t.cpp - test OPER type
#undef NDEBUG
#include <cassert>
#include <functional>
#define XLOPERX XLOPER
#include "../xll/xll.h"

using namespace xll;

int test_defines = []()
{
	return 0;
}();

int test_oper_adt = []()
{
	{
		OPER4 o("abc");
		OPER4 o2(o);
		o = o2;
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 3);
		ensure(o.val.str[3] == 'c');
	}
	{
		const char* abc = "abc";
		OPER4 o(abc);
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 3);
		ensure(o.val.str[3] == 'c');
		const char* de = "de";
		o = de;
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 2);
		ensure(o.val.str[2] == 'e');
	}
	{
		OPER12 o(L"abc");
		OPER12 o2(o);
		o = o2;
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 3);
		ensure(o.val.str[3] == 'c');
	}
	{
		const wchar_t* abc = L"abc";
		OPER12 o(abc);
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 3);
		ensure(o.val.str[3] == 'c');
		const wchar_t* de = L"de";
		o = de;
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 2);
		ensure(o.val.str[2] == 'e');
	}
	{
		const char* s = nullptr;
		OPER o(s);
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 0);
	}

	return 0;
}
();

int test_oper_default = []()
{
	{
		OPER4 o;
		ensure(o.xltype == xltypeNil);
	}
	{
		OPER12 o;
		ensure(o.xltype == xltypeNil);
	}

	return 0;
}
();

int test_oper_num = []()
{
	{
		OPER4 o(1.23);
		ensure(o.xltype == xltypeNum);
		ensure(o.val.num == 1.23);
		ensure(o == 1.23);
		o = 2.34;
		ensure(o == 2.34);
	}
	{
		OPER12 o(1.23);
		ensure(o.xltype == xltypeNum);
		ensure(o.val.num == 1.23);
	}
	{
		OPER4 o;
		o = 1.23;
		ensure(o.xltype == xltypeNum);
		ensure(o.val.num == 1.23);
	}
	{
		OPER12 o;
		o = 1.23;
		ensure(o.xltype == xltypeNum);
		ensure(o.val.num == 1.23);
	}
	{
		OPER4 o;
		o = 1;
		ensure(o.xltype == xltypeInt);
		ensure(o.val.w == 1);
		ensure(o == 1);
	}
	{
		OPER4 o;
		o = 2.;
		ensure(o.xltype == xltypeNum);
		ensure(o.val.num == 2);
		ensure(o.val.w != 2);
		ensure(o.val.xbool != 2);
		ensure(o == 2.);
		ensure(o == 2);
	}
	/* !!! operator double() not working
	{
		OPER x(1.2), y(2.3);
		ensure(x);
		double z = x + y;
		ensure(z == 1.2 + 2.3);
		x = true;
		ensure(x.xltype == xltypeBool);
		z = x + y;
		ensure(z == 1 + 2.3);
		y = 4;
		ensure(y.xltype == xltypeInt);
		z = x + y;
		ensure(z == 1 + 4);
	}
	*/

	return 0;
}
();

int test_oper_str = []()
{
	{
		OPER4 o("abc");
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == (char)strlen("abc"));
		ensure(0 == strncmp("abc", o.val.str + 1, o.val.str[0]));
		ensure(o == "abc");
		o = "def";
		ensure(o == "def");
		o &= "abc";
		ensure(o == "defabc");
	}
	{
		const char* abc = "abc";
		OPER4 o(abc);
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == (char)strlen("abc"));
		ensure(0 == strncmp("abc", o.val.str + 1, o.val.str[0]));
	}
	{
		auto abc = "abc";
		OPER4 o(abc);
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == (char)strlen("abc"));
		ensure(0 == strncmp("abc", o.val.str + 1, o.val.str[0]));
	}
	{
		OPER12 o(L"abc");
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == (wchar_t)wcslen(L"abc"));
		ensure(0 == wcsncmp(L"abc", o.val.str + 1, o.val.str[0]));
	}
	{
		const wchar_t* abc = L"abc";
		OPER12 o(abc);
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == (char)wcslen(abc));
		ensure(0 == wcsncmp(abc, o.val.str + 1, o.val.str[0]));
	}
	{
		auto abc = L"abc";
		OPER12 o(abc);
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == (char)wcslen(L"abc"));
		ensure(0 == wcsncmp(L"abc", o.val.str + 1, o.val.str[0]));
	}
	{
		OPER4 o;
		o = "abc";
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == (char)strlen("abc"));
		ensure(0 == strncmp("abc", o.val.str + 1, o.val.str[0]));
	}
	{
		OPER12 o;
		auto abc = L"abc";
		o = abc;
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == (wchar_t)wcslen(abc));
		ensure(0 == wcsncmp(abc, o.val.str + 1, o.val.str[0]));
	}
	{
		OPER4 o;
		o.append("abc");
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == (char)strlen("abc"));
		ensure(0 == strncmp("abc", o.val.str + 1, o.val.str[0]));
	}
	{
		OPER12 o;
		const XCHAR abc[] = L"abc";
		o.append(abc);
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == (wchar_t)wcslen(abc));
		ensure(0 == wcsncmp(abc, o.val.str + 1, o.val.str[0]));
	}
	{
		const char* abc = "\03abc";
		const XLOPER o2 = { .val = { .str = (LPSTR)abc }, .xltype = xltypeStr };
		OPER4 o;
		o.append(o2);
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == (char)strlen("abc"));
		ensure(0 == strncmp("abc", o.val.str + 1, o.val.str[0]));
	}

	return 0;
}
();

int test_oper_multi = []()
{
	{
		OPER m(2, 3);
		ensure(m.xltype == xltypeMulti);
		ensure(m.rows() == 2);
		ensure(m.columns() == 3);
		ensure(m.size() == 6);
		ensure(m[1] == OPER{});

		for (auto i : m) {
			ensure(i.xltype == xltypeNil);
			ensure(i == Nil);
		}
		for (auto& i : m) {
			ensure(i.xltype == xltypeNil);
			ensure(i == Nil);
		}
		for (const auto& i : m) {
			ensure(i.xltype == xltypeNil);
			ensure(i == Nil);
		}

		m(1,0) = "foo";
		ensure(m(1,0).xltype == xltypeStr);
		ensure(m(1,0) == "foo");
		ensure(m[3] == "foo");

		m.resize(3, 2);
		ensure(m.xltype == xltypeMulti);
		ensure(m.rows() == 3);
		ensure(m.columns() == 2);
		ensure(m.size() == 6);
		ensure(m[3] == "foo");
		ensure(m(1, 1) == "foo");

		m.resize(2, 2);
		ensure(m.xltype == xltypeMulti);
		ensure(m.rows() == 2);
		ensure(m.columns() == 2);
		ensure(m.size() == 4);
		ensure(m[3] == "foo");
		ensure(m(1, 1) == "foo");

		m.resize(3, 3);
		ensure(m.xltype == xltypeMulti);
		ensure(m.rows() == 3);
		ensure(m.columns() == 3);
		ensure(m.size() == 9);
		ensure(m[3] == "foo");
		ensure(m(1, 0) == "foo");

		m(1, 2) = OPER("bar");
		ensure(m(1, 2) == "bar");

		m(2, 1) = m; // multis can nest
		ensure(m(2, 1)(1, 2) == "bar");
	}
	{
		OPER o({});
		ensure(o.xltype == xltypeNil);
		ensure(o.size() == 0);
		o.append_bottom(o);
		ensure(o.xltype == xltypeNil);
	}
	{
		OPER o({ OPER(1.23) } );
		ensure(o.xltype == xltypeNum);
		ensure(o.rows() == 1);
		ensure(o.columns() == 1);
		ensure(o[0] == 1.23);
		ensure(o == 1.23);
		o.append_bottom(o);
		ensure(o.rows() == 2);
		ensure(o.columns() == 1);
		ensure(o[0] == 1.23);
		ensure(o[1] == 1.23);
	}
	{
		OPER o{ { OPER(1.23) } };
		ensure(o.xltype == xltypeMulti);
		ensure(o.rows() == 1);
		ensure(o.columns() == 1);
		ensure(o[0] == 1.23);
	}
	{
		OPER o({ OPER(1.23), OPER("foo") } );
		ensure(o.xltype == xltypeMulti);
		ensure(o.rows() == 1);
		ensure(o.columns() == 2);
		ensure(o[0] == 1.23);
		ensure(o[1] == "foo");
		o.append_bottom(o);
		ensure(o.rows() == 2);
		ensure(o.columns() == 2);
		ensure(o[0] == 1.23);
		ensure(o[1] == "foo");
		ensure(o[2] == 1.23);
		ensure(o[3] == "foo");
	}
	{
		OPER o({ OPER(1.23), OPER("foo") });
		ensure(o.xltype == xltypeMulti);
		ensure(o.rows() == 1);
		ensure(o.columns() == 2);
		ensure(o[0] == 1.23);
		ensure(o[1] == "foo");
		XLOPERX x;
		XLOPERX x2[2];
		x2[0].xltype = xltypeMissing;
		x2[1].xltype = xltypeErr;
		x.xltype = xltypeMulti;
		x.val.array.rows = 1;
		x.val.array.columns = 2;
		x.val.array.lparray = x2;
		o.append_bottom(x);
		ensure(o.rows() == 2);
		ensure(o.columns() == 2);
		ensure(o[0] == 1.23);
		ensure(o[1] == "foo");
		ensure(o[2].xltype == xltypeMissing);
		ensure(o[3].xltype == xltypeErr);
	}

	return 0;
}
();

int test_oper_bool = []()
{
	{
		OPER o(true);
		ensure(o.xltype == xltypeBool);
		ensure(o.val.xbool == TRUE);
		ensure(o == true);
		o = false;
		ensure(o == false);
	}
	{
		OPER12 o(false);
		ensure(o.xltype == xltypeBool);
		ensure(o.val.xbool == FALSE);
	}
	{
		OPER o;
		ensure(o.xltype == xltypeNil);
		o = true;
		ensure(o.xltype == xltypeBool);
		ensure(o.val.xbool == TRUE);
		o = false;
		ensure(o.xltype == xltypeBool);
		ensure(o.val.xbool == FALSE);
	}

	return 0;
}
();

int test_oper_int = []()
{
	{
		OPER o(123);
		ensure(o.xltype == xltypeInt);
		ensure(o.val.w == 123);
	}
	{
		OPER12 o(-123);
		ensure(o.xltype == xltypeInt);
		ensure(o.val.w == -123);
	}

	return 0;
}
();

int test_compare = []()
{
	{
		OPER o1(1.23);
		XLOPERX o1_ = { .val = { .num = 2.34 }, .xltype = xltypeNum };
		ensure(o1 == o1);
		ensure(o1_ == o1_);
		ensure(o1 < o1_);
		ensure(o1 <= o1_);
		ensure(o1_ > o1);
		ensure(o1_ >= o1);
	}
	{
		OPER o1(1.23), o1_(2.34);
		ensure(o1 == o1);
		ensure(o1 < o1_);
		//ensure(o1 != o1_);
		ensure(o1 <= o1_);
		ensure(o1_ > o1);
		ensure(o1_ >= o1);
	}
	{
		OPER s1("abc"), s1_("def");
		ensure(s1 < s1_);
		ensure(s1 == s1);
		ensure(s1 < s1_);
		ensure(s1 <= s1_);
		ensure(!(s1 == s1_));
	}

	return 0;
}
();

int test_fp = []()
{
	using xll::FP4;
	using xll::FP12;
	using xll::FPX;

	{
		_FPX a = { 1, 1 };
		ensure(size(a) == 1);
		a.array[0] = 1.23;
		FPX ax(a);
		ensure(ax.size() == 1);
		ensure(ax[0] == 1.23);
		ensure(ax(0,0) == 1.23);
		double s = 0;
		for (const auto& i : a) {
			s += i;
		}
		ensure(s == 1.23);
		s = 0;
		for (const auto& i : ax) {
			s += i;
		}
		ensure(s == 1.23);
	}
	{
		FPX a;
 		ensure(a.rows() == 0);
		ensure(a.columns() == 0);
		a[0] = 1.23; // always allocates for one double

		FPX a2{ a };
		a = a2;
		ensure(a == a2);
	}
	{
		FPX a(2,3);
		FPX b{ a };
		a = b;
		ensure(a.rows() == 2);
		ensure(a.columns() == 3);
		ensure(a.size() == 6);

		a(1, 0) = 1.23;
		ensure(a[3] == 1.23);
		ensure(a == a);
		ensure(a != b);

		a.resize(3, 2);
		ensure(a.rows() == 3);
		ensure(a.columns() == 2);
		ensure(a.size() == 6);
		ensure(a[3] == 1.23);

		a.resize(3, 3);
		ensure(a.rows() == 3);
		ensure(a.columns() == 3);
		ensure(a.size() == 9);
		ensure(a[3] == 1.23);

		a.resize(2, 2);
		ensure(a.rows() == 2);
		ensure(a.columns() == 2);
		ensure(a.size() == 4);
		ensure(a[3] == 1.23);

		a.resize(0, 0);
	}

	return 0;
}
();

int test_handle = []()
{
	{
  		int* pi = new int(2);
		HANDLEX h = to_handle(pi);
		ensure(h);
		int* p = to_pointer<int>(h);
		ensure(p == pi);

		delete pi;
	}

	return 0;
}
();


int test_xloper = []()
{
	{
		ensure(0 == strcmp(XLL_DOUBLE, "B"));
		ensure(0 == strcmp(XLL_FP12, "K%"));
	}

	return 0;
}
();


int test_oper_cvt = []() {
	{
		OPER4 o(L"abc");
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 3);
		ensure(0 == strncmp(o.val.str + 1, "abc", 3));
	}
	{
		OPER12 o("abc");
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 3);
		ensure(0 == wcsncmp(o.val.str + 1, L"abc", 3));
	}
	{
		OPER o(L"abc");
		ensure(o == OPER("abc"));
	}
	{
		OPER o(L"abc");
		ensure(o == OPER(L"abc"));
	}

	return 0;
}();

int test_oper_err = []() {
	ensure(ErrValue4.xltype == xltypeErr);
	ensure(ErrValue4.val.err == xlerrValue);

	ensure(ErrValue12.xltype == xltypeErr);
	ensure(ErrValue12.val.err == xlerrValue);

	ensure(ErrValue.xltype == xltypeErr);
	ensure(ErrValue.val.err == xlerrValue);

	return 0;
}();
