// oper.t.cpp - test OPER type
#undef NDEBUG
#include <cassert>
#include <algorithm>
#include <functional>
#include "../xll/xll.h"

using namespace xll;

int XLL_ERROR(const char*, bool)
{
	return 0;
}
int XLL_INFO(const char*, bool)
{
	return 0;
}

int test_defines()
{
	return 0;
}
int test_defines_ = test_defines();

int test_oper_default()
{
	{
		OPER4 o;
		ensure(o.xltype == xltypeNil);
		OPER4 o2(o);
		ensure(o2.xltype == xltypeNil);
		o = o2;
		ensure(o.xltype == xltypeNil);
	}
	{
		OPER12 o;
		ensure(o.xltype == xltypeNil);
		OPER12 o2(o);
		ensure(o2.xltype == xltypeNil);
		o = o2;
		ensure(o.xltype == xltypeNil);
	}
	{
		OPER o;
		ensure(o.xltype == xltypeNil);
		ensure(o.is_nil());
	}

	return 0;
}
int test_oper_default_ = test_oper_default();

int test_oper_num()
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
		ensure(o.xltype == xltypeNum); // just like Excel
		ensure(o.val.num == 1);
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
int test_oper_num_ = test_oper_num();

int test_oper_str()
{
	{
		OPER4 o4("");
		ensure(o4.is_str());
		ensure(o4.val.str[0] == 0);

		OPER12 o12("");
		ensure(o12.is_str());
		ensure(o12.val.str[0] == 0);
	}
	{
		OPER o;
		o = "";
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 0);
		o.append("");
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 0);
	}
	{
		OPER o(_T(""), 3);
		ensure(o.val.str[0] == 3);
		ensure(o.val.str[1] == 0);
	}
	{
		OPER o("abc");
		auto s = o.as_cstr();
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 3);
		ensure(o.val.str[3] == 'c');
		ensure(o.val.str[4] == 0);
		ensure(s[0] == 'a');
		ensure(3 == traits<XLOPERX>::len(s));

		s = o.as_cstr();
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 3);
		ensure(o.val.str[3] == 'c');
		ensure(o.val.str[4] == 0);
		ensure(s[0] == 'a');
		ensure(3 == traits<XLOPERX>::len(s));
	}
	{
		OPER4 o("abc");
		OPER4 o2(o);
		o = o2;
		ensure(o.xltype == xltypeStr);
		ensure(o.val.str[0] == 3);
		ensure(o.val.str[3] == 'c');
		std::string s = o.to_string();
		assert(s == "abc");
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
		std::string s = o.to_string();
		assert(s == "abc");
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
	{
		OPER o;
		o.append(OPER{});
		ensure(o.xltype == xltypeNil);
		OPER foo("foo");
		o.append(foo);
		ensure(o == foo);
	}
	{
		OPER o("foo");
		o.append(OPER{});
		ensure(o == "foo");
	}
	{
		OPER4 o4("foo");
		OPER12 o12(L"bar");
		o4.append(o12);
		ensure(o4 == "foobar");
		o12.append(o4);
		ensure(o12 == "barfoobar");
	}

	return 0;
}
int test_oper_str_ = test_oper_str();

int test_oper_multi()
{
	{
		OPER m(1, 1);
		m[0] = "abc";
		ensure(m[0] == "abc");
		ensure(m != "abc");
	}
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
		ensure(m(1,0) == _T("foo"));
		ensure(m[3] == _T("foo"));

		m.resize(3, 2);
		ensure(m.xltype == xltypeMulti);
		ensure(m.rows() == 3);
		ensure(m.columns() == 2);
		ensure(m.size() == 6);
		ensure(m[3] == _T("foo"));
		ensure(m(1, 1) == _T("foo"));

		m.resize(2, 2);
		ensure(m.xltype == xltypeMulti);
		ensure(m.rows() == 2);
		ensure(m.columns() == 2);
		ensure(m.size() == 4);
		ensure(m[3] == _T("foo"));
		ensure(m(1, 1) == _T("foo"));

		m.resize(3, 3);
		ensure(m.xltype == xltypeMulti);
		ensure(m.rows() == 3);
		ensure(m.columns() == 3);
		ensure(m.size() == 9);
		ensure(m[3] == _T("foo"));
		ensure(m(1, 0) == _T("foo"));

		m(1, 2) = OPER("bar");
		ensure(m(1, 2) == _T("bar"));

		m(2, 1) = m; // multis can nest
		ensure(m(2, 1)(1, 2) == _T("bar"));
	}
	{
		OPER o;
		ensure(o.xltype == xltypeNil);
		ensure(o.size() == 0);
		o.push_back(o);
		ensure(o.xltype == xltypeNil);
	}
	{
		OPER o({ OPER(1.23) } );
		ensure(o.xltype == xltypeNum);
		ensure(o.rows() == 1);
		ensure(o.columns() == 1);
		ensure(o[0] == 1.23);
		ensure(o == 1.23);
		o.push_back(o);
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
		ensure(o[1] == _T("foo"));
		o.push_back(o);
		ensure(o.rows() == 2);
		ensure(o.columns() == 2);
		ensure(o[0] == 1.23);
		ensure(o[1] == _T("foo"));
		ensure(o[2] == 1.23);
		ensure(o[3] == _T("foo"));
	}
	{
		OPER o({ OPER(1.23), OPER("foo") });
		ensure(o.xltype == xltypeMulti);
		ensure(o.rows() == 1);
		ensure(o.columns() == 2);
		ensure(o[0] == 1.23);
		ensure(o[1] == _T("foo"));
		XLOPERX x;
		XLOPERX x2[2];
		x2[0].xltype = xltypeMissing;
		x2[1].xltype = xltypeErr;
		x.xltype = xltypeMulti;
		x.val.array.rows = 1;
		x.val.array.columns = 2;
		x.val.array.lparray = x2;
		o.push_back(x);
		ensure(o.rows() == 2);
		ensure(o.columns() == 2);
		ensure(o[0] == 1.23);
		ensure(o[1] == _T("foo"));
		ensure(o[2].xltype == xltypeMissing);
		ensure(o[3].xltype == xltypeErr);
	}
	{
		OPER o(2, 1);
		o(0, 0) = 0;
		o(1, 0) = 10;
		OPER o2(2, 1);
		o2(0, 0) = 1;
		o2(1, 0) = 11;
		o.push_back(o2, OPER::Side::Right); // push back column
		//o.push_back(o, OPER::Side::Right);
		assert(o(0, 0) == 0);
		assert(o(0, 1) == 1);
		assert(o(1, 0) == 10);
		assert(o(1, 1) == 11);
	}

	return 0;
}
int test_oper_multi_ = test_oper_multi();

OPER oiota(unsigned a, unsigned n)
{
	OPER o(1, n);

	for (unsigned i = 0; i < n; ++i)
		o[i] = a + i;
	
	return o;
}

int test_oper_drop()
{
	{
		auto o1 = oiota(0, 6);
		o1.drop(0);
		assert(o1 == oiota(0, 6));
		o1.resize(2, 3);
		o1.drop(0);
		assert(o1 == oiota(0, 6).resize(2,3));
	}
	{
		assert(oiota(0, 2).drop(2).size() == 0);
		assert(oiota(0, 2).drop(-2).size() == 0);
		assert(oiota(0, 2).drop(3).size() == 0);
		assert(oiota(0, 2).drop(-3).size() == 0);
	}
	{
		assert(oiota(0, 6).drop(1) == oiota(1, 5));
		assert(oiota(0, 6).drop(-1) == oiota(0, 5));
		assert(oiota(0, 6).resize(3, 2).drop(1) == oiota(2, 4).resize(2,2));
		assert(oiota(0, 6).resize(3, 2).drop(-1) == oiota(0, 4).resize(2,2));
		assert(oiota(0, 6).resize(3, 2).drop(3).size() == 0);
		assert(oiota(0, 6).resize(3, 2).drop(-3).size() == 0);
	}

	return 0;
}
int test_oper_drop_ = test_oper_drop();

int test_oper_take()
{
	{
		auto o1 = oiota(0, 6);
		o1.take(0);
		assert(o1.size() == 0);
		o1.resize(2, 3);
		o1.take(0);
		assert(o1.size() == 0);
	}
	{
		assert(oiota(0, 2).take(2) == oiota(0, 2));
		assert(oiota(0, 2).take(-2) == oiota(0, 2));
		assert(oiota(0, 2).take(3) == oiota(0, 2));
		assert(oiota(0, 2).take(-3) == oiota(0, 2));
	}
	{
		assert(oiota(0, 6).take(1) == oiota(0, 1));
		auto a = oiota(0, 6).take(-1);
		auto b = oiota(5, 1);
		assert(a == b);
		assert(oiota(0, 6).take(-1) == oiota(5, 1));
		assert(oiota(0, 6).resize(3, 2).take(2) == oiota(0, 4).resize(2, 2));
		assert(oiota(0, 6).resize(3, 2).take(-2) == oiota(2, 4).resize(2, 2));
		assert(oiota(0, 6).resize(3, 2).take(3) == oiota(0, 6).resize(3, 2));
		assert(oiota(0, 6).resize(3, 2).take(-3) == oiota(0, 6).resize(3, 2));
	}

	return 0;
}
int test_oper_take_ = test_oper_take();

int test_oper_bool()
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
int test_oper_bool_ = test_oper_bool();

// ints converted to double, just like Excel
int test_oper_int()
{
	{
		OPER o(123);
		ensure(o.xltype == xltypeNum);
		ensure(o.val.num == 123);
	}
	{
		OPER12 o(-123);
		ensure(o.xltype == xltypeNum);
		ensure(o.val.num == -123);
	}

	return 0;
}
int test_oper_int_ = test_oper_int();

int test_compare()
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
int test_compare_ = test_compare();

inline FPX aiota(double t, unsigned n) {
	FPX a(1, n);

	for (unsigned i = 0; i < n; ++i) {
		a[i] = t + i;
	}

	return a;
};

int test_fp()
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
		for (const auto& i : ax) {
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
 		ensure(a.rows() == 1);
		ensure(a.columns() == 1);
		ensure(a.size() == 1);
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
	{
		FPX a = aiota(0, 3);
		ensure(a.drop(1) == aiota(1, 2));
	}
	{
		FPX a = aiota(0, 3);
		ensure(a.drop(-1) == aiota(0, 2));
	}
	{
		FPX a = aiota(0, 6);
		a.resize(2, 3);
		ensure(a.drop(1) == aiota(3, 3));
	}
	{
		FPX a = aiota(0, 6);
		a.resize(3, 2);
		ensure(a.drop(-1) == aiota(0, 4).resize(2,2));
	}

	{
		FPX a = aiota(0, 3);
		ensure(a.take(1) == aiota(0, 1));
	}
	{
		FPX a = aiota(0, 3);
		ensure(a.take(-1) == aiota(2, 1));
	}
	{
		FPX a = aiota(0, 6);
		a.resize(2, 3);
		ensure(a.take(1) == aiota(0, 3));
	}
	{
		FPX a = aiota(0, 6);
		a.resize(3, 2);
		ensure(a.take(-1) == aiota(4, 2));
	}
	{
		FPX a = aiota(1, 3);
		scan(std::plus<double>{}, a.get());
		ensure(a[0] == 1);
		ensure(a[1] == 3);
		ensure(a[2] == 6);
	}
	{
		FPX a = aiota(1, 3); // {1,2,3}
		FPX m = aiota(0, 2); // {0, 1, 0, 1, ...}
		mask(a.get(), m.get());
		ensure(a.size() == 1);
		ensure(a[0] == 2);
	}
	{
		FPX a = aiota(1, 6); // {1,2,3,4,5,6}
		a.resize(3, 2);
		FPX m = aiota(0, 2); // {0, 1, 0, 1, ...}
		mask(a.get(), m.get());
		ensure(a.rows() == 1);
		ensure(a.columns() == 2);
		ensure(a(0,0) == 3);
		ensure(a(0, 1) == 4);
	}
	{
		FPX a = aiota(1, 3);
		rotate(a.get(), 1);
		ensure(a.size() == 3);
		ensure(a[0] == 2);
		ensure(a[1] == 3);
		ensure(a[2] == 1);
	}
	{
		FPX a = aiota(1, 3);
		rotate(a.get(), -1);
		ensure(a.size() == 3);
		ensure(a[0] == 3);
		ensure(a[1] == 1);
		ensure(a[2] == 2);
	}
	{
		FPX a = aiota(1, 3);
		shift(a.get(), 1);
		ensure(a.size() == 3);
		ensure(a[0] == 0);
		ensure(a[1] == 1);
		ensure(a[2] == 2);
	}
	{
		FPX a = aiota(1, 3);
		shift(a.get(), -1);
		ensure(a.size() == 3);
		ensure(a[0] == 2);
		ensure(a[1] == 3);
		ensure(a[2] == 0);
	}

	return 0;
}
int test_fp_ = test_fp();

int test_handle()
{
	{
  		int* pi = new int(2);
		HANDLEX h = to_handle(pi);
		ensure(h);
		int* p = to_pointer<int>(h);
		ensure(p == pi);

		delete pi;
	}
	{
		handle<int> h(new int(2));
		ensure(*h.ptr() == 2);
		ensure(*h == 2);
		//handle<int> h2(h); // deleted
		handle<int> h2(std::move(h));
		ensure(*h == 2);
		ensure(*h2 == 2);
		//h = h2; // deleted
		h = std::move(h2);
		ensure(*h == 2);
		ensure(*h2 == 2);
		h2.swap(h);
		ensure(*h == 2);
		ensure(*h2 == 2);

		handle<int> h3(new int(3));
		ensure(*h == 2);
		ensure(*h3 == 3);
	}
	{
		handle<int>::codec<XLOPERX> hc("foo[0x", "]");
		handle<int> _h(new int{ 123 });
		OPER H = hc.encode(_h.get());
		ensure(0 == _tcsncmp(H.val.str + 1, _T("foo[0x"), 6));
		/*
		ensure(pre == Excel(xlfLeft, H, Excel(xlfLen, pre)));
		ensure(suf == Excel(xlfRight, H, Excel(xlfLen, suf)));
		*/
		HANDLEX h_ = hc.decode(H);
		ensure(h_ == _h.get());
	}

	return 0;
}
int test_handle_ = test_handle();


int test_xloper()
{
	{
		ensure(0 == strcmp(XLL_DOUBLE, "B"));
		ensure(0 == strcmp(XLL_FP12, "K%"));
	}

	return 0;
}
int test_xloper_ = test_xloper();


int test_oper_cvt() 
{
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
}
int test_oper_cvt_ = test_oper_cvt();

int test_oper_err()
{
	ensure(ErrValue4.xltype == xltypeErr);
	ensure(ErrValue4.val.err == xlerrValue);

	ensure(ErrValue12.xltype == xltypeErr);
	ensure(ErrValue12.val.err == xlerrValue);

	ensure(ErrValue.xltype == xltypeErr);
	ensure(ErrValue.val.err == xlerrValue);

	assert(OPER(OPER::Err::Value) == ErrValue);

	return 0;
}
int test_oper_err_ = test_oper_err();

int test_mref()
{
	{
		REF r(1, 1);
		OPER mr({ r,r,r });
		assert(mr.size() == 1);
		assert(mr.type() == xltypeRef);
		assert(mr == mr);
		OPER mr2(mr);
		mr = mr2;
		assert(mr == mr2);
	}
	{
		REF12 r(1, 1);
		OPER12 mr({ r,r,r });
		assert(mr.size() == 1);
		assert(mr == mr);
		OPER12 mr2(mr);
		mr = mr2;
		assert(mr == mr2);
	}

	return 0;
}
int test_mref_ = test_mref();
/*
int test_max()
{
	{
		OPER x(1.23);
		OPER u = xll::max(x);
		assert(x == (std::max)(x, u));
		assert(x == (std::max)(u, x));
	}
	{
		OPER x("abc");
		OPER u = xll::max(x);
		assert(u <= x);
		assert(x >= u);
	}
	{
		OPER x(true);
		OPER u = xll::max(x);
		assert(u <= x);
		assert(x >= u);
	}
	{
		OPER x(false);
		OPER u = xll::max(x);
		assert(u <= x);
		assert(x >= u);
	}
	{
		XLOPERX x = { .val = {.w = 123}, .xltype = xltypeInt };
		OPER u = xll::max(x);
		assert(u <= x);
		assert(x >= u);
	}
	{
		OPER x(2, 3);
		OPER u = xll::max(x);
		assert(u <= x);
		assert(x >= u);
	}

	return 0;
}
int test_max_ = test_max();
*/