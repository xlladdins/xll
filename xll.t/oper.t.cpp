// oper.t.cpp - test OPER type
#include <cassert>
#include "../xll/oper.h"

using namespace xll;

int test_oper_adt()
{
	{
		OPER o("abc");
		OPER o2(o);
		o = o2;
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == 3);
		assert(o.val.str[3] == 'c');
	}
	{
		const char* abc = "abc";
		OPER o(abc);
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == 3);
		assert(o.val.str[3] == 'c');
		const char* de = "de";
		o = de;
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == 2);
		assert(o.val.str[2] == 'e');
	}
	{
		OPER12 o(L"abc");
		OPER12 o2(o);
		o = o2;
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == 3);
		assert(o.val.str[3] == 'c');
	}
	{
		const wchar_t* abc = L"abc";
		OPER12 o(abc);
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == 3);
		assert(o.val.str[3] == 'c');
		const wchar_t* de = L"de";
		o = de;
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == 2);
		assert(o.val.str[2] == 'e');
	}

	return 0;
}
int test_oper_adt_ = test_oper_adt();

int test_oper_default()
{
	{
		OPER o;
		assert(o.xltype == xltypeNil);
	}
	{
		OPER12 o;
		assert(o.xltype == xltypeNil);
	}

	return 0;
}
int test_oper_default_ = test_oper_default();

int test_oper_num()
{
	{
		OPER o(1.23);
		assert(o.xltype == xltypeNum);
		assert(o.val.num == 1.23);
	}
	{
		OPER12 o(1.23);
		assert(o.xltype == xltypeNum);
		assert(o.val.num == 1.23);
	}
	{
		OPER o;
		o = 1.23;
		assert(o.xltype == xltypeNum);
		assert(o.val.num == 1.23);
	}
	{
		OPER12 o;
		o = 1.23;
		assert(o.xltype == xltypeNum);
		assert(o.val.num == 1.23);
	}

	return 0;
}
int test_oper_num_ = test_oper_num();

int test_oper_str()
{
	{
		OPER o("abc");
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == (char)strlen("abc"));
		assert(0 == strncmp("abc", o.val.str + 1, o.val.str[0]));
	}
	{
		const char* abc = "abc";
		OPER o(abc);
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == (char)strlen("abc"));
		assert(0 == strncmp("abc", o.val.str + 1, o.val.str[0]));
	}
	{
		auto abc = "abc";
		OPER o(abc);
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == (char)strlen("abc"));
		assert(0 == strncmp("abc", o.val.str + 1, o.val.str[0]));
	}
	{
		OPER12 o(L"abc");
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == (wchar_t)wcslen(L"abc"));
		assert(0 == wcsncmp(L"abc", o.val.str + 1, o.val.str[0]));
	}
	{
		const wchar_t* abc = L"abc";
		OPER12 o(abc);
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == (char)wcslen(abc));
		assert(0 == wcsncmp(abc, o.val.str + 1, o.val.str[0]));
	}
	{
		auto abc = L"abc";
		OPER12 o(abc);
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == (char)wcslen(L"abc"));
		assert(0 == wcsncmp(L"abc", o.val.str + 1, o.val.str[0]));
	}
	{
		OPER o;
		o = "abc";
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == (char)strlen("abc"));
		assert(0 == strncmp("abc", o.val.str + 1, o.val.str[0]));
	}
	{
		OPER12 o;
		auto abc = L"abc";
		o = abc;
		assert(o.xltype == xltypeStr);
		assert(o.val.str[0] == (wchar_t)wcslen(abc));
		assert(0 == wcsncmp(abc, o.val.str + 1, o.val.str[0]));
	}

	return 0;
}
int test_oper_str_ = test_oper_str();

int test_oper_multi()
{
	{
		OPER m(2, 3);
		assert(m.xltype == xltypeMulti);
		assert(m.rows() == 2);
		assert(m.columns() == 3);
		assert(m.size() == 6);
		assert(m[1] == OPER{});
		m[1] = "foo";
		assert(m[1].xltype == xltypeStr);
		assert(m[1] == OPER("foo"));
		m.resize(3, 2);
		assert(m.xltype == xltypeMulti);
		assert(m.rows() == 3);
		assert(m.columns() == 2);
		assert(m.size() == 6);
		assert(m[1] == OPER("foo"));
	}

	return 0;
}
int test_oper_multi_ = test_oper_multi();

int test_oper_bool()
{
	{
		OPER o(true);
		assert(o.xltype == xltypeBool);
		assert(o.val.xbool == TRUE);
	}
	{
		OPER12 o(false);
		assert(o.xltype == xltypeBool);
		assert(o.val.xbool == FALSE);
	}
	{
		OPER o;
		assert(o.xltype == xltypeNil);
		o = true;
		assert(o.xltype == xltypeBool);
		assert(o.val.xbool == TRUE);
		o = false;
		assert(o.xltype == xltypeBool);
		assert(o.val.xbool == FALSE);
	}

	return 0;
}

int test_oper_int()
{
	{
		OPER o(123);
		assert(o.xltype == xltypeInt);
		assert(o.val.w == 123);
	}
	{
		OPER12 o(-123);
		assert(o.xltype == xltypeInt);
		assert(o.val.w == -123);
	}

	return 0;
}
int test_oper_bool_ = test_oper_bool();

int test_compare()
{
	{
		OPER o1(1.23);
		XLOPER o1_ = { .val = { .num = 2.34 }, .xltype = xltypeNum };
		assert(o1 == o1);
		assert(o1_ == o1_);
		assert(o1 < o1_);
		assert(o1 <= o1_);
		assert(o1_ > o1);
		assert(o1_ >= o1);
	}
	{
		OPER o1(1.23), o1_(2.34);
		assert(o1 == o1);
		assert(o1 < o1_);
		//assert(o1 != o1_);
		assert(o1 <= o1_);
		assert(o1_ > o1);
		assert(o1_ >= o1);
	}
	{
		OPER s1("abc"), s1_("def");
		assert(xloper_cmp(s1, s1_) < 0);
		assert(s1 == s1);
		assert(s1 < s1_);
		assert(s1 <= s1_);
		assert(!(s1 == s1_));
	}

	return 0;
}
int test_compare_ = test_compare();
