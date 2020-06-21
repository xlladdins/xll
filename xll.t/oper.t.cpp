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
	}
	return 0;
}

int test_oper_default()
{
	{
		OPER o;
		assert(o.xltype == xltypeMissing);
	}
	{
		OPER12 o;
		assert(o.xltype == xltypeMissing);
	}

	return 0;
}

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

	return 0;
}

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

	return 0;
}

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

int test_oper_adt_ = test_oper_adt();
int test_oper_default_ = test_oper_default();
int test_oper_num_ = test_oper_num();
int test_oper_str_ = test_oper_str();
int test_oper_bool_ = test_oper_bool();