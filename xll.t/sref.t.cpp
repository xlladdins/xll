// sref.t.cpp - test XREF
#include "../xll/sref.h"

using namespace xll;

int test_sref = [] {
	{
		REF4 r4;
		REF12 r12;
		REF r;

		REF4 r4_(r4);
		r12 = r12;
	}
	{
		XLREF r{ 1, 1, 1, 1 };
		//ensure(rows(r) == 1);
		//ensure(columns(r) == 1);
		ensure(r == r);
		ensure(!(r < r));
		ensure(!(r > r));
	}
	
	{
		REF r(2, 3);
		ensure(height(r) == 1);
		ensure(width(r) == 1);
		REF r2(r);
		ensure(r2 == r);
		r = r2;
		ensure(r2 == r);
	}
	

	return 0;
}();