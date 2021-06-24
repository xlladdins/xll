// codec.h - encoder/decoder for XLOPERs
// XLOPER is a variant type num|str|err|multi|...
#pragma once
#include <concepts>
#include <iostream>
#include "excel.h"

namespace xll {

	template<class X>
	inline XOPER<X> decode(const typename traits<X>::xcstr str, DWORD len = 0)
	{
		XLOPER<X> o = len == 0 ? XOPER<X>(str) : XOPER<X>(str, len);
		
		// convert to number, date, boolean, or error
		XLOPER<X> o_ = XExcel<X>(xlfEvaluate, o);

		// error strings evaluate to xltypeErr
		if (!o_.is_err() or o == XOPER<X>(xll_err_str[o_.val.err])) {
			o = o_;
		}

		return o;
	}

}
