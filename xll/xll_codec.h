// xll_codec.h - encoder/decoder for XLOPERs
#pragma once
#include <charconv>
#include <chrono>
#include <concepts>
#include <cstring>
#include <iostream>
#include <limits>
#include "excel.h"
#include "utf8.h"
#include "fms_view.h"

namespace xll {

	// simple and inefficent decoder for Excel types
	template<class X, class T = typename traits<X>::xcstr>
	inline XOPER<X> decode(const fms::view<const T>& v)
	{
		XOPER<X> o(v.buf, v.len);

		// num/date
		XOPER<X> o_ = Excel(xlfValue, o);
		if (!o_.is_err()) {
			o = o_;
		}
		else {
			// bool or err
			o_ = Excel(xlfEvaluate, o);
			if (!o_.is_err() or o == XOPER<X>(xll_err_str[o_.val.err])) {
				o = o_;
			}
		}

		return o;
	}

#ifdef _DEBUG

#define XLL_CODEC_TEST(X) \
	X("", "") \
	X("123", 123) \
	X("1.23", 1.23) \
	X("TRUE", true) \
	X("fAlSe", false) \
	X("2001-2-3", Excel(xlfDate, OPER(2001), OPER(2), OPER(3))) \

	inline int codec_test()
	{
#define CODEC_CHECK(a,b) ensure(decode<XLOPERX>(fms::view<const TCHAR>(_T(a))) == b);
		XLL_CODEC_TEST(CODEC_CHECK)
#undef CODEC_CHECK
#define CODEC_CHECK(a,b,...) ensure(decode<XLOPERX>(fms::view<const TCHAR>(_T(b))) == XOPER<XLOPERX>(XOPER<XLOPERX>::Err::a));
		XLL_ERR(CODEC_CHECK)
#undef CODEC_CHECK

		return 0;
	}

#undef XLL_CODEC_TEST

#endif // _DEBUG

} // namespace xll
