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
#include "view.h"

namespace xll {

	// return NaN if not a floating point number
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> to_num(xll::view<const T>& v)
	{
		XOPER<X> num = std::numeric_limits<double>::quiet_NaN();

		if (v.len == 0) {
			return num;
		}

		const char* buf = nullptr; // if wide char
		const char* str = nullptr;
		size_t len = 0;
		
		if constexpr (std::is_same_v<T, typename traits<XLOPER>::xchar>) {
			str = v.buf;
			len = v.len;
		}
		else if constexpr (std::is_same_v<T, typename traits<XLOPER12>::xchar>) {
			buf = utf8::wcstombs(v.buf, v.len);
			str = buf + 1;
			len = buf[0];
		}
		else {
			return num;
		}

		auto [ptr, err] = std::from_chars(str, str + len, num.val.num);
		if (ptr != str and err == std::errc()) {
			v = v.advance(static_cast<DWORD>(ptr - str));
		}

		if (buf) {
			free((void*)buf);
		}

		return num;
	}

#ifdef _DEBUG

	inline int xll_to_num_test()
	{
		OPER num;

		{
			view<const TCHAR> v(_T(""));
			num = to_num<XLOPERX>(v);
			ensure(num.is_num() and isnan(num.val.num));
		}
		{
			view<const TCHAR> v(_T("123"));
			num = to_num<XLOPERX>(v);
			ensure(num == 123);
		}
		{
			view<const TCHAR> v(_T("1.23"));
			num = to_num<XLOPERX>(v);
			ensure(num == 1.23);
		}
		{
			view<const TCHAR> v(_T("1.23 abc"));
			num = to_num<XLOPERX>(v);
			ensure(num == 1.23);
			ensure(v.equal(view<const TCHAR>(_T(" abc"))));
		}

		return 0;
	}

#endif // _DEBUG

	// case insensitve, like Excel
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> to_bool(xll::view<const T>& v)
	{
		XOPER<X> xbool = XOPER<X>(XOPER<X>::Err::NA);

		if (v.len >= 4 and XOPER<X>(v.buf, 4) == "TRUE") {
			xbool = true;
			v.advance(4);
		}
		else if (v.len >= 5 and XOPER<X>(v.buf, 5) == "FALSE") {
			xbool = false;
			v.advance(5);
		}

		return xbool;
	}

#ifdef _DEBUG

	inline int xll_to_bool_test()
	{
		OPER xbool;

		{
			view<const TCHAR> v(_T("true"));
			ensure(to_bool<XLOPERX>(v) == true);
		}
		{
			view<const TCHAR> v(_T("fAlSe"));
			ensure(to_bool<XLOPERX>(v) == false);
		}
		{
			view<const TCHAR> v(_T("truedat"));
			ensure(to_bool<XLOPERX>(v) == true);
			ensure(v.equal(view<const TCHAR>(_T("dat"))));
		}
		{
			view<const TCHAR> v(_T("dunno"));
			xbool = to_bool<XLOPERX>(v);
			ensure(xbool.is_err());
			ensure(xbool == ErrNA);
		}

		return 0;
	}

#endif // _DEBUG

	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> to_date(xll::view<const T>& v)
	{
		XOPER<X> date = XOPER<X>(XOPER<X>::Err::NA);
		XOPER<X> xv(v.buf, v.len);

		// number/date, boolean, or error
		XOPER<X> val = Excel(xlfValue, xv);
		if (val) {
			XOPER<X> eval = Excel(xlfEvaluate, xv);
			if (val != eval) {
				date = val; // date, not just a number
				v.advance(v.len);
			}
		}

		return date;
	}

#ifdef _DEBUG

	inline int xll_to_date_test()
	{
		OPER o;

		{
			view<const TCHAR> v(_T("2001-2-3"));
			o = to_date<XLOPERX>(v);
			ensure(o == Excel(xlfDate, OPER(2001), OPER(2), OPER(3)));
		}

		return 0;
	}

#endif // _DEBUG

	template<class X>
	struct codec_traits {};

	template<> struct codec_traits<int>    { typedef int type;    };
	template<> struct codec_traits<double> { typedef double type; };
	template<> struct codec_traits<bool>   { typedef bool type;   };
	template<> struct codec_traits<time_t> { typedef time_t type; };

	// decode string of T to X
	template<class X, class T>
	struct decoder {
		enum class data_type {
			integer,
			number, // floating point
			boolean,
			date,
			time,
			datetime,
			
		};
		virtual ~decoder()
		{ }
		static X to_num(view<const T>& v)
		{
			return to_num_(v);
		}
		// to_bool
		// to_date
		// ...
	private:
		virtual X to_num_(view<const T>& v) = 0;
	};

	template<class X, class T = typename traits<X>::xchar>
	struct excel_decoder : public decoder<X,T> {
		XOPER<X> to_num_(view<const T>& v) override
		{
			return xll::to_num<T>(v);
		}
	};

	template<class X>
	inline XOPER<X> decode(const typename traits<X>::xcstr str, DWORD len = 0)
	{
		XOPER<X> o = len == 0 ? XOPER<X>(str) : XOPER<X>(str, len);

		// date or boolean
		XOPER<X> o_ = Excel(xlfValue, o);
		if (!o_.is_err()) {
			o = o_;
		}

		// convert to number, boolean, or error
		o_ = XExcel<X>(xlfEvaluate, o);

		// error strings evaluate to xltypeErr
		if (!o_.is_err() or o == XOPER<X>(xll_err_str[o_.val.err])) {
			o = o_;
		}


		return o;
	}


#ifdef _DEBUG

	int codec_test()
	{
		xll_to_num_test();
		xll_to_bool_test();
#ifndef _CONSOLE
		xll_to_date_test(); // Excel not available in xll.t
#endif
		return 0;
	}

#endif // _DEBUG

} // namespace xll
