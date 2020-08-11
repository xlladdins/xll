// oper.h - Wrapper for XLOPER
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <algorithm>
#include <concepts>
#include <iostream>
#include <utility>
#include "utf8.h"
#include "xloper.h"

#pragma warning(disable: 4996)

namespace xll {

	/// <summary>
	///  Wrapper for XLOPER/XLOPER12 structs.
	/// </summary>
	/// <typeparam name="XLOPERX">
	/// Either XLOPER or XLOPER12 
	/// </typeparam>
	/// An OPER corresponds to a cell or a 2-dimensional range of cells.
	/// It is a variant type that can be either a number, string, boolean,
	/// error, range, single reference, or integer. 
	template<class X>  
		requires (std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>)
	class XOPER final : public X {
		typedef typename traits<X>::typex X_; // the _other_ type
		typedef typename traits<X>::xchar xchar;
		typedef typename traits<X_>::xchar charx;
		typedef typename traits<X>::xcstr xcstr;
		typedef typename traits<X_>::xcstr cstrx;
		typedef typename traits<X>::xint xint;
		typedef typename traits<X>::xrw xrw;
		typedef typename traits<X>::xcol xcol;
		typedef typename traits<X>::xint xint;

	public:
		using X::xltype;
		using X::val;

		friend void swap(XOPER& a, XOPER& b) noexcept
		{
			using std::swap;

			swap(a.xltype, b.xltype);
			swap(a.val, b.val);
		}

		XOPER()
		{
			xltype = xltypeNil;
		}
		XOPER(const X& x)
		{
			if (x.xltype == xltypeStr) {
				str_alloc(x.val.str + 1, x.val.str[0]);
			}
			else if (x.xltype == xltypeMulti) {
				multi_alloc(x.val.array.rows, x.val.array.columns);
				std::copy(x.val.array.lparray, x.val.array.lparray + size(), begin());
			}
			else {
				xltype = x.xltype;
				val = x.val;
			}
		}
		explicit XOPER(const XOPER& o)
			: XOPER(static_cast<const X&>(o))
		{ }
		XOPER(XOPER&& o) noexcept
			: XOPER()
		{
			swap(*this, o);
		}
		XOPER& operator=(XOPER o)
		{
			swap(*this, o);

			return *this;
		}
		/*
		XOPER& operator=(const X& x)
		{
			swap(*this, XOPER(x));

			return *this;
		}
		XOPER& operator=(XOPER&& o) noexcept
		{
			if (this != &o) {
				xltype = o.xltype;
				val = o.val;
				o.xltype = xltypeMissing;
			}

			return *this;
		}
		*/
		~XOPER()
		{
			oper_free();
		}
		/*
		operator bool() const
		{
			switch (xltype) {
			case xltypeNum:
				return val.num != 0;
				break;
			case xltypeStr:
				return val.str[0] != 0;
				break;
			case xltypeBool:
				return val.xbool != 0;
				break;
			case xltypeMulti:
				return std::any_of(begin(), end(), [](const auto& x) { return x; });
				break;
			case xltypeInt:
				return val.w != 0;
			}

			return false; // xltypeErr, xltypeMissig, xltypeNil
		}
		*/
		// floating point number
		explicit XOPER(double num)
		{
			xltype = xltypeNum;
			val.num = num;
		}
		XOPER operator=(double num)
		{
			oper_free();
			operator=(XOPER(num));

			return *this;
		}
		bool operator==(double num) const
		{
			return xltype == xltypeNum && val.num == num;
		}
		/* not working !!!
		// handy for using OPERs in numerical expressions
		operator double()
		{
			return xltype == xltypeNum ? val.num 
				 : xltype == xltypeInt ? val.w 
				 : xltype == xltypeBool ? val.xbool 
				 : std::numeric_limits<double>::quiet_NaN();
		}
		*/

		// xltypeStr given length
		XOPER(xcstr str, int n)
		{
			str_alloc(str, n);
		}
		// xltypeStr from NULL terminated string
		explicit XOPER(xcstr str)
			: XOPER(str, traits<X>::len(str))
		{
		}
		// xltypeStr from string literal
		template<size_t N>
		XOPER(xcstr(&str)[N])
			: XOPER(str, N)
		{
		}
		// convert to type appropriate for X
		explicit XOPER(cstrx str)
		{
			str_alloc(traits<X>::cvt(str), (size_t)-1);
		}
		// Construct from string literal
		template<size_t N>
		XOPER(cstrx (&str)[N])
		{
			str_alloc(traits<X>::cvt(str), (size_t)-1);
		}
		bool operator==(xcstr str) const
		{
			if (xltype != xltypeStr) {
				return false;
			}

			size_t n = traits<X>::len(str);
			ensure(n < static_cast<size_t>(std::numeric_limits<xchar>::max()));
			
			if (val.str[0] != static_cast<xchar>(n))
				return false;

			return 0 == traits<X>::cmp(str, val.str + 1, n);
		}
		// no operator== for cstrx
		XOPER operator=(xcstr str)
		{
			oper_free();
			str_alloc(str, traits<X>::len(str));

			return *this;
		}
		XOPER operator=(cstrx str)
		{
			oper_free();
			str_alloc(traits<X>::cvt(str), (size_t)-1);

			return *this;
		}
		XOPER& append(const X& x)
		{
			ensure(x.xltype == xltypeStr);
			str_append(x.val.str + 1, x.val.str[0]);

			return *this;
		}
		XOPER& append(const X_& x)
		{
			ensure(x.xltype == xltypeStr);
			str_append(traits<X>::cvt(x.val.str + 1, x.val.str[0]), (size_t)-1);

			return *this;
		}
		// Like the Excel ampersand operator
		XOPER& operator&=(const X& x)
		{
			return append(x);
		}
		XOPER& operator&=(const X_& x)
		{
			return append(x);
		}
		XOPER& append(xcstr str)
		{
			str_append(str, traits<X>::len(str));

			return *this;
		}
		XOPER& append(cstrx str)
		{
			str_append(traits<X>::cvt(str), (size_t)-1);

			return *this;
		}
		XOPER& operator&=(xcstr str)
		{
			return append(str);
		}
		XOPER& operator&=(cstrx str)
		{
			return append(str);
		}

		// xltypeBool
		explicit XOPER(bool xbool)
		{
			xltype = xltypeBool;
			val.xbool = xbool;
		}
		bool operator==(bool xbool) const
		{
			return xltype == xltypeBool && val.xbool == static_cast<typename traits<X>::xbool>(xbool);
		}
		XOPER operator=(bool xbool)
		{
			oper_free();
			operator=(XOPER(xbool));

			return *this;
		}

		// xltypeRef
		// xltypeErr - predifined as ErrXXX
		// xltypeFlow - not used

		// xltypeMulti
		XOPER(xrw rw, xcol col)
		{
			multi_alloc(rw, col);
		}
		//??? XOPER(std::initializer_list<XOPER> o)
		XOPER& resize(xrw rw, xcol col)
		{
			multi_realloc(rw, col);

			return *this;
		}
		size_t rows() const
		{
			return xltype == xltypeMulti ? val.array.rows : 1;
		}
		size_t columns() const
		{
			return xltype == xltypeMulti ? val.array.columns : 1;
		}
		size_t size() const
		{
			return rows()*columns();
		}
		// STL friendly
		const XOPER* begin() const
		{
			return xltype == xltypeMulti ? static_cast<const XOPER*>(val.array.lparray) : this;
		}
		XOPER* begin()
		{
			return xltype == xltypeMulti ? static_cast<XOPER*>(val.array.lparray) : this;
		}
		const XOPER* end() const
		{
			return xltype == xltypeMulti ? begin() + size() : this + 1;
		}
		XOPER* end()
		{
			return xltype == xltypeMulti ? begin() + size() : this + 1;
		}
		// one-dimensional index
		XOPER& operator[](size_t i)
		{
			ensure(i < size());

			if (xltype == xltypeMulti) {
				return static_cast<XOPER&>(val.array.lparray[i]);
			}
			else {
				ensure(i == 0);
				return *this;
			}
		}
		const XOPER& operator[](size_t i) const
		{
			return operator[](i);
		}
		// two-dimensional index
		XOPER& operator()(xrw rw, xcol col)
		{
			return operator[](columns() * rw + col);
		}
		const XOPER& operator()(xrw rw, xcol col) const
		{
			return operator()(rw, col);
		}

		// xltypeMissing - predefined as Missing
		// xltypeNil - predefined as Nil

		// xltypeSRef - reference to a single range
		XOPER(xrw row, xcol col, xrw height, xcol width)
		{
			xltype = xltypeSRef;
			val.sref.count = 1;
			val.sref.rwFirst = row;
			val.sref.rwLast = row + height;
			val.sref.colFirst = col;
			val.sref.colLast = col + width;
		}

		// xltypeInt. Excel usually converts this to num.
		explicit XOPER(int w)
		{
			xltype = xltypeInt;
			val.w = static_cast<xint>(w);
		}
		bool operator==(int w) const
		{
			return xltype == xltypeInt ? val.w == w
				: xltype == xltypeNum ? val.num == w
				: xltype == xltypeBool ? val.xbool == w
				: false;
		}
		XOPER operator=(int w)
		{
			oper_free();
			operator=(XOPER(w));

			return *this;
		}

	private:
		void oper_free()
		{
			if (xltype == xltypeStr) {
				str_free();
			}
			else if (xltype == xltypeMulti) {
				multi_free();
			}
			
			xltype = xltypeNil;
		}

		// allocate unless counted
		void str_alloc(xcstr str, size_t n)
		{
			xltype = xltypeStr;

			if (n == (size_t)-1) {
				val.str = const_cast<xchar*>(str); // move
			}
			else {
				val.str = (xchar*)malloc((n + 1) * sizeof(xchar));
				// first character is count
				if (val.str) {
					memcpy_s(val.str + 1, n * sizeof(xchar), str, n * sizeof(xchar));
					// ensure (n <= ...MAX);
					val.str[0] = static_cast<xchar>(n);
				}
			}
		}
		void str_append(xcstr str, size_t n)
		{
			if (xltype == xltypeNil) {
				str_alloc(str, n);
			}
			else {
				ensure(xltype == xltypeStr);
				bool counted = false;
				if (n == (size_t)-1) {
					counted = true;
					n = str[0];
				}
				if (n == 0) {
					n = traits<X>::len(str);
				}
				xchar* tmp = (xchar*)realloc(val.str, (val.str[0] + n + 1) * sizeof(xchar));
				if (tmp != nullptr) {
					val.str = tmp;
					memcpy_s(val.str + 1 + val.str[0], n * sizeof(xchar), str + counted, n * sizeof(xchar));
					val.str[0] = static_cast<xchar>(val.str[0] + n);
				}
				if (counted) {
					free(const_cast<xchar*>(str));
				}
			}
		}
		void str_free()
		{
			ensure(xltype == xltypeStr);
			free(val.str);
			xltype = xltypeNil;
		}
		
		// xltypeMulti
		void multi_alloc(xrw r, xcol c)
		{
			xltype = xltypeMulti;
			val.array.rows = r;
			val.array.columns = c;
			val.array.lparray = (X*)malloc(size() * sizeof(X));
			if (val.array.lparray) {
				std::fill(begin(), end(), XOPER{});
			}
		}
		void multi_realloc(xrw r, xcol c)
		{
			ensure(xltype == xltypeMulti);
	
			// current size
			auto n = size();
			val.array.rows = r;
			val.array.columns = c;
			if (n > size()) {
				for (XOPER* po = begin() + size(); po != begin() + n; ++po) {
					po->oper_free();
				}
				val.array.lparray = (X*)realloc(val.array.lparray, size() * sizeof(X));
				ensure(val.array.lparray);
			}
			else if (n < size()) {
				val.array.lparray = (X*)realloc(val.array.lparray, size() * sizeof(X));
				ensure(val.array.lparray);
				for (XOPER* po = begin() + n; po != begin() + size(); ++po) {
					*po = XOPER{};
				}
			}
		}
		void multi_free()
		{
			ensure(xltype == xltypeMulti);
			for (XOPER* po = begin(); po != end(); ++po) {
				po->oper_free();
			}
			free(val.array.lparray);

			xltype = xltypeNil;
		}
	};

	typedef XOPER<XLOPER> OPER4;
	typedef XOPER<XLOPER12> OPER12;
	typedef XOPER<XLOPERX> OPER;

	typedef OPER4* LPOPER4;
	typedef OPER12* LPOPER12;
	typedef OPER* LPOPER;

}

// Just like Excel.
inline xll::OPER4 operator&(const XLOPER& x, const XLOPER& y)
{
	return xll::OPER4(x) &= y;
}
inline xll::OPER4 operator&(const XLOPER& x, const char* y)
{
	xll::OPER4 z(x);

	return z.append(y);
}
inline xll::OPER12 operator&(const XLOPER12& x, const XLOPER12& y)
{
	return xll::OPER12(x) &= y;
}
inline xll::OPER12 operator&(const XLOPER12& x, const wchar_t* y)
{
	return xll::OPER12(x).append(y);
}

inline auto& operator<<(std::basic_ostream<xll::traits<XLOPER>::xchar>& os, const XLOPER& x)
{
	switch (x.xltype) {
	case xltypeNum:
		os << x.val.num;
		break;
	case xltypeStr:
		os.write(x.val.str + 1, x.val.str[0]);
		break;
	case xltypeBool:
		os << static_cast<bool>(x.val.xbool);
		break;
	case xltypeErr:
		os << xll_err_str[x.val.err];
		break;
	case xltypeMulti:
		os << "{";
		for (size_t r = 0; r < xll::rows(x); ++r) {
			os << ":";
		}
		os << "}";
		break;
	default:
		break;
	}

	return os;
}
