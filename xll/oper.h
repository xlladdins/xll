// oper.h - Wrapper for XLOPER
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <concepts>
#include <utility>
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
		typedef typename traits<X>::xchar xchar;
		typedef typename traits<X>::xcstr xcstr;
		typedef typename traits<X>::xint xint;
		typedef typename traits<X>::xrw xrw;
		typedef typename traits<X>::xcol xcol;
		//!!!typedef typename traits<X>::xint xint;

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
		XOPER(const XOPER& o)
			: XOPER(static_cast<const X&>(o))
		{ }
		XOPER(const X& x)
		{
			if (x.xltype == xltypeStr) {
				str_alloc(x.val.str[0]);
				std::copy(x.val.str + 1, x.val.str + 1 + x.val.str[0], val.str + 1);
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
			swap(*this, OPERX(x));

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

		// xltypeStr given length
		XOPER(xcstr str, size_t n)
		{
			str_alloc(n);
			std::copy(str, str + n, val.str + 1);
		}
		// xltypeStr from NULL terminated string
		explicit XOPER(xcstr str)
			: XOPER(str, traits<X>::len(str))
		{
		}
		// Construct from string literal
		template<size_t N>
		XOPER(const xchar (&str)[N])
			: XOPER(str, N)
		{
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

			return 0 == traits<X>::cmp(str, val.str + 1, val.str[0]);
		}
		XOPER operator=(xcstr str)
		{
			oper_free();
			operator=(XOPER(str));

			return *this;
		}
		// Like the Excel ampersand operator
		XOPER& operator&=(const XOPER& str)
		{
			ensure(str.xltype == xltypeStr);
			str_append(str.val.str + 1, str.val.str[0]);

			return *this;
		}
		XOPER& operator&=(xcstr str)
		{
			str_append(str, traits<X>::len(str));

			return *this;
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
		// xltypeErr
		// xltypeFlow
		// xltypeMulti
		
		XOPER(xrw rw, xcol col)
		{
			multi_alloc(rw, col);
		}
		XOPER& resize(xrw rw, xcol col)
		{
			multi_realloc(rw, col);

			return *this;
		}
		xrw rows() const
		{
			return xltype == xltypeMulti ? val.array.rows : (xrw)1;
		}
		xcol columns() const
		{
			return static_cast<xcol>(xltype == xltypeMulti ? val.array.columns : 1);
		}
		auto size() const
		{
			return xltype == xltypeMulti ? rows() * columns() : 1;
		}
		const XOPER* begin() const
		{
			return xltype == xltypeMulti ? static_cast<XOPER*>(val.array.lparray) : this;
		}
		XOPER* begin()
		{
			return xltype == xltypeMulti ? static_cast<XOPER*>(val.array.lparray) : this;
		}
		const XOPER* end() const
		{
			return xltype == xltypeMulti ? static_cast<XOPER*>(val.array.lparray + size()) : this + 1;
		}
		XOPER* end()
		{
			return xltype == xltypeMulti ? static_cast<XOPER*>(val.array.lparray + size()) : this + 1;
		}
		// one-dimensional index
		const XOPER& operator[](xint i) const
		{
			ensure(i < size());

			if (xltype == xltypeMulti) {
				return static_cast<const XOPER&>(val.array.lparray[i]);
			}
			else {
				ensure(i == 0);
				return *this;
			}
		}
		XOPER& operator[](xint i)
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
		// two-dimensional index
		const XOPER& operator()(xrw rw, xcol col) const
		{
			return operator[](columns() * rw + col);
		}
		XOPER& operator()(xrw rw, xcol col)
		{
			return operator[](columns() * rw + col);
		}

		// xltypeMissing
		// xltypeNil
		// xltypeSRef
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

		// xltypeStr
		void str_alloc(size_t n)
		{
			xltype = xltypeStr;
			val.str = (xchar*)malloc((n + 1) * sizeof(xchar));
			// first character is count
			if (val.str) {
				val.str[0] = static_cast<xchar>(n);
			}
		}
		void str_realloc(size_t n)
		{
			ensure(xltype == xltypeStr);
			
			val.str = (xchar*)realloc(val.str, (n + 1) * sizeof(xchar));
			if (val.str) {
				val.str[0] = static_cast<xchar>(n);
			}
		}
		void str_append(xcstr str, size_t n)
		{
			ensure(xltype == xltypeNil || xltype == xltypeStr);

			if (xltype == xltypeNil) {
				operator=(XOPER(str, n));
			}
			else {
				if (n == 0) {
					n = traits<X>::len(str);
				}
				xchar len = val.str[0];
				str_realloc(len + n);
				if (val.str) {
					std::copy(str, str + n, val.str + 1 + len);
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

	using OPER = XOPER<XLOPER>;
	using OPER12 = XOPER<XLOPER12>;
	using OPERX = XOPER<XLOPERX>;

	typedef OPER* LPOPER;
	typedef OPER12* LPOPER12;
	typedef OPERX* LPOPERX;

	// Just like Excel.
	inline xll::OPER operator&(const XLOPER& x, const XLOPER& y)
	{
		return xll::OPER(x) &= y;
	}
	inline xll::OPER12 operator&(const XLOPER12& x, const XLOPER12& y)
	{
		return xll::OPER12(x) &= y;
	}
}
