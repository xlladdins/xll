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
	/// It is a variant type that can be either a number, string,
	/// error, range, integer, or boolean. 
	template<class X>  
	class XOPER : public X {
	public:
		using X::xltype;
		using X::val;
		typedef typename traits<X>::xchar xchar;
		typedef typename traits<X>::xcstr xcstr;
		typedef typename traits<X>::xint xint;
		typedef typename traits<X>::xrw xrw;
		typedef typename traits<X>::xcol xcol;

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
				str_cpy(x.val.str + 1, x.val.str[0]);
			}
			// else if (x.xltype == xltypeMulti) { ... }
			else {
				xltype = x.xltype;
				val = x.val;
			}
		}
		XOPER(XOPER&& o) noexcept
		{
			xltype = o.xltype;
			val = o.val;
			o.xltype = xltypeMissing;
		}
		XOPER& operator=(const XOPER& o)
		{
			return operator=(static_cast<const X&>(o));
		}
		XOPER& operator=(const X& x)
		{
			if (this != &x) {
				oper_free();
				if (x.xltype == xltypeStr) {
					str_cpy(x.val.str + 1, x.val.str[0]);
				}
				// else if (x.xltype == xltypeMulti) { }
				else {
					xltype = x.xltype;
					val = x.val;
				}
			}

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
		~XOPER()
		{
			oper_free();
		}

		/*
		template<class T>
		XOPER& operator=(const T& t)
		{
			operator=(XOPER(t));

			return *this;
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

		// NULL terminated string
		explicit XOPER(xcstr str)
		{
			str_cpy(str, 0);
		}
		// Construct from string literal
		template<size_t N>
		XOPER(const xchar (&str)[N])
		{
			str_cpy(str, N - 1);
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
			str_append(str, 0);

			return *this;
		}

		explicit XOPER(bool xbool)
		{
			xltype = xltypeBool;
			val.xbool = xbool;
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
			alloc_multi(rw, col);
		}
		
		// xltypeMissing
		// xltypeNil
		// xltypeSRef

		// Excel usually converts this to num.
		explicit XOPER(int w)
		{
			xltype = xltypeInt;
			val.w = static_cast<xint>(w);
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
				free_str();
			}
			else if (xltype == xltypeMulti) {
				free_multi();
			}
			else {
				xltype = xltypeNil;
			}
		}
		void str_alloc(size_t n)
		{
			xltype = xltypeStr;
			val.str = (xchar*)malloc((n + 1) * sizeof(xchar));
			// first character is count
			if (val.str) {
				val.str[0] = static_cast<xchar>(n);
			}
		}
		void str_cpy(xcstr str, size_t n)
		{
			if (n == 0) {
				n = traits<X>::len(str);
			}
			str_alloc(n);
			if (val.str) {
				traits<X>::cpy(val.str + 1, str, n);
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
			if (xltype == xltypeNil || xltype == xltypeMissing) {
				str_cpy(str, n);
			}
			else {
				if (n == 0) {
					n = traits<X>::len(str);
				}
				realloc_str(val.str[0] + n);
				if (val.str) {
					traits<X>::cpy(val.str + val.str[0] + 1, str, n);
				}
			}
		}
		void free_str()
		{
			//ensure(xltype == xltypeStr);
			free(val.str);
			xltype = xltypeNil;
		}
		void alloc_multi(xrw rw, xcol col)
		{
			xltype = xltypeMulti;
			val.array.rw = rw;
			val.array.col = col;
			val.array.lparray = (X*)malloc(sizeof(X) + rw * col * sizeof(X*));
			if (val.array.lparray) {
				for (size_t i = 0; i < rw * col; ++i) {
					val.array.lparray = XOPER{};
				}
			}
			xltype = xltypeNil;
		}
		void free_multi()
		{

		}
	};

	using OPER = XOPER<XLOPER>;
	using OPER12 = XOPER<XLOPER12>;
	using OPERX = XOPER<XLOPERX>;
}
