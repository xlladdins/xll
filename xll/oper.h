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
		typedef typename traits<X>::xint xint;

		XOPER()
		{
			xltype = xltypeNil;
		}
		XOPER(const XOPER& o)
		{
			if (o.xltype == xltypeStr) {
				alloc_str(o.val.str + 1, o.val.str[0]);
			}
			// else if (o.xltype == xltypeMulti) { ... }
			else {
				xltype = o.xltype;
				val = o.val;
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
			if (this != &o) {
				this->~XOPER();
				if (o.xltype == xltypeStr) {
					alloc_str(o.val.str + 1, o.val.str[0]);
				}
				// else if (o.xltype == xltypeMulti) { }
				else {
					xltype = o.xltype;
					val = o.val;
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
			if (xltype == xltypeStr) {
				free_str();
			}
			// else if (xltype == xltypeMulti) { }
		}
		XOPER(const X& x)
		{
			if (x.xltype == xltypeStr) {
				alloc_str(x.val.str + 1, x.val.str[0]);
			}
			// else if (x.xltype == xltypeMulti) ...
			else {
				xltype = x.xltype;
				val = x.val;
			}
		}
		XOPER& operator=(const X& x)
		{
			return operator=(XOPER(x));
		}
		
		template<class T>
		XOPER& operator=(const T& t)
		{
			operator=(XOPER(t));

			return *this;
		}

		bool operator==(const XOPER& o) const
		{
			if (xltype != o.xltype) {
				return false;
			}

			switch (xltype) {
			case xltypeNum:
				return val.num == o.val.num;
			case xltypeStr:
				return val.str[0] != o.val.str[0]
					? false
				    : 0 == traits<X>::cmp(val.str + 1, o.val.str + 1, val.str[0]);
			//case xltypeMulti:
			case xltypeInt:
				return val.w == o.val.w;
			case xltypeBool:
				return val.xbool == o.val.xbool;
			default:
				return false;
			}
		}
		bool operator<(const XOPER& o) const
		{
			if (xltype != o.xltype) {
				return false;
			}

			switch (xltype) {
			case xltypeNum:
				return val.num < o.val.num;
			case xltypeStr:
				return val.str[0] != o.val.str[0]
					? val.str[0] < o.val.str[0]
					: traits<X>::cmp(val.str + 1, o.val.str + 1, val.str[0]) < 0;
				//case xltypeMulti:
			case xltypeInt:
				return val.w < o.val.w;
			case xltypeBool:
				return val.xbool < o.val.xbool;
			default:
				return false;
			}
		}

		// floating point number
		explicit XOPER(double num)
		{
			xltype = xltypeNum;
			val.num = num;
		}

		// Counted character string
		explicit XOPER(const xchar* str)
		{
			alloc_str(str, 0);
		}		
		// Construct from string literal
		template<size_t N>
		XOPER(const xchar (&str)[N])
		{
			alloc_str(&str[0], N);
		}
		// Like the Excel ampersand operator
		XOPER& operator&=(const XOPER& str)
		{
			//ensure(str.xltype == xltypeStr);

			append(str.val.str + 1, str.val.str[0]);

			return *this;
		}
		XOPER& operator&=(const xchar* str)
		{
			append(str, 0);

			return *this;
		}

		explicit XOPER(bool xbool)
		{
			xltype = xltypeBool;
			val.xbool = xbool;
		}

		// xltypeRef
		// xltypeErr
		// xltypeFlow
		// xltypeMulti
		/*
		OPER(xrw rw, xcol col)
		{

		}
		*/
		// xltypeMissing
		// xltypeNil
		// xltypeSRef

		// Excel usually converts this to num.
		explicit XOPER(int w)
		{
			xltype = xltypeInt;
			val.w = static_cast<xint>(w);
		}


	private:
		void alloc_str(const xchar* str, size_t n)
		{
			if (n == 0) {
				n = traits<X>::len(str);
			}
			xltype = xltypeStr;
			val.str = (xchar*)malloc((n + 1) * sizeof(xchar));
			// ensure (val.str);
			if (val.str != nullptr) {
				traits<X>::cpy(val.str + 1, str, n);
				val.str[0] = static_cast<xchar>(n);
			}
		}
		void append(const xchar* str, size_t n)
		{
			//ensure(xltype == xltypeStr);
			if (n == 0) {
				n = traits<X>::len(str);
			}
			if (xltype == xltypeNil) {
				alloc_str(str, n);
			}
			else {
				val.str = (xchar*)realloc(val.str, (val.str[0] + n + 1) * sizeof(xchar));
				// ensure (val.str);
				if (val.str != nullptr) {
					traits<X>::cpy(val.str + val.str[0] + 1, str, n);
					val.str[0] = static_cast<xchar>(val.str[0] + n);
				}
			}
		}
		void free_str()
		{
			//ensure(xltype == xltypeStr);
			free(val.str);
		}
	};

	using OPER = XOPER<XLOPER>;
	using OPER12 = XOPER<XLOPER12>;
	using OPERX = XOPER<XLOPERX>;
}
