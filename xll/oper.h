// oper.h - Wrapper for XLOPER
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
	template<class X>  
	class XOPER : public X {
	public:
		using X::xltype;
		using X::val;
		typedef typename traits<X>::xchar xchar;
		typedef typename traits<X>::xint xint;

		XOPER()
		{
			xltype = xltypeMissing;
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
		
		template<class T>
		XOPER& operator=(const T& t)
		{
			operator=(XOPER(t));

			return *this;
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
			// ensure val.str!!!
			val.str[0] = static_cast<xchar>(n);
			traits<X>::cpy(val.str + 1, str, n);
		}
		void free_str()
		{
			//ensure(xltype == xltypeStr);
			free(val.str);
		}
	};


	using OPER = XOPER<XLOPER>;
	using OPER12 = XOPER<XLOPER12>;
	//using OPERX = XOPER<XLOPERX>;
}
