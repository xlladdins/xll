// oper.h - Wrapper for XLOPER
#pragma once
#include <utility>
#include "xloper.h"

#pragma warning(disable: 4996)

namespace xll {

	template<class X>
	class OPERX : public X {
	public:
		using X::xltype;
		using X::val;
		typedef typename traits<X>::xchar xchar;
		typedef typename traits<X>::xint xint;

		OPERX()
		{
			xltype = xltypeMissing;
		}
		OPERX(const OPERX& o)
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
		OPERX(OPERX&& o) noexcept
		{
			xltype = o.xltype;
			val = o.val;
			o.xltype = xltypeMissing;
		}
		OPERX& operator=(const OPERX& o)
		{
			if (this != &o) {
				this->~OPERX();
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
		OPERX& operator=(OPERX&& o) noexcept
		{
			if (this != &o) {
				xltype = o.xltype;
				val = o.val;
				o.xltype = xltypeMissing;
			}

			return *this;
		}
		~OPERX()
		{
			if (xltype == xltypeStr) {
				free_str();
			}
			// else if (xltype == xltypeMulti) { }
		}
		
		template<class T>
		OPERX& operator=(const T& t)
		{
			operator=(OPERX(t));

			return *this;
		}
		
		// floating point number
		explicit OPERX(double num)
		{
			xltype = xltypeNum;
			val.num = num;
		}

		// Counted character string
		explicit OPERX(const xchar* str)
		{
			alloc_str(str, 0);
		}		
		// Construct from string literal
		template<size_t N>
		OPERX(const xchar (&str)[N])
		{
			alloc_str(&str[0], N);
		}

		explicit OPERX(bool xbool)
		{
			xltype = xltypeBool;
			val.xbool = xbool;
		}

		// xltypeRef
		// xltypeErr
		// xltypeFlow
		// xltypeMulti
		// xltypeMissing
		// xltypeNil
		// xltypeSRef

		// Excel usually converts this to num.
		explicit OPERX(int w)
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
			val.str[0] = static_cast<xchar>(n);
			traits<X>::cpy(val.str + 1, str, n);
		}
		void free_str()
		{
			//ensure(xltype == xltypeStr);
			free(val.str);
		}
	};


	using OPER = OPERX<XLOPER>;
	using OPER12 = OPERX<XLOPER12>;

}
