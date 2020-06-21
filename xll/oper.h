// oper.h - Wrapper for XLOPER
#pragma once
#include <utility>
#include <Windows.h>
#include "XLCALL.H"

#pragma warning(disable: 4996)

namespace xll {

	// predefined XLOPERs
	inline const XLOPER Missing 
		= { .xltype = xltypeMissing };
	inline const XLOPER12 Missing12 
		= { .xltype = xltypeMissing };
	inline const XLOPER Nil 
		= { .xltype = xltypeNil };
	inline const XLOPER12 Nil12 
		= { .xltype = xltypeNil };

	inline const XLOPER ErrNull 
		= { .val = { .err = xlerrNull }, .xltype = xltypeErr };
	inline const XLOPER12 ErrNull12 
		= { .val = { .err = xlerrNull }, .xltype = xltypeErr };
	inline const XLOPER ErrDiv0
		= { .val = { .err = xlerrDiv0 }, .xltype = xltypeErr };
	inline const XLOPER12 ErrDiv012
		= { .val = { .err = xlerrDiv0 }, .xltype = xltypeErr };
	inline const XLOPER ErrValue
		= { .val = { .err = xlerrValue }, .xltype = xltypeErr };
	inline const XLOPER12 ErrValue12
		= { .val = { .err = xlerrValue }, .xltype = xltypeErr };
	inline const XLOPER ErrRef
		= { .val = { .err = xlerrRef }, .xltype = xltypeErr };
	inline const XLOPER12 ErrRef12
		= { .val = { .err = xlerrRef }, .xltype = xltypeErr };
	inline const XLOPER ErrName
		= { .val = { .err = xlerrName }, .xltype = xltypeErr };
	inline const XLOPER12 ErrName12
		= { .val = { .err = xlerrName }, .xltype = xltypeErr };
	inline const XLOPER ErrNum
		= { .val = { .err = xlerrNum }, .xltype = xltypeErr };
	inline const XLOPER12 ErrNum12
		= { .val = { .err = xlerrNum }, .xltype = xltypeErr };
	inline const XLOPER ErrNA
		= { .val = { .err = xlerrNA }, .xltype = xltypeErr };
	inline const XLOPER12 ErrNA12
		= { .val = { .err = xlerrNA }, .xltype = xltypeErr };

	// XLOPER/XLOPER12 traits
	// ???Put in traits.h???
	template<class X>
	struct traits {
	};
	template<>
	struct traits<XLOPER> {
		typedef CHAR xchar;
		typedef short int xint;
		static size_t len(const xchar* s)
		{
			return strlen(s);
		}
		static xchar* cpy(xchar* dest, const xchar* src, size_t n)
		{
			return strncpy(dest, src, n);
		}
	};
	template<>
	struct traits<XLOPER12> {
		typedef XCHAR xchar;
		typedef int xint;
		static size_t len(const xchar* s)
		{
			return wcslen(s);
		}
		static xchar* cpy(xchar* dest, const xchar* src, size_t n)
		{
			return wcsncpy(dest, src, n);
		}
	};

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
			else {
				xltype = o.xltype;
				val = o.val;
			}
		}
		OPERX(OPERX&& o) noexcept
		{
			xltype = o.xltype;
			std::exchange(val, o.val);
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
				std::exchange(val, o.val);
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
		/*
		template<class T>
		OPERX& operator=(const T& t)
		{
			return OPERX(t);
		}
		*/
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
