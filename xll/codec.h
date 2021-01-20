// codec.h - encoder/decoder for XLOPERs
#pragma once
#include <concepts>
#include "traits.h"

namespace xll {

	// ??? use encoder/decoder to serialize/unserialize
	template<class O, class X> // ouput
	concept encoder {
		static O& operator()(O& o, double num);
		static O& operator()(O& o, traits<X>::cstr str);
		...
			template<class T>
		static O& operator()(O& o, T t);
	};

	template<class E, class O, class X>
	O& encode(O& o, const XOPER<X>& x)
	{
		switch (x.xltype & xlbitmask) {
		case xltypeNum: E(o, x.val.num); break;
		case xltypeStr: E(o, x.val.str); break; // counted string
			...

		}

		return o;
	}

	class decoder {
	}



	// ??? use encoder/decoder to serialize/unserialize
	template<class O, class X> // ouput
	struct encoder {
		static O& operator()(O& o, double num);
		static O& operator()(O& o, traits<X>::cstr str);
		...
			template<class T>
		static O& operator()(O& o, T t);
	};

	template<class E, class O, class X>
	O& encode(O& o, const XOPER<X>& x)
	{
		switch (x.xltype & xlbitmask) {
		case xltypeNum: E(o, x.val.num); break;
		case xltypeStr: E(o, x.val.str); break; // counted string
			...

		}

		return o;
	}

	class decoder {
	}



}
