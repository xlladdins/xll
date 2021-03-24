// codec.h - encoder/decoder for XLOPERs
// XLOPER is a variant type num|str|err|multi|...
#pragma once
#include <concepts>
#include <iostream>
#include "traits.h"

namespace xll {


	template<class X, class C = traits<X>::xchar>
	std::basic_ostream<C>& encode(std::basic_ostream<C>& os, const X& x)
	{
		// for xltypeMulti
		static C lb = '{', rb = '}', fs = ',', rs = ';';

		switch (x.xltype & xlbitmask) {
		case xltypeNum:
			os << x.val.num;
			break;
		case xltypeStr:
			os.write(x.val.str + 1, x.val.str[0]);
			break;
		case xltypeErr:
			os << xll_err_desc[x.val.err];
			break;
		case xltypeMulti:
			os << rb;
			for (unsigned i = 0; i < rows(x); ++i) {
				if (i > 0) {
					os << rs;
				}
				for (unsigned j = 0; j < columns(x); ++j) {
					if (j > 0) {
						os << fs;
					}
					os << index(x, i, j);
				}
			}
			os << lb;
			break;
		case xltypeSRef:
			X o;
			int result = traits<X>::Excelv(xlfReftext, &o, 1, &x);
			ensure(result == xlretSuccess and o.val.xltype == xltypeStr);
			os.write(o.val.str + 1, o.val.str[0]);
			traits<X>::Excelv(xlFree, 0, 1, &o);
			break;
		case xltypeRef:
			ensure(!"not implemented");
			break;
		case xltypeInt:
			os << x.val.w;
			break;
		case xltypeBool:
			os << x.val.xbool != 0;
			break;
		default:
			ensure(!"unknown type");
		}

		return os;
	}

	template<class X, class C = traits<X>::xchar>
	inline X decode(std::basic_istream<C>& is)
	{
		C c = 0; // length char at front
		std::basic_ostringstream<C> os;
		os << c;
		os << is;
		std::basic_string<C> s(os.str());
		s[0] = static_cast<C>(s.length());
		X x, o;
		x.xltype = xltypeStr;
		x.val.str = &s[0];
		int result = traits<X>::Excelv(xlfEvaluate, &o, 1, &x);
		ensure(result == xlretSuccess);
		x.xlbit |= xlbitXLFree;

		return x;
	}
}
