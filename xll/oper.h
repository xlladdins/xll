// oper.h - Wrapper for XLOPER
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <algorithm>
#include <concepts>
#include <iostream>
#include <utility>
#include "utf8.h"
#include "xloper.h"
#include "sref.h"

#pragma warning(disable: 4996)

namespace xll {

	/// <summary>
	///  Wrapper for XLOPER/XLOPER12 structs.
	/// </summary>
	/// <typeparam name="X">
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
		typedef typename traits<X>::xbool xbool;
		typedef typename traits<X>::xint xint;
		typedef typename traits<X>::xrw xrw;
		typedef typename traits<X>::xcol xcol;
		typedef typename traits<X>::xint xint;
		typedef typename traits<X>::xref xref;

	public:
		using X::xltype;
		using X::val;

		void swap(XOPER& x) noexcept
		{
			using std::swap;

			swap(xltype, x.xltype);
			swap(val, x.val);
		}

		//
		// constructors
		//
		XOPER()
		{
			xltype = xltypeNil;
		}
		XOPER(const X& x)
		{
			if (x.xltype & xltypeScalar) {
				xltype = x.xltype;
				val = x.val;
			}
			else if (x.xltype & xltypeStr) {
				str_alloc(x.val.str + 1, x.val.str[0]);
			}
			else if (x.xltype & xltypeMulti) {
				multi_alloc(xll::rows(x), xll::columns(x));
				std::copy(xll::begin(x), xll::end(x), begin());
			}
			else if (x.xltype & xltypeRef) {
				ref_alloc(x.val.mref.lpmref->count);
				val.mref.idSheet = x.val.mref.idSheet;
				std::copy(ref_begin(x), ref_end(x), val.mref.lpmref->reftbl);
			}
			else if (x.xltype & xltypeBigData) {
				throw std::runtime_error("XOPER: xltypeBigData not supported yet");
			}
			else {
				throw std::runtime_error("XOPER: unknown xltype");
			}
		}
		XOPER(const XOPER& o)
			: XOPER(static_cast<const X&>(o))
		{ }

		XOPER(XOPER&& o) noexcept
			: XOPER()
		{
			xltype = std::exchange(o.xltype, xltype);
			val = std::exchange(o.val, val);
		}
		//
		// Assignment
		//
		XOPER& operator=(const XOPER& o)
		{
			return *this = XOPER(o);
		}
		XOPER& operator=(XOPER&& o) noexcept
		{
			swap(o);

			return *this;
		}
		~XOPER()
		{
			oper_free();
		}

		// Does not involve memory needing to be freed.
		bool is_scalar() const
		{
			return xltype & xltypeScalar;
		}
		
		explicit operator bool() const
		{
			switch (xltype & xlbitmask) {
			case xltypeNum:
				return val.num != 0;
			case xltypeStr:
				return val.str[0] != 0;
			case xltypeBool:
				return val.xbool != 0;
			case xltypeMulti:
				return std::any_of(begin(), end(), [](const XOPER& x) { return !!x; });
			case xltypeInt:
				return val.w != 0;
			case xltypeBigData:
				return val.bigdata.cbData != 0;
			case xltypeMissing: case xltypeSRef: case xltypeRef:
				return true;
			}

			return false; // xltypeErr, xltypeNil
		}
		/**/
		// IEEE 64-bit floating point number
		template<class T> 
			requires std::is_convertible_v<T,double>
		XOPER(T num)
		{
			xltype = xltypeNum;
			val.num = static_cast<double>(num);
		}
		template<class T>
			requires std::is_convertible_v<T, double>
		XOPER& operator=(T num)
		{
			oper_free();

			xltype = xltypeNum;
			val.num = static_cast<double>(num);

			return *this;
		}
		template<class T>
			requires std::is_convertible_v<T, double>
		bool operator==(T num) const
		{
			return xltype == xltypeNum && val.num == static_cast<double>(num);
		}
		bool is_num() const
		{
			return xltype == xltypeNum;
		}
		const double& as_num() const
		{
			ensure(is_num());

			return val.num;
		}
		double& as_num()
		{
			ensure(is_num());

			return val.num;
		}

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
			str_alloc(traits<X>::cvt(str), -1);
		}
		// Construct from string literal
		template<size_t N>
		XOPER(cstrx (&str)[N])
		{
			str_alloc(traits<X>::cvt(str), -1);
		}
		bool operator==(xcstr str) const
		{
			if (!(xltype & xltypeStr)) {
				return false;
			}

			unsigned n = traits<X>::len(str);
			//ensure(n < static_cast<size_t>(std::numeric_limits<xchar>::max()));
			
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
			str_alloc(traits<X>::cvt(str), -1);

			return *this;
		}
		bool is_str() const
		{
			return xltype & xltypeStr;
		}
		// reference to counted string
		const xcstr& as_str() const
		{
			ensure(is_str());

			return val.str;
		}
		// xcstr& as_str() not provided
		std::string to_string() const
		{
			if (xltype & xltypeNil) {
				return "";
			}

			ensure(is_str());

			if constexpr (std::is_same_v<X, XLOPER>) {
				return std::string(val.str + 1, val.str[0]);
			}
			else {
				return utf8::wcstostring(val.str + 1, val.str[0]);
			}
		}
		// string building
		XOPER& append(const X& x)
		{
			if (x.xltype == xltypeNil) {
				return *this;
			}

			ensure(x.xltype & xltypeStr);
			
			if (xltype & xlbitXLFree) {
				XOPER tmp(*this);
				oper_free();
				swap(tmp);
			}
			str_append(x.val.str + 1, x.val.str[0]);

			return *this;
		}
		XOPER& append(const X_& x)
		{
			if (x.xltype == xltypeNil) {
				return *this;
			}

			ensure(x.is_str());
			str_append(traits<X>::cvt(x.val.str + 1, x.val.str[0]), (unsigned)-1);

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
		XOPER& append(xcstr str = nullptr)
		{
			if (!str) {
				// null terminate a string
				xchar null[1] = { 0 };
				str_append(null, 1);
			}
			else {
				str_append(str, traits<X>::len(str));
			}

			return *this;
		}
		XOPER& append(cstrx str)
		{
			str_append(traits<X>::cvt(str), -1);

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
			return xltype == xltypeBool and val.xbool == static_cast<typename traits<X>::xbool>(xbool);
		}
		XOPER operator=(bool xbool)
		{
			oper_free();
			operator=(XOPER(xbool));

			return *this;
		}
		bool is_bool() const
		{
			return xltype == xltypeBool;
		}
		xbool& as_bool()
		{
			ensure(is_bool());

			return val.xbool;
		}
		const xbool& as_bool() const
		{
			ensure(is_bool());

			return val.xbool;
		}

		// xltypeErr - predifined as ErrXXX
#define XLL_ERR_ENUM(a, b, c) a = xlerr##a,
		enum class Err {
			XLL_ERR_TYPE(XLL_ERR_ENUM)
		};
#undef XLL_ERR_ENUM
		XOPER(enum Err err)
		{
			xltype = xltypeErr;
			val.err = static_cast<WORD>(err);
		}
		XOPER& operator=(enum Err err)
		{
			oper_free();
			xltype = xltypeErr;
			val.err = static_cast<WORD>(err);

			return *this;
		}
		bool is_err() const
		{
			return xltype == xltypeErr;
		}
		// xltypeFlow - not used

		// xltypeMulti
		bool is_multi() const
		{
			return xltype & xltypeMulti;
		}
		XOPER(unsigned rw, unsigned col)
		{
			multi_alloc(rw, col);
		}
		// XOPER({XOPER(a), ...}). Use resize if needed.
		explicit XOPER(std::initializer_list<XOPER> x)
		{
			//ensure(x.size() <= std::numeric_limits<xcol>::max());

			multi_alloc(1, static_cast<xcol>(x.size()));
			std::copy(x.begin(), x.end(), begin());
		}
		XOPER& resize(unsigned rw, unsigned col)
		{
			multi_realloc(rw, col);

			return *this;
		}
		unsigned rows() const
		{
			return xll::rows(*this);
		}
		unsigned columns() const
		{
			return xll::columns(*this);
		}
		unsigned size() const
		{
			return xll::size(*this);
		}
		// STL friendly
		XOPER* begin()
		{
			return xll::begin(*this);
		}
		const XOPER* begin() const
		{
			return xll::begin(*this);
		}
		XOPER* end()
		{
			return xll::end(*this);
		}
		const XOPER* end() const
		{
			return xll::end(*this);
		}
		// one-dimensional index
		XOPER& operator[](unsigned i)
		{
			return xll::index(*this, i);
		}
		const XOPER& operator[](unsigned i) const
		{
			return xll::index(*this, i);
		}
		// two-dimensional index
		XOPER& operator()(unsigned rw, unsigned col)
		{
			return xll::index(*this, rw, col);
		}
		const XOPER& operator()(unsigned rw, unsigned col) const
		{
			return xll::index(*this, rw, col);
		}
		const XOPER& append_bottom(const X& x)
		{
			if (xltype == xltypeNil) {
				operator=(x);

				return *this;
			}
			// make a copy if memory overlap
			const XOPER& o = overlap(x) ? XOPER(x) : x;
			if (xltype == xltypeMulti) {
				ensure(columns() == xll::columns(x));
				multi_realloc(rows() + xll::rows(x), columns());
				std::copy(o.begin(), o.end(), end() - o.size());
			}
			else {
				XOPER tmp(1, 1);
				swap(tmp);
				operator[](0) = tmp;

				return append_bottom(o);
			}

			return *this;
		}
		// append_top
		// append_right
		// append_left

		// xltypeMissing - predefined as Missing
		// xltypeNil - predefined as Nil

		// xltypeSRef - reference to a single range
		XOPER(xrw row, xcol col, xrw height, xcol width)
		{
			xltype = xltypeSRef;
			val.sref.count = 1;
			val.sref.ref = xref(row, col, height, width);
		}
		XOPER(const xref& ref)
		{
			xltype = xltypeSRef;
			val.sref.count = 1;
			val.sref.ref = ref;
		}
		bool is_sref() const
		{
			return xltype == xltypeSRef;
		}
		const xref& as_sref() const
		{
			return val.sref.ref;
		}
		xref& as_sref()
		{
			return val.sref.ref;
		}

		// xltypeRef - reference to multiple refs
		XOPER(const std::initializer_list<xref>& ref)
		{
			xltype = xltypeRef;
			ref_alloc(static_cast<WORD>(ref.size()));
			std::copy(ref.begin(), ref.end(), val.mref.lpmref->reftbl);
		}
		bool is_ref() const
		{
			return xltype & xltypeRef;
		}
		// size() returns number of refs
		const xref* as_ref() const
		{
			ensure(is_ref());

			return val.mref.lpmref->reftbl;
		}
		xref* as_ref()
		{
			ensure(is_ref());

			return val.mref.lpmref->reftbl;
		}

		// xltypeInt. Excel usually converts this to num.
		bool is_int() const
		{
			return xltype & xltypeInt;
		}
		// ints get converted to double, just like Excel
		const xint& as_int() const
		{
			ensure(is_int());

			return val.w;
		}
		xint& as_int()
		{
			ensure(is_int());

			return val.w;
		}

	private:
		// true if memory overlaps with x
		bool overlap(const X& x) const
		{
			return (begin() <= xll::begin(x) and xll::begin(x) < end()) 
				|| (begin() < xll::end(x)    and xll::end(x)  <= end());
		}

		void oper_free()
		{
			if (xltype & xlbitXLFree) {
				X* px[1];
				px[0] = this;
				ensure (xlretSuccess ==  traits<X>::Excelv(xlFree, 0, 1, px));
			}
			else if (xltype == xltypeStr) {
				str_free();
			}
			else if (xltype == xltypeMulti) {
				multi_free();
			}
			else if (xltype == xltypeRef) {
				ref_free();
			}
			
			xltype = xltypeNil;
		}

		// allocate unless counted
		void str_alloc(xcstr str, int n)
		{
			xltype = xltypeStr;

			if (n == -1) {
				val.str = const_cast<xchar*>(str); // move
			}
			else {
				val.str = (xchar*)malloc(((size_t)n + 1) * sizeof(xchar));
				// first character is count
				if (val.str) {
					memcpy_s(val.str + 1, n * sizeof(xchar), str, n * sizeof(xchar));
					// ensure (n <= ...MAX);
					val.str[0] = static_cast<xchar>(n);
				}
			}
		}
		void str_append(xcstr str, int n)
		{
			if (xltype == xltypeNil) {
				str_alloc(str, n);
			}
			else if (xltype == (xltypeStr & xlbitXLFree)) {
				*this = XExcel<X>(xlfConcatenate, *this, XOPER<X>(str));
			}
			else {
				ensure(xltype == xltypeStr);
				bool counted = false;
				if (n == -1) {
					counted = true;
					n = str[0];
				}
				if (n == 0) {
					n = traits<X>::len(str);
				}

				xchar* tmp = (xchar*)realloc(val.str, ((size_t)val.str[0] + n + 1) * sizeof(xchar));
				if (!tmp) {
					throw std::bad_alloc{};
				}
				// ensure(tmp);
				val.str = tmp;
				memcpy_s(val.str + 1 + val.str[0], n * sizeof(xchar), str + counted, n * sizeof(xchar));
				val.str[0] = static_cast<xchar>(val.str[0] + n);

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
		void multi_alloc(unsigned r, unsigned c)
		{
			ensure(r <= (unsigned)(std::numeric_limits<xrw>::max)());
			ensure(c <= (unsigned)(std::numeric_limits<xcol>::max)());

			if (r * c == 0) {
				xltype = XOPER().xltype;

				return;
			}

			xltype = xltypeMulti;
			val.array.rows = static_cast<xrw>(r);
			val.array.columns = static_cast<xcol>(c);
			val.array.lparray = static_cast<X*>(malloc(size() * sizeof(X)));
			if (val.array.lparray) {
				std::fill(begin(), end(), XOPER{});
			}
		}
		void multi_realloc(unsigned r, unsigned c)
		{
			ensure(r <= (unsigned)(std::numeric_limits<xrw>::max)());
			ensure(c <= (unsigned)(std::numeric_limits<xcol>::max)());

			if (!(xltype & xltypeMulti)) {
				oper_free();
				multi_alloc(r, c);

				return;
			}

			if (r * c == 0) {
				multi_free();
				operator=(XOPER{});

				return;
			}
	
			// current size
			auto n = size();
			val.array.rows = static_cast<xrw>(r);
			val.array.columns = static_cast<xcol>(c);
			if (n > size()) {
				std::for_each(end(), begin() + n, [](auto& o) { o.oper_free(); });
				X* tmp = (X*)realloc(val.array.lparray, size() * sizeof(X));
				ensure(tmp);
				val.array.lparray = tmp;
			}
			else if (n < size()) {
				X* tmp = (X*)realloc(val.array.lparray, size() * sizeof(X));
				ensure(tmp);
				val.array.lparray = tmp;
				std::fill(begin() + n, end(), XOPER{});
			}
		}
		void multi_free()
		{
			ensure(xltype == xltypeMulti);
			std::for_each(begin(), end(), [](auto& o) { o.oper_free(); });
			free(val.array.lparray);

			xltype = xltypeNil;
		}
		void ref_alloc(WORD count)
		{
			using xref = typename traits<X>::xref;
			using xmref = typename traits<X>::xmref;

			xltype = xltypeRef;
			val.mref.lpmref = (xmref*)malloc(sizeof(xmref) + count * sizeof(xref));
			ensure(val.mref.lpmref);
			val.mref.lpmref->count = count;
		}
		void ref_realloc(WORD count)
		{
			using xref = typename traits<X>::xref;
			using xmref = typename traits<X>::xmref;

			ensure(xltype == xltypeRef);
			void* tmp = realloc(val.mref.lpmref, sizeof(xmref) + count * sizeof(xref));
			if (tmp) {
				val.mref.lpmref = (xmref*)tmp;
				val.mref.lpmref->count = count;
			}
		}
		void ref_free()
		{
			ensure(xltype == xltypeRef);
			free(val.mref.lpmref);

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

// Just like Excel concatenation
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

/*
inline auto& operator<<(std::ostream& os, const xll::OPER& o)
{
	switch (o.xltype) {
	case xltypeNum:
		os << o.val.num;
		break;
	case xltypeStr:
		os.write(o.val.str + 1, o.val.str[0]);
		break;
	case xltypeBool:
		os << static_cast<bool>(o.val.xbool);
		break;
	case xltypeErr:
		os << xll_err_str[o.val.err];
		break;
	case xltypeMulti: // "{a,b...;c,d,...}"
		os << "{";
		for (unsigned r = 0; r < xll::rows(o); ++r) {
			for (unsigned c = 0; c < xll::columns(o)) {
				if (c > 0)
					os << ",";
				else if (r > 0)
					os << ";";
				os << o(r, c);
			}
		}
		os << "}";
		break;
	case xltypeInt:
		os << o.val.w;
		break;
	default:
		break;
	}

	return os;
}
*/