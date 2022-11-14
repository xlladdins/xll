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
#include "error.h"

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
	class XOPER : public X {
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

		// xltype without xlbitXXX flags
		[[nodiscard]] decltype(xltype) type() const
		{
			return xltype & xlbitmask;
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
			if (x.xltype == xltypeBigData) {
				xltype = x.xltype;
				val.bigdata = x.val.bigdata;
			}
			else if (x.xltype & xltypeScalar) {
				xltype = (x.xltype & xltypeScalar);
				val = x.val;
			}
			else if (x.xltype & xltypeStr) {
				str_alloc(x.val.str + 1, x.val.str[0]);
			}
			else if (x.xltype & xltypeMulti) {
				multi_alloc(xll::rows(x), xll::columns(x));
				std::copy(xll::begin(x), xll::end(x), begin());
			}
			/*
			else if (x.xltype & xltypeRef) {
				ref_alloc(x.val.mref.lpmref->count);
				val.mref.idSheet = x.val.mref.idSheet;
				std::copy(ref_begin(x), ref_end(x), val.mref.lpmref->reftbl);
			}
			*/
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
			swap(o);
			/*
			xltype = std::exchange(o.xltype, xltype);
			val = std::exchange(o.val, val);
			*/
		}
		// convert from other type of OPER
		XOPER(const X_& x)
		{
			xltype = static_cast<WORD>(x.xltype);
			switch (x.xltype & xlbitmask) {
			case xltypeNum:
				val.num = x.val.num;
				break;
			case xltypeStr:
				str_alloc(traits<X>::cvt(x.val.str + 1, x.val.str[0]), -1);
				break;
			case xltypeBool:
				val.xbool = static_cast<typename traits<X>::xbool>(x.val.xbool);
				break;
			//case xltypeRef: //!!! not implemented
			case xltypeErr:
				val.err = static_cast<typename traits<X>::xerr>(x.val.err);
				break;
			case xltypeMulti:
				multi_alloc(x.val.array.rows, x.val.array.columns);
				for (unsigned i = 0; i < size(); ++i)
					val.array.lparray[i] = XOPER(x.val.array.lparray[i]);
				break;
			case xltypeSRef:
				val.sref.count = x.val.sref.count;
				val.sref.ref.rwFirst = static_cast<typename traits<X>::xrw>(x.val.sref.ref.rwFirst);
				val.sref.ref.rwLast = static_cast<typename traits<X>::xrw>(x.val.sref.ref.rwLast);
				val.sref.ref.colFirst = static_cast<typename traits<X>::xcol>(x.val.sref.ref.colFirst);
				val.sref.ref.colLast = static_cast<typename traits<X>::xcol>(x.val.sref.ref.colLast);
				break;
			case xltypeInt:
				val.w = static_cast<typename traits<X>::xint>(x.val.w);
				break;
			default:
				xltype = xltypeErr;
				val.err = xlerrNA;
			}
		}
		//
		// Assignment
		//
		XOPER& operator=(const XOPER& o)
		{
			operator=(static_cast<const X&>(o));

			return *this;
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
		[[nodiscard]] bool is_scalar() const
		{
			return xltype & xltypeScalar;
		}
		
		explicit operator bool() const
		{
			switch (type()) {
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
			case xltypeSRef: case xltypeRef:
				return true;
			}

			return false; // xltypeErr, xltypeNil, xltypeMissing
		}
		
		// IEEE 64-bit floating point number
		[[nodiscard]] bool is_num() const
		{
			return type() == xltypeNum;
		}
		template<class T>
			requires std::is_arithmetic_v<T>
		XOPER(T num)
		{
			xltype = xltypeNum;
			val.num = static_cast<double>(num);
		}
		template<class T>
			requires std::is_arithmetic_v<T>
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
		double& as_num()
		{
			return val.num;
		}
		double as_num() const
		{
			return val.num;
		}

		// xltypeStr
		[[nodiscard]] bool is_str() const
		{
			return type() == xltypeStr;
		}

		// Use, e.g., OPER("", 255) to allocate a string buffer
		XOPER(xcstr str, int n)
		{
			str_alloc(str, n);
		}
		// xltypeStr from string literal
		template<size_t N>
		XOPER(xcstr(&str)[N])
			: XOPER(str, N)
		{
		}
		// xltypeStr from NULL terminated string
		/*explicit*/ XOPER(xcstr str)
			: XOPER(str, traits<X>::len(str))
		{ }

		// convert to type appropriate for X
		XOPER(cstrx str, int len)
		{
			xltype = xltypeStr;
			val.str = traits<X>::cvt(str, len);
		}
		// Construct from string literal
		template<size_t N>
		XOPER(cstrx(&str)[N])
			: XOPER(str, N)
		{ }
		// convert to type appropriate for X
		/*explicit*/ XOPER(cstrx str)
			: XOPER(str, traits<X_>::len(str))
		{ }

		bool operator==(xcstr str) const
		{
			if (type() != xltypeStr) {
				return false;
			}

			unsigned n = traits<X>::len(str);
			ensure(n <= traits<X>::charmax);
			
			if (val.str[0] != static_cast<xchar>(n))
				return false;

			return 0 == traits<X>::cmp(str, val.str + 1, n);
		}
		XOPER operator=(xcstr str)
		{
			oper_free();
			str_alloc(str, traits<X>::len(str));

			return *this;
		}

		XOPER operator=(cstrx str)
		{
			oper_free();

			xltype = xltypeStr;
			val.str = traits<X>::cvt(str, -1);

			return *this;
		}
		//??? is this a good idea
		bool operator==(cstrx str) const
		{
			return *this == XOPER(str);
		}

		// string building
		XOPER& append(xcstr str, int n)
		{
			str_append(str, n);

			return *this;
		}
		XOPER& append(xcstr str)
		{
			return append(str, traits<X>::len(str));
		}
		XOPER& append(cstrx str)
		{
			return append(traits<X>::cvt(str), -1);
		}
		XOPER& append(const X& x)
		{
			if (xltype == xltypeNil) {
				operator=(x);
			}
			else if (x.xltype == xltypeNum) {
				static const XOPER<X> xGeneral("General");
				XOPER<X> xNum = XExcel<X>(xlfText, x, xGeneral);
				append(xNum.val.str + 1, xNum.val.str[0]);
			}
			else if (x.xltype == xltypeStr) {
				append(x.val.str + 1, x.val.str[0]);
			}
			else {
				operator=(XErrValue<X>);
			}

			return *this;
		}
		XOPER& append(const X_& x)
		{
			return append(XOPER(x));
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
		XOPER& operator&=(xcstr str)
		{
			return append(str);
		}
		XOPER& operator&=(cstrx str)
		{
			return append(str);
		}
		// reference to counted string
		const xcstr& as_str() const
		{
			ensure(is_str());

			return val.str;
		}
		// pointer to null terminated string
		xcstr as_cstr() const
		{
			ensure(is_str());

			return val.str + 1;
		}
		// cstrx& as_str() not provided

		std::string to_string() const
		{
			if (xltype == xltypeNil) {
				return "";
			}

			if constexpr (std::is_same_v<X, XLOPER>) {
				return std::string(val.str + 1, val.str[0]);
			}
			else {
				return utf8::wcstostring(val.str + 1, val.str[0]);
			}
		}

		// replace non alphanumeric or period '.' with underscore
		XOPER safe(xchar u = '_') const
		{
			XOPER s(*this);

			if (s.is_str()) {
				for (int i = 1; i <= s.val.str[0]; ++i) {
					if (s.val.str[i] != '.' && !traits<X>::alnum(s.val.str[i])) {
						s.val.str[i] = u;
					}
				}
			}

			return s;
		}

		// xltypeBool
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
		bool operator==(bool xbool) const
		{
			return xltype == xltypeBool && val.xbool == static_cast<typename traits<X>::xbool>(xbool);
		}

		[[nodiscard]] bool is_bool() const
		{
			return type() == xltypeBool;
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

		/*
		// xltypeRef - reference to multiple refs
		XOPER(const std::initializer_list<xref>& ref)
		{
			xltype = xltypeRef;
			ref_alloc(static_cast<WORD>(ref.size()));
			std::copy(ref.begin(), ref.end(), val.mref.lpmref->reftbl);
		}
		*/
		[[nodiscard]] bool is_ref() const
		{
			return type() == xltypeRef;
		}
		// size() returns number of refs
		const xref* as_ref(unsigned i) const
		{
			ensure(is_ref());

			return val.mref.lpmref->reftbl;
		}
		xref* as_ref()
		{
			ensure(is_ref());
	
			return val.mref.lpmref->reftbl;
		}

		// xltypeErr - predefined as ErrXXX
#define XLL_ERR_ENUM(a, b, c) a = xlerr##a,
		enum class Err {
			XLL_ERR(XLL_ERR_ENUM) // Err::NA, etc
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
		// xltypeErr - predefined as ErrXXX
		[[nodiscard]] bool is_err() const
		{
			return xltype == xltypeErr;
		}

		// xltypeRef
		// xltypeFlow - not used

		// xltypeMulti
		[[nodiscard]] bool is_multi() const
		{
			return type() == xltypeMulti;
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
		// XOPER(a, ...)
		template<class... Os>
		static XOPER make(const Os&... os)
		{
			return XOPER({ XOPER(os)... });
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
		// key-value lookup using MATCH
		const XOPER& operator[](const XOPER& key) const
		{
			if (size() % 2 != 0) {
				return XErrNA<XLOPERX>;
			}

			XOPER i = Excel(xlfMatch, key, *this, XOPER(0));

			return i ? operator[](i.as_int() + 1) : XErrNA<XLOPERX>;
		}
		const XOPER& operator[](const char* key) const
		{
			return operator[](XOPER(key));
		}

		XOPER<X>& drop(int n)
		{
			if (static_cast<unsigned>(abs(n)) >= size()) {
				oper_free();

				return *this;
			}
			if (n == 0) {
				return *this;
			}

			X x = *this; // copy bits
			x = xll::drop(x, n);

			if (n > 0) {
				for (unsigned i = 0; i < size() - xll::size(x); ++i) {
					operator[](i).oper_free();
				}
				MoveMemory(begin(), xll::begin(x), xll::size(x) * sizeof(X));
			}
			else {
				for (unsigned i = xll::size(x); i < size(); ++i) {
					operator[](i).oper_free();
				}
			}

			val.array.rows = x.val.array.rows;
			val.array.columns = x.val.array.columns;

			return *this;
		}
		XOPER<X>& take(int n)
		{
			if (static_cast<unsigned>(abs(n)) >= size()) {
				return *this;
			}
			if (n == 0) {
				oper_free();

				return *this;
			}

			X x = *this; // copy bits
			x = xll::take(x, n);

			if (n > 0) {
				for (unsigned i = xll::size(x); i < size(); ++i) {
					operator[](i).oper_free();
				}
			}
			else {
				for (unsigned i = 0; i < size() - xll::size(x); ++i) {
					operator[](i).oper_free();
				}
				MoveMemory(begin(), xll::begin(x), xll::size(x) * sizeof(X));
			}

			val.array.rows = x.val.array.rows;
			val.array.columns = x.val.array.columns;

			return *this;
		}

		enum class Side {
			Bottom, Right, Top, Left
		};
		const XOPER& push_back(const X& x, Side side = Side::Bottom)
		{
			if (x.xltype == xltypeNil) {
				return *this;
			}
			if (xltype == xltypeNil) {
				operator=(x);

				return *this;
			}

			const XOPER& o = (overlap(x) ? XOPER(x) : x);
			if (type() == xltypeMulti) {
				if (type() != xltype) {
					operator=(*this); // free old and make new copy
				}
				unsigned r = rows();
				unsigned c = columns();
				unsigned n = r * c;
				if (side == Side::Bottom) {
					ensure(c == o.columns());
					multi_realloc(r + o.rows(), c);
					std::copy(o.begin(), o.end(), begin() + n);
				}
				else if (side == Side::Right) {
					ensure(r == o.rows());
					multi_realloc(r, c + o.columns());
					// copy to back of old range
					std::copy(o.begin(), o.end(), begin() + n);
					// rotate new rows into place
					auto b = begin() + c;
					auto m = b + c;
					auto e = m + o.columns();
					while (--r) {
						std::rotate(b, m, e);
						b += columns();
						m += columns();
						e += columns();
					}
				}
				// else if (side == Side::Top) { ... }
				// else if (side == Side::Left) { ... }
				else {
					ensure(!"XOPER::push_back: unimplimented Side");
				}
			}
			else {
				resize(1, 1);

				return push_back(o, side);
			}

			return *this;
		}
		const XOPER& push_bottom(const X& x)
		{
			return push_back(x, Side::Bottom);
		}
		const XOPER& push_right(const X& x)
		{
			return push_back(x, Side::Right);
		}

		// xltypeMissing - predefined as Missing
		bool is_missing() const
		{
			return type() == xltypeMissing;
		}
		// xltypeNil - predefined as Nil
		bool is_nil() const
		{
			return type() == xltypeNil;
		}

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
			return type() == xltypeSRef;
		}
		xref& as_sref(unsigned i = 0)
		{
			if (type() == xltypeRef) {
				ensure(i < val.mref.lpmref->count);

				return val.mref.lpmref->reftbl[i];
			}

			ensure(i == 0 && type() == xltypeSRef);

			return val.sref.ref;
		}
		const xref& as_sref(unsigned i = 0) const
		{
			if (type() == xltypeRef) {
				ensure(i < val.mref.lpmref->count);

				return val.mref.lpmref->reftbl[i];
			}

			ensure(i == 0 && type() == xltypeSRef);

			return val.sref.ref;
		}
		// xltypeInt. Excel usually converts this to num.
		bool is_int() const
		{
			return type() == xltypeInt;
		}
		xint as_int() const
		{
			xint w = 0;

			switch (type()) {
			case xltypeInt:
				w = val.w;
				break;
			case xltypeBool:
				w = val.xbool;
				break;
			case xltypeNum:
				w = static_cast<xint>(val.num);
				break;
			default:
				ensure(!"OPER::as_int: non-numeric type");
			}
			
			return w;
		}

		// xltypeBigData
		bool is_binary() const
		{
			return xltype == xltypeBigData;
		}
		// Used to allocate persistent storage for an xltypeBigData XLOPER/ XLOPER12. 
		// Data with a defined binary name is saved with the workbook, and can be accessed by name at any time. 
		// xlDefineBinaryName, xlGetBinaryName

	private:
		// true if memory overlaps with x
		bool overlap(const X& x) const
		{
			return (begin() <= xll::begin(x) && xll::begin(x) < end()) 
				|| (begin() < xll::end(x)    && xll::end(x)  <= end());
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
			/*
			else if (xltype == xltypeRef) {
				ref_free();
			}
			*/
			
			xltype = xltypeNil;
		}

		// allocate unless counted
		void str_alloc(xcstr str, int n)
		{
			ensure(str);
			xltype = xltypeStr;

			if (n == -1) {
				// str is counted string allocated by malloc
				val.str = const_cast<xchar*>(str);

				return;
			}

			ensure (n <= traits<X>::charmax);

			xchar* tmp = (xchar*)malloc(((size_t)n + 2) * sizeof(xchar));
			ensure(tmp);
			// always null terminate
			if (n != 0) {
				traits<X>::cpy(tmp + 1, n + 1, str, n);
				tmp[n + 1] = 0;
			}
			else {
				tmp[0] = tmp[1] = 0;
			}
			// first character is count
			tmp[0] = static_cast<xchar>(n);
			val.str = tmp;
		}

		// append str up to null and always allocate n more chars
		// if n == -1 then str is counted and will be freed
		void str_append(xcstr str, int n)
		{
			if (xltype == xltypeNil) {
				str_alloc(str, n);

				return;
			}

			ensure(type() == xltypeStr);

			bool counted = (n == -1);
			if (counted) {
				n = str[0];
				if (n == 0 || str[1] == 0) {
					free(const_cast<xchar*>(str));

					return; // noop
				}
			}
			else if (n == 0 || str[0] == 0) {
				return; // noop
			}

			if (xltype & xlbitXLFree) {
				XOPER tmp(val.str + 1, val.str[0]);
				swap(tmp);
				//swap(XOPER(val.str + 1, val.str[0]));
			}

			ensure(val.str[0] + n < traits<X>::charmax);

			xchar* tmp = (xchar*)realloc(val.str, ((size_t)val.str[0] + n + 2) * sizeof(xchar));
			ensure(tmp);
			traits<X>::cpy(tmp + 1 + tmp[0], n + 1, str + counted, n);
			tmp[0] = static_cast<xchar>(val.str[0] + n);
			tmp[tmp[0] + 1] = 0;
			val.str = tmp;
			//memcpy_s(val.str + 1 + val.str[0], n * sizeof(xchar), str + counted, n * sizeof(xchar));
			//if (n == 1 and str[0] == 0) {
			//	--n; // don't count null terminator
			//}
			//val.str[0] = static_cast<xchar>(val.str[0] + n);

			if (counted) {
				free(const_cast<xchar*>(str));
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
				operator=(XOPER{});

				return;
			}

			xltype = xltypeMulti;
			val.array.rows = static_cast<xrw>(r);
			val.array.columns = static_cast<xcol>(c);
			val.array.lparray = static_cast<X*>(malloc(size() * sizeof(X)));
			if (!val.array.lparray) {
				throw std::bad_alloc{};
			}
			for (unsigned i = 0; i < size(); ++i) {
				new (val.array.lparray + i) XOPER{};
			}
		}
		void multi_realloc(unsigned r, unsigned c)
		{
			ensure(r <= (unsigned)(std::numeric_limits<xrw>::max)());
			ensure(c <= (unsigned)(std::numeric_limits<xcol>::max)());

			if (!(xltype & xltypeMulti)) {
				XOPER tmp = std::move(*this);
				multi_alloc(r, c);
				operator[](0) = std::move(tmp);

				return;
			}

			if (r * c == 0) {
				multi_free();
				ensure(xltype == XOPER{}.xltype);

				return;
			}
	
			// current size
			auto n = size();
			if (n > r*c) {
				std::for_each(end(), begin() + n, [](auto& o) { o.oper_free(); });
				val.array.rows = static_cast<xrw>(r);
				val.array.columns = static_cast<xcol>(c);
			}
			else if (n < r*c) {
				XOPER tmp(r, c);
				for (unsigned i = 0; i < n; ++i) {
					tmp[i] = std::move(operator[](i));
				}
				for (unsigned i = n; i < size(); ++i) {
					new (tmp.val.array.lparray + i) XOPER{};
				}
				swap(tmp);
			}
			else {
				val.array.rows = static_cast<xrw>(r);
				val.array.columns = static_cast<xcol>(c);
			}
		}
		void multi_free()
		{
			ensure(xltype == xltypeMulti);
			std::for_each(begin(), end(), [](auto& o) { o.oper_free(); });
			free(val.array.lparray);

			xltype = xltypeNil;
		}
#if 0
		void ref_alloc(WORD count)
		{
			using xref = typename traits<X>::xref;
			using xmref = typename traits<X>::xmref;

			xltype = xltypeRef;
			val.mref.lpmref = (xmref*)malloc(sizeof(xmref) + count * sizeof(xref));
			if (!val.mref.lpmref) {
				throw std::bad_alloc{};
			}
			val.mref.lpmref->count = count;
		}
		void ref_realloc(WORD count)
		{
			using xref = typename traits<X>::xref;
			using xmref = typename traits<X>::xmref;

			ensure(xltype == xltypeRef);
			void* tmp = realloc(val.mref.lpmref, sizeof(xmref) + count * sizeof(xref));
			if (!tmp) {
				throw std::bad_alloc{};
			}
			val.mref.lpmref = (xmref*)tmp;
			val.mref.lpmref->count = count;
		}
		void ref_free()
		{
			ensure(xltype == xltypeRef);
			//free(val.mref.lpmref);

			xltype = xltypeNil;
		}
#endif // 0
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