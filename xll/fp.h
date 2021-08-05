// fp.h - Two-dimensional array of doubles
#pragma once
#include "traits.h"

namespace xll {

	inline unsigned short size(const _FP& a)
	{
		return a.rows * a.columns;
	}

	inline double& index(_FP& a, unsigned short i)
	{
		return a.array[xmod<unsigned short>(i, size(a))];
	}
	inline double& index(_FP& a, unsigned short i, unsigned short j)
	{
		return index(a, a.columns * xmod<unsigned short>(i, a.rows) + xmod<unsigned short>(j, a.columns));
	}

	// Make FP STL friendly.
	inline double* begin(_FP& a)
	{
		return a.array;
	}
	inline const double* begin(const _FP& a)
	{
		return a.array;
	}
	inline double* end(_FP& a)
	{
		return a.array + size(a);
	}
	inline const double* end(const _FP& a)
	{
		return a.array + size(a);
	}

	inline unsigned size(const _FP12& a)
	{
		return a.rows * a.columns;
	}

	inline double& index(_FP12& a, unsigned i)
	{
		return a.array[xmod<unsigned>(i, size(a))];
	}
	inline double& index(_FP12& a, unsigned i, unsigned j)
	{
		return index(a, a.columns * xmod<unsigned>(i, a.rows) + xmod<unsigned>(j, a.columns));
	}

	inline double* begin(_FP12& a)
	{
		return a.array;
	}
	inline const double* begin(const _FP12& a)
	{
		return a.array;
	}
	inline double* end(_FP12& a)
	{
		return a.array + size(a);
	}
	inline const double* end(const _FP12& a)
	{
		return a.array + size(a);
	}

	inline _FP12* drop(_FP12* pa, int n)
	{
		if (n == 0) {
			return pa;
		}

		int na = static_cast<int>(size(*pa));
		n = std::clamp(n, -na, na);

		int off;
		if (pa->rows == 1) {
			off = abs(n);
			pa->columns -= abs(n);
		}
		else {
			off = abs(n) * pa->columns;
			pa->rows -= abs(n);
		}

		if (n > 0) {
			MoveMemory(begin(*pa), begin(*pa) + off, size(*pa) * sizeof(double));
		}

		return pa;
	}

	inline _FP12* take(_FP12* pa, int n)
	{
		if (n == 0) {
			pa->rows = 0;
			pa->columns = 0;

			return pa;
		}

		int na = static_cast<int>(size(*pa));
		n = std::clamp(n, -na, na);

		int off;
		auto e = end(*pa);
		if (pa->rows == 1) {
			off = abs(n);
			pa->columns = abs(n);
		}
		else {
			off = abs(n) * pa->columns;
			pa->rows = abs(n);
		}

		if (n < 0) {
			MoveMemory(begin(*pa), e - off, size(*pa) * sizeof(double));
		}

		return pa;
	}

	template<class Op>
	inline _FP12* scan(const Op& op, _FP12* pa)
	{
		if (pa->rows == 1) {
			for (unsigned i = 1; i < pa->columns; ++i) {
				pa->array[i] = op(pa->array[i - 1], pa->array[i]);
			}
		}
		else {
			// scan each column
			for (unsigned j = 0; j < pa->columns; ++j) {
				for (unsigned i = 1; i < pa->rows; ++i) {
					index(*pa, i, j) = op(index(*pa, i - 1, j), index(*pa, i, j));
				}
			}
		}

		return *pa;
	}

	template<class Op>
	inline _FP12* diff(_FP12* pa/*, const Op& op = std::minus<double>{}*/)
	{
		if (xll::size(*pa) <= 1) {
			return pa;
		}

		Op op = Op{};
		if (pa->rows == 1) {
			for (unsigned i = pa->columns - 1; i; --i) {
				pa->array[i] = op(pa->array[i], pa->array[i - 1]);
			}
		}
		else {
			// diff each column
			for (int j = 0; j < pa->columns; ++j) {
				for (int i = pa->rows - 1; i; --i) {
					index(*pa, i, j) = op(index(*pa, i, j), index(*pa, i - 1, j));
				}
			}
		}

		return pa;
	}

	// fixed size FP types
	template<unsigned R, unsigned C = 1>
	struct FP_ {
		unsigned short int rows;
		unsigned short int columns;
		double array[R*C];
		FP_()
			: rows(R), columns(C)
		{
		}
		auto size() const
		{
			return rows * columns;
		}
		FP_& resize(unsigned short int r, unsigned short int c)
		{
			ensure(r * c <= R * C);

			rows = r;
			columns = c;

			return *this;
		}
		double& operator[](unsigned short int i)
		{
			return index(*(_FP*)this, i);
		}
		double operator[](unsigned short int i) const
		{
			return index(*(_FP*)this, i);
		}
	};

	template<unsigned R, unsigned C = 1>
	struct FP12_ {
		INT32 rows;
		INT32 columns;
		double array[R*C];
		FP12_()
			: rows(R), columns(C)
		{ }
		unsigned size() const
		{
			return rows * columns;
		}
		FP12_& resize(int r, int c)
		{
			ensure(r * c <= R * C);

			rows = r;
			columns = c;

			return *this;
		}
		double& operator[](unsigned i)
		{
			return index(*(_FP12*)this, i);
		}
		double operator[](unsigned i) const
		{
			return index(*(_FP12*)this, i);
		}
	};

	template<unsigned R, unsigned C = 1>
#if XLL_VERSION == 12
	using FPX_ = FP12_<R,C>;
#else
	using FPX_ = FP_<R,C>;
#endif

	/// <summary>
	/// C++ wrapper for FP and FP12 data types
	/// </summary>
	template<class X>
		requires (std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>)
	class XFP {
		typedef typename traits<X>::xint xint;
		typedef typename traits<X>::xfp xfp;
		xfp* fp;
	public:
		XFP(xint r = 1, xint c = 1)
		{
			fp_alloc(r, c);
		}
		XFP(const xfp& a)
			: XFP(a.rows, a.columns)
		{
			MoveMemory(begin(), a.array, size() * sizeof(double));
		}
		XFP(const XFP& a)
			: XFP(*a.fp)
		{ }
		XFP(XFP&& a) noexcept
			: fp(std::exchange(a.fp, nullptr))
		{ }
		XFP& operator=(const XFP& a)
		{
			if (this != &a) {
				fp_realloc(a.rows(), a.columns());
				MoveMemory(begin(), a.begin(), size() * sizeof(double));
			}

			return *this;
		}
		XFP& operator=(const xfp& a)
		{
			fp_realloc(a.rows, a.columns);
			MoveMemory(begin(), xll::begin(a), size() * sizeof(double));

			return *this;
		}
		XFP& operator=(XFP&& a) noexcept
		{
			fp = std::exchange(a.fp, nullptr);

			return *this;
		}
		~XFP()
		{
			fp_free();
		}

		// Convert to native Excel FP type pointer.
		xfp* get()
		{
			return fp;
		}
		const xfp* get() const
		{
			return fp;
		}
		/*
		xfp* operator&()
		{
			return get();
		}
		*/

		bool operator==(const XFP& a) const
		{
			if (rows() != a.rows()) {
				return false;
			}
			if (columns() != a.columns()) {
				return false;
			}
			
			return std::equal(begin(), end(), a.begin());
		}

		XFP& resize(xint r, xint c)
		{
			fp_realloc(r, c);

			return *this;
		}

		xint rows() const
		{
			return fp ? fp->rows : 0;
		}
		xint columns() const
		{
			return fp ? fp->columns : 0;
		}
		xint size() const
		{
			return rows() * columns();
		}
		bool is_empty() const
		{
			return fp == nullptr;
		}
		double* array()
		{
			return fp->array;
		}
		const double* array() const
		{
			return fp->array;
		}
		double& operator[](xint i)
		{
			ensure(fp != nullptr);

			return fp->array[xmod(i, size())];
		}
		const double& operator[](xint i) const
		{
			ensure(fp != nullptr);

			return fp->array[xmod(i, size())];
		}
		double& operator()(xint i, xint j)
		{
			return operator[](xmod(i, rows())*columns() + xmod(j, columns()));
		}
		const double& operator()(xint i, xint j) const
		{
			return operator[](xmod(i, rows())* columns() + xmod(j, columns()));
		}
		double* begin()
		{
			return fp ? fp->array : nullptr;
		}
		double* end()
		{
			return fp ? fp->array + size() : nullptr;
		}
		const double* begin() const
		{
			return fp ? fp->array : nullptr;
		}
		const double* end() const
		{
			return fp ? fp->array + size() : nullptr;
		}
		XFP& drop(int n)
		{
			fp = xll::drop(fp, n);

			return *this;
		}
		XFP& take(int n)
		{
			fp = xll::take(fp, n);

			return *this;
		}
	private:
		void fp_alloc(xint r, xint c)
		{
			ensure(r >= 0 and c >= 0);
			xint n = r * c;
			if (n == 0) {
				fp = nullptr;
			}
			else {
				fp = (xfp*)malloc(sizeof(xfp) + n * sizeof(double));
				if (!fp) {
					throw std::bad_alloc{};
				}
				fp->rows = r;
				fp->columns = c;
			}
		}
		void fp_realloc(xint r, xint c)
		{
			ensure(r >= 0 and c >= 0);
			xint n = r * c;
			if (n == 0) {
				fp_free();
			}
			else {
				if (n != size()) {
					void* tmp = realloc(fp, sizeof(xfp) + n * sizeof(double));
					if (!tmp) {
						throw std::bad_alloc{};
					}
					fp = (xfp*)tmp;
				}
				fp->rows = r;
				fp->columns = c;
			}
		}
		void fp_free()
		{
			free(fp);
			fp = nullptr;
		}
	};

	using FP4 = XFP<XLOPER>;
	using FP12 = XFP<XLOPER12>;
	using FPX = XFP<XLOPERX>;
}