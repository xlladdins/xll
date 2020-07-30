// fp.h - Two-dimensional array of doubles
#pragma once
#include "traits.h"

inline auto size(const _FP& a)
{
	return a.rows * a.columns;
}

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

inline auto size(const _FP12& a)
{
	return a.rows * a.columns;
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

namespace xll {

	/// <summary>
	/// C++ wrapper for FP and FP12 data types
	/// </summary>
	template<class X>
		requires (std::is_same_v<XLOPER, X> || std::is_same_v<XLOPER12, X>)
	class XFP {
		typedef typename traits<X>::xint xint;
		typedef typename traits<X>::xfp xfp;
		void* fp;
	public:
		XFP(xint r = 0, xint c = 0)
		{
			fp_alloc(r, c);
		}
		XFP(const xfp& a)
			: XFP(a.rows, a.columns)
		{
			std::copy(::begin(a), ::end(a), begin());
		}
		XFP(const XFP& a)
		{
			fp_alloc(a.rows(), a.columns());
			std::copy(a.begin(), a.end(), begin());
		}
		XFP(XFP&& a) = default;
		XFP& operator=(const XFP& a)
		{
			if (this != &a) {
				fp_realloc(a.rows(), a.columns());
				std::copy(a.begin(), a.end(), begin());
			}

			return *this;
		}
		XFP& operator=(XFP&& a) = default;
		~XFP()
		{
			free(fp);
		}

		// Convert to pointer to native Excel FP type.
		xfp* get()
		{
			return reinterpret_cast<xfp*>(fp);
		}
		const xfp* get() const
		{
			return reinterpret_cast<const xfp*>(fp);
		}

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
			return get()->rows;
		}
		xint columns() const
		{
			return get()->columns;
		}
		xint size() const
		{
			return rows() * columns();
		}
		bool is_empty() const
		{
			return rows() == 0 && columns() == 0;
		}
		double* array()
		{
			return get()->array;
		}
		const double* array() const
		{
			return get()->array;
		}
		double& operator[](xint i)
		{
			return array()[i];
		}
		const double& operator[](xint i) const
		{
			return array()[i];
		}
		double& operator()(xint i, xint j)
		{
			return operator[](i*columns() + j);
		}
		const double& operator()(xint i, xint j) const
		{
			return operator[](i* columns() + j);
		}
		double* begin()
		{
			return array();
		}
		double* end()
		{
			return array() + size();
		}
		const double* begin() const
		{
			return array();
		}
		const double* end() const
		{
			return array() + size();
		}
	private:
		void fp_alloc(xint r, xint c)
		{
			size_t n = r * c;
			fp = malloc(sizeof(xfp) + n * sizeof(double));
			if (fp) {
				xfp* pfx = get();
				pfx->rows = r;
				pfx->columns = c;
			}
		}
		void fp_realloc(xint r, xint c)
		{
			xint n = r * c;
			if (n != size()) {
				fp = realloc(fp, sizeof(xfp) + n * sizeof(double));
			}
			if (fp) {
				xfp* pfx = get();
				pfx->rows = r;
				pfx->columns = c;
			}
		}
		void fp_free()
		{
			free(fp);
		}
	};

	using FP = XFP<XLOPER>;
	using FP12 = XFP<XLOPER12>;
	using FPX = XFP<XLOPERX>;
}