// fp.h - Two-dimensional array of doubles
#pragma once
#include "traits.h"

namespace xll {

	/// <summary>
	/// C++ wrapper for FP and FP12 data types
	/// </summary>
	template<class X>
	class XFP {
		void* fp;
	public:
		typedef typename traits<X>::xint xint;
		typedef typename traits<X>::xfp xfp;
		XFP(xint r = 0, xint c = 0)
		{
			fp_alloc(r, c);
		}
		XFP(const XFP& fp)
		{
			fp_alloc(fp.rows, fp.columns);
			std::copy(fp.begin(), fp.end(), begin(), end());
		}
		~XFP()
		{
			free(fp);
		}

		operator xfp& ()
		{
			return reinterpret_cast<xfp&>(*(xfp*)fp);
		}
		operator const xfp& () const
		{
			return reinterpret_cast<const xfp&>(*(xfp*)fp);
		}

		xint rows() const
		{
			return operator const xfp&().rows;
		}
		xint columns() const
		{
			return operator const xfp&().columns;
		}
		xint size() const
		{
			return rows() * columns();
		}
		double* array()
		{
			return operator xfp&().array;
		}
		const double* array() const
		{
			return operator const xfp&().array;
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
			fp = malloc(sizeof(xfp) + r * c * sizeof(double));
			if (fp) {
				auto& pfx = operator xfp&();
				pfx.rows = r;
				pfx.columns = c;
				pfx.array[0] = 0;
			}
		}
		void fp_realloc(xint r, xint c)
		{
			fp = realloc(fp, sizeof(xfp) + r * c * sizeof(double));
			if (fp) {
				auto& pfx = operator xfp&();
				pfx.rows = r;
				pfx.columns = c;
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