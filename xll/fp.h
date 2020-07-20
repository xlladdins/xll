// fp.h - Two-dimensional array of doubles
#pragma once
#include "traits.h"

namespace xll {

	/// <summary>
	/// C++ wrapper for FP and FP12 data types
	/// </summary>
	template<class X>
	class XFP {
		typedef typename traits<X>::xint xint;
		typedef typename traits<X>::xfp xfp;
		void* fp;
	public:
		XFP(xint r = 0, xint c = 0)
		{
			fp_alloc(r, c);
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
			return reinterpret_cast<const xfp*>(fp)->rows;
		}
		xint columns() const
		{
			return reinterpret_cast<const xfp*>(fp)->columns;
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
			return reinterpret_cast<xfp*>(fp)->array;
		}
		const double* array() const
		{
			return reinterpret_cast<const xfp*>(fp)->array;
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
				xfp* pfx = reinterpret_cast<xfp*>(fp);
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
				xfp* pfx = reinterpret_cast<xfp*>(fp);
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