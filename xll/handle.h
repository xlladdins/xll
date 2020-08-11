// handle.h - handles to C++ objects
#pragma once
#include <memory>
#include <set>
#include "excel.h"

// handle data type
using HANDLEX = double;

// handle argument types for add-ins
inline const auto XLL_HANDLE = XLL_DOUBLE;
inline const auto XLL_HANDLEX = XLL_DOUBLE;

namespace xll {

	/// <summary>
	/// Convert a pointer to a handle.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <param name="p">is a pointer</param>
	/// <returns>Returns a handle encoding the pointer bits.</returns>
	/// 
	template<class T>
	inline HANDLEX to_handle(T* p)
	{
		// first 16 bits of p are 0 and 2^(64 - 16) = 2^48
		// doubles exactly represent integers < 2^53
		return (HANDLEX)((uint64_t)p);
	}
	/// <summary>
	/// Convert a handle to a pointer.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	/// <param name="h">is a handle</param>
	/// <returns>Returns the decoded handle bits.</returns>
	/// 
	template<class T>
	inline T* to_pointer(HANDLEX h)
	{
		return (T*)((uint64_t)h);
	}

	/// <summary>
	/// Collection of handles parameterized by type.
	/// </summary>
	template<class T>
	class handle {
		inline static std::set<std::unique_ptr<T>> ps;
		T* p;
	public:
		/// <summary>
		/// Add a handle to the collection.
		/// If the calling cell has a known handle then
		/// delete the object and erase pointer.
		/// </summary>
		handle(T* p) noexcept
			: p(p)
		{
			// value in calling cell
			OPER x = XExcel<XLOPERX>(xlCoerce, XExcel<XLOPERX>(xlfCaller));
			if (x.xltype == xltypeNum && x.val.num != 0) {
				// looks like a handle
				std::unique_ptr<T> px(to_pointer<T>(x.val.num));
				auto pi = ps.find(px);
				if (pi != ps.end()) {
					ps.erase(pi);
				}
				px.release();
			}

			ps.insert(std::unique_ptr<T>(p));
		}
		handle(HANDLEX h) noexcept
			: p(to_pointer<T>(h))
		{
			// unknown handle
			std::unique_ptr<T> px(p);
			if (ps.find(px) == ps.end()) {
				p = nullptr;
			}
			px.release();
		}
		operator bool() const
		{
			return p != nullptr;
		}
		// return value for Excel
		HANDLEX get()
		{
			return to_handle(p);
		}
		T* ptr()
		{
			return p;
		}
		T* operator->()
		{
			return p;
		}
		const T* operator->() const
		{
			return p;
		}
	};

}
