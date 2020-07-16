// handle.h - handles to C++ objects
#pragma once
#include <algorithm>
#include <set>
#include "excel.h"

typedef double HANDLEX;
inline const auto XLL_HANDLEX = XLL_DOUBLEX;

namespace xll {

	template<class T>
	inline HANDLEX to_handle(T* p)
	{
		// first 16 bits are 0
		// double perfectly represents ints < 2^53
		if (((uint64_t)p) >> (64 - 16)) {
			MessageBox(GetForegroundWindow(), _T("handle > 2^48"), 0, MB_OK);
		}
		return (HANDLEX)((uint64_t)p);
	}
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
		/// garbage collect handles on exit
		struct pointers : public std::set<T*> {
			using std::set<T*>::begin;
			using std::set<T*>::end;
			~pointers()
			{
				std::for_each(begin(), end(), [](auto i) { delete i; });
			}
		};
		inline static pointers ps;
		T* p;
	public:
		/// <summary>
		/// Add a handle to the collection.
		/// If the calling cell has a known handle then
		/// delete the object and remove handle from collecton.
		/// </summary>
		handle(T* p) noexcept
			: p(p)
		{
			// value in calling cell
			OPERX x = Excel<XLOPERX>(xlCoerce, Excel<XLOPERX>(xlfCaller));
			if (x.xltype == xltypeNum && x.val.num != 0) {
				// looks like a handle
				T* px = to_pointer<T>(x.val.num);
				auto pi = ps.find(px);
				// already seen so garbage collect
				if (pi != ps.end()) {
					delete* pi;
					ps.erase(pi);
				}
			}

			ps.insert(p);
		}
		handle(HANDLEX h) noexcept
			: p(to_pointer<T>(h))
		{
			// some measure of type safety
			if (!ps.contains(p)) {
				p = nullptr;
			}
		}
		operator bool() const
		{
			return p != nullptr;
		}
		T* operator->()
		{
			return p;
		}
		T* ptr()
		{
			return p;
		}
		// return value for Excel
		HANDLEX get()
		{
			return to_handle(p);
		}

	};

}
