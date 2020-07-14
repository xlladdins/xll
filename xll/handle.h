// handle.h - handles to C++ objects
#pragma once
#include <algorithm>
#include <set>
#include "excel.h"

typedef double HANDLEX;
inline const auto XLL_HANDLEX = XLL_DOUBLEX;

namespace xll {

	inline HANDLEX p2h(void* p)
	{
		// first 16 bits are 0 so (uint64_t)p < 2^48
		// double perfectly represents ints < 2^53
		return (HANDLEX)((uint64_t)p);
	}
	inline void* h2p(HANDLEX h)
	{
		return (void*)((uint64_t)h);
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
		/// If the calling cell has what looks like a handle then
		/// delete the object and remove handle from collecton.
		/// </summary>
		handle(T* p) noexcept
			: p(p)
		{
			// value in calling cell
			OPERX x = Excel<XLOPERX>(xlCoerce, Excel<XLOPERX>(xlfCaller));
			if (x.xltype == xltypeNum && x.val.num != 0) {
				// looks like a handle
				T* px = (T*)h2p(x.val.num);
				auto pi = ps.find(px);
				// if so, garbage collect
				if (pi != ps.end()) {
					delete* pi;
					ps.erase(pi);
				}
			}

			ps.insert(p);
		}
		handle(HANDLEX h) noexcept
			: p((T*)h2p(h))
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
			return p2h(p);
		}

	};

}
