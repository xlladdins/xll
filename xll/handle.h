// handle.h - handles to C++ objects
#pragma once
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

	template<class T>
	class handle {
		inline static std::set<T*> ps;
		T* p;
	public:
		handle(T* p = nullptr) noexcept
			: p(p)
		{
			OPERX x = Excel<XLOPERX>(xlCoerce, Excel<XLOPERX>(xlfCaller));
			if (x.xltype == xltypeNum && x.val.num != 0) {
				T* px = (T*)h2p(x.val.num);
				auto pi = ps.find(px);
				if (pi != ps.end()) {
					ps.erase(pi);
					delete *pi;
				}
			}

			if (p) {
				ps.insert(p);
			}
		}
		handle(HANDLEX h) noexcept
			: p((T*)h2p(h))
		{
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
		HANDLEX get()
		{
			return p2h(p);
		}

	};

}
