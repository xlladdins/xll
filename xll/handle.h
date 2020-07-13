// handle.h - handles to C++ objects
#pragma once
#include <set>

typedef double HANDLEX;

namespace xll {

	inline HANDLEX p2h(void* p)
	{
		HANDLEX h = 0;
	
		union {
			void* p;
			uint64_t ui;
		} u = { .p = p };
		h = static_cast<HANDLEX>(u.ui);

		return h;
	}
	inline void* h2p(HANDLEX /*h*/)
	{
		return nullptr;
	}

	template<class T>
	class handle {
		static std::set<T*> ptr_;
		T* pt;
	public:
		handle(T* pt)
			: pt(pt)
		{
			// if coerce caller in map delete
		}
		handle(HANDLEX h)
		{
			// convert to pointer
			// look up in set
		}
		T* ptr()
		{
			return pt;
		}
		T* operator->()
		{
			return pt;
		}
		HANDLEX get()
		{
			return p2h(pt);
		}

	};

}
