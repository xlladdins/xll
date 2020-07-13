// handle.h - handles to C++ objects
#pragma once
#include <set>

namespace xll {

	using HANDLEX = double;

	inline HANDLEX p2h(void* p)
	{
		return 0;
	}
	inline void* h2p(HANDLEX h)
	{
		return nullptr;
	}

	template<class T>
	class handle {
		static set<T> ptr_;
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
