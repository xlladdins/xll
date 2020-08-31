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
	/// Use <c>handle<T> h(new T(...))</c> to create a handle.
	/// Use <c>handle<T> h_(h)</c> to lookup h returned by <c>get()</c>.
	/// Unknown handles return null pointers.
	/// </summary>
	template<class T>
	class handle {
		// all active pointers of type T*
		inline static std::set<std::unique_ptr<T>> ps;
		// underlying pointer
		T* p;

		// Convert handle in caller to pointer.
		static T* caller()
		{
			OPER x = Excel(xlCoerce, Excel(xlfCaller));

			return x.xltype == xltypeNum ? to_pointer<T>(x.val.num) : nullptr;
		}

		/// If calling cell has a known handle then delete corresponding object
		void gc(void)
		{
			if (T* p_ = caller()) {
				std::unique_ptr<T> px(p_);
				auto pi = ps.find(px);
				px.release(); // do not call delete on p_!
				// garbage collect
				if (pi != ps.end()) {
					ps.erase(pi); // calls delete
				}
			}
		}
	public:
		/// <summary>
		/// Add a handle to the collection.
		/// </summary>
		handle(T* p) noexcept
			: p(p)
		{
			ensure(nullptr != p);

			gc();

			ps.emplace(p);
		}
		handle(HANDLEX h, bool check = true) noexcept
			: p(to_pointer<T>(h))
		{
			ensure(0 != h);

			if (check) {
				std::unique_ptr<T> p_(p);
				if (ps.find(p_) == ps.end()) {
					// unknown handle
					p = nullptr;
				}
				p_.release(); // do not call delete on p!
			}
		}
		~handle()
		{
			// do nothing
		}
		operator bool() const
		{
			return p != nullptr;
		}
		// return value for Excel
		HANDLEX get() const
		{
			return to_handle(p);
		}
		// underlying pointer
		T* ptr()
		{
			return p;
		}
		// act like a pointer
		T* operator->()
		{
			return p;
		}
		const T* operator->() const
		{
			return p;
		}
	};

	// endode/decode???

}
