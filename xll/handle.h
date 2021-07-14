// handle.h - handles to C++ objects
#pragma once
#include <limits>
#include <map>
#include <memory>
#include <set>
#include <typeinfo>
#include <utility>
#include "excel.h"

// handle data type
using HANDLEX = double;
inline constexpr HANDLEX INVALID_HANDLEX = std::numeric_limits<HANDLEX>::quiet_NaN();
inline constexpr HANDLEX XLL_NAN = std::numeric_limits<HANDLEX>::quiet_NaN();

// handle argument types for add-ins
inline LPCSTR XLL_HANDLEX = XLL_DOUBLE;
inline LPCSTR XLL_HANDLE = XLL_DOUBLE;

namespace xll {

	/// <summary>
	/// Convert a pointer to a handle.
	/// </summary>
	/// <typeparam name="T">Any type</typeparam>
	/// <param name="p">A pointer to T</param>
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

	// keep track of handles returned to Excel
	inline std::set<void*> safe_pointers;
	
	template<class T>
	inline HANDLEX safe_handle(T* p)
	{
		safe_pointers.insert(p);

		return to_handle<T>(p);
	}

	template<class T>
	inline T* safe_pointer(HANDLEX h)
	{
		T* p = to_pointer<T>(h);

		return safe_pointers.contains(p) ? p : nullptr;
	}

	// typeid<T>.name() given pointer
	inline std::map<void*, const char*> handle_typename;

	// compare underlying raw pointers
	template<class T>
	struct unique_ptrcmp {
		using is_transparent = void;

		template<class T>
		bool operator()(const std::unique_ptr<T>& a, const T* b) const
		{
			return a.get() < b;
		}

		template<class T>
		bool operator()(const T* a, const std::unique_ptr<T>& b) const
		{
			return a < b.get();
		}

		template<class T>
		bool operator()(const std::unique_ptr<T>& a, const std::unique_ptr<T>& b) const
		{
			return a.get() < b.get();
		}
	};

	/// <summary>
	/// Collection of handles parameterized by type.
	/// They behave very much like <c>std::unique_ptr</c>
	/// but they are owned by the cell they are created in.
	/// When a new handle is created in a cell the destructor
	/// for the old handle is called.
	/// 
	/// Use <c>handle<T> h(new T(...))</c> to create a handle
	/// and <c>get()</c> to return a <c>HANDLEX</c> to Excel.
	/// Functions that create handles must be uncalced.
	/// 
	/// Use <c>handle<T> h_(h)</c> to lookup <c>h</c> returned by <c>get()</c>.
	/// Functions that use handles do not need to be uncalced.
	/// Unknown handles return null pointers.
	/// This can be circumvented by using <c>handle<T> h_(h, false)</c>
	/// to prevent the lookup.
	/// </summary>
	template<class T>
	class handle {
		// all active pointers of type T*
		inline static std::set<std::unique_ptr<T>, unique_ptrcmp<T>> ps;
		// caller of handle
		inline static std::map<T*, OPER> caller;

		static void erase(T* p) noexcept
		{
			if (p != nullptr) {
				auto p_ = ps.find(p);
				if (p_ != ps.end()) {
					ps.erase(p_);
				}
				auto ph = handle_typename.find(p);
				if (ph != handle_typename.end()) {
					handle_typename.erase(ph);
				}
			}
		}

		// Convert handle in caller to pointer.
		template<class T>
		static T* coerce(const OPER& caller)
		{
			OPER o = Excel(xlCoerce, caller);

			return o.xltype == xltypeNum ? to_pointer<T>(o.val.num) : nullptr;
		}

		// underlying pointer
		T* p;
	public:
		/// <summary>
		/// Add a handle to the collection.
		/// </summary>
		handle(T* p) noexcept
			: p{ p }
		{
			// caller has previously created handle
			erase(coerce<T>(caller[p] = Excel(xlfCaller)));

			ps.emplace(std::unique_ptr<T>(p));
			handle_typename[p] = typeid(*p).name();
		}
		/// <summary>
		/// Lookup an existing handle.
		/// </summary>
		handle(HANDLEX h, bool check = true) noexcept
			: p(to_pointer<T>(h))
		{
			if (check and p) {
				if (ps.find(p) == ps.end()) {
					// unknown handle
					p = nullptr;
				}
			}
			// handle was created by a function argument
			if (check and caller[p] == Excel(xlfCaller)) {
				caller[p] = ErrNA; // mark for erasure
			}
		}
		handle(const handle&) = delete;
		handle& operator=(const handle&) = delete;
		handle(handle&& h) noexcept
			: p(h.p)
		{ }
		handle& operator=(handle&& h) noexcept
		{
			if (p != h.p) {
				swap(h);
			}

			return *this;
		}
		~handle()
		{
			// temporary handle
			if (caller[p] == ErrNA) {
				erase(p);
			}
		}

		void swap(handle& h)
		{
			using std::swap;

			// ps unchanged
			swap(p, h.p);
		}

		explicit operator bool() const
		{
			return p != nullptr;
		}

		// return value for Excel
		HANDLEX get() const
		{
			return to_handle(p);
		}
		// underlying pointer
		T* ptr() const
		{
			return p;
		}

		// act like a unique pointer
		typename std::add_lvalue_reference<T>::type operator*() const
		{
			return *p;
		}
		T* operator->()
		{
			return p;
		}
		const T* operator->() const
		{
			return p;
		}

		// downcast to U
		template<class U> 
			requires std::is_base_of_v<T,U>
		U* as()
		{
			return dynamic_cast<U*>(p);
		}

		// encode/decode handles to strings
		template<class X>
		class codec {
			using xchar = traits<X>::xchar;
			static uint8_t c2h(xchar c) // assumes ASCII
			{
				return static_cast<uint8_t>(c <= '9' ? c - '0' : 10 + c - 'A');
			}
			static xchar h2c(uint8_t h) // assumes ASCII
			{
				return static_cast<xchar>(h <= 9 ? '0' + h : 'A' + h - 10);
			}
			XOPER<X> H;  // "prefix0123456789ABCDEFsuffix
			unsigned off; // size of prefix
		public:
			// e.g., codec c(OPER("\\MyClass["), OPER("]"));
			template<class X_>
				requires std::is_base_of_v<X, X_>
			codec(const X_& prefix, const X_& suffix)
				: H(prefix), off(prefix.val.str[0])
			{
				H.append("0123456789ABCDEF");
				H.append(suffix);
			}

			// does not allocate memory
			const XOPER<X>& encode(HANDLEX h)
			{
				union {
					T* h;
					uint8_t c[8];
				} hc;
				hc.h = to_pointer<T>(h);

				// h -> "prefix01..Fsuffix"
				xchar* pc = H.val.str + 1 + off;
				for (unsigned i = 0; i < 8; ++i) {
					pc[2 * i] = h2c(hc.c[7 - i] >> 4);
					pc[2 * i + 1] = h2c(hc.c[7 - i] & 0x0F);
				}

				return H;
			}

			// does not allocate memory
			HANDLEX decode(const XOPER<X>& _H)
			{
				ensure(_H.is_str());

				// Extra chars appended to suffix ok.
				// Could use this to add, e.g., a timestamp.
				ensure(_H.val.str[0] >= H.val.str[0]);
				
				// No prefix or suffix check.It will fail when used.
				
				union {
					T* h;
					uint8_t c[8];
				} hc;
				// "prefix01..Fsuffix" -> h
				xchar* pc = _H.val.str + 1 + off;
				for (unsigned i = 0; i < 8; ++i) {
					hc.c[7 - i] = (c2h(pc[2 * i]) << 4) + c2h(pc[2 * i + 1]);
				}

				return to_handle<T>(hc.h);
			}
		};
	};

	template<class T>
	void swap(handle<T> &h, handle<T>& k)
	{
		h.swap(k);
	}

}
