// codec.h - encoder/decoder for XLOPERs
// XLOPER is a variant type num|str|err|multi|...
#pragma once
#include <charconv>
#include <concepts>
#include <chrono>
#include <cstring>
#include "view.h"

namespace xll {

	// decode string of T to X
	template<class X>
	struct decoder {
		virtual ~decoder()
		{ }
		/*
		static X convert(view<const T>& v)
		{
			auto [ptr, err] = std::from_chars(str, str + len, num.val.num);
			if (ptr != str and err == std::errc()) {
				v = v.advance(static_cast<DWORD>(ptr - str));

		}
		// to_bool
		// to_date
		// ...
	private:
		virtual X to_num_(view<const T>& v) = 0;
		*/
	};


} // namespace xll
