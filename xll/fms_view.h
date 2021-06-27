// fms_view.h - Platform independent view of contiguous memory
#pragma once
#include <cstdint>
#include <cctype>
#include <compare>
#include <stdexcept>
#include <utility>

namespace fms {

	/// <summary>
	/// View of contiguous memory.
	/// </summary>
	/// <typeparam name="T">type of data</typeparam>
	template<class T>
	struct view {
		using view_type = T;
		T* buf;
		uint32_t len;

		view()
			: buf(nullptr), len(0)
		{ }
		view(T* buf, uint32_t len)
			: buf(buf), len(len)
		{ }
		template<size_t N>
		view(T(&buf)[N])
			: view(buf, static_cast<uint32_t>(N))
		{
			if (len)
				--len; // ignore terminating null
		}
		view(const view&) = default;
		view& operator=(const view&) = default;
		virtual ~view()
		{ }

		explicit operator bool() const
		{
			return len != 0;
		}

		// compare views
		auto operator<=>(const view&) const = default;

		// compare contents
		bool equal(const view& v) const
		{
			return len == v.len and len == 0 or std::equal(buf, buf + len, v.buf);
		}

		// do not advance past end
		view& advance(uint32_t n)
		{
			if (n > len) 
				n = len;
			buf += n;
			len -= n;

			return *this;
		}

		T back() const
		{
			return buf[len - 1];
		}
		T front() const
		{
			return buf[0];
		}

		view& eat(T c)
		{
			if (front() != c) {
				throw std::runtime_error("view: eat cannot digest character");
			}
			advance(1);

			return *this;
		}

		view& skipws()
		{
			while (std::isspace(front()))
				advance(1);

			return *this;
		}
		view& trimws()
		{
			while (len and std::isspace(back())) {
				--len;
			}

			return *this;
		}
	};

} // namespace xll
