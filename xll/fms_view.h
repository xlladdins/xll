// fms_view.h - Platform independent view of contiguous memory
#pragma once
#include <algorithm>
#include <cstdint>
#include <cctype>
#include <algorithm>
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

		view& drop(int32_t n)
		{
			n = std::clamp(n, -(int32_t)len, (int32_t)len);

			if (n > 0) {
				buf += n;
				len -= n;
			}
			else if (n < 0) {
				len += n;
			}

			return *this;
		}
		view& take(int32_t n)
		{
			n = std::clamp(n, -(int32_t)len, (int32_t)len);

			if (n >= 0) {
				len = n;
			}
			else {
				buf += len + n;
				len = -n;
			}

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
			drop(1);

			return *this;
		}

		view& skipws()
		{
			while (len and std::isspace(front()))
				drop(1);

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
