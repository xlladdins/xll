// registry.h - Windows registry wrappers
#pragma once
#include <algorithm>
#include <compare>
#include <iterator>
#include <stdexcept>
#include <string>
#include <tuple>
#include <type_traits>
#include <vector>
#include <Windows.h>
#include <tchar.h>

namespace Reg {

	// remove characters from first c to end
	inline std::basic_string<TCHAR> chop(std::basic_string<TCHAR>& s, TCHAR c = TEXT('\0'))
	{
		auto i = s.find(c);
		if (i != s.npos) {
			s.erase(i);
		}

		return s;
	}

	// add trailing character if missing
	inline std::basic_string<TCHAR> tack(std::basic_string<TCHAR>& s, TCHAR c = TEXT('\\'))
	{
		if (s.length() != 0 and s.back() != c) {
			s.push_back(c);
		}

		return s;
	}
	

	inline char* GetFormatMessage(DWORD id)
	{
		// not thread safe
		static constexpr DWORD size = 1024;
		static char buf[size];

		FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM, 0, id, 0, buf, size, 0);

		return buf;
	}

#define REG_KEY(X) \
	X(HKCR, CLASSES_ROOT, "Defines types (or classes) of documents and the properties associated with those types.") \
	X(HKCC, CURRENT_CONFIG, "Contains information about the current hardware profile of the local computer system.") \
	X(HKCU, CURRENT_USER, "Defines the preferences of the current user.") \
	X(HKLS, CURRENT_USER_LOCAL_SETTINGS, "Defines preferences of the current user that are local to the machine.") \
	X(HKLM, LOCAL_MACHINE, "Defines the physical state of the computer.") \
	X(HKPD, PERFORMANCE_DATA, "Allows you to access performance data.") \
	X(HKPN, PERFORMANCE_NLSTEXT, "References the text strings that describe counters.") \
	X(HKPT, PERFORMANCE_TEXT, "References the text strings that describe counters in US English.") \
	X(HKUS, USERS, "Defines the default user configuration for new users on the local computer and the user configuration for the current user.") \

#define REG_SAM(X) \
	X(ALL_ACCESS, "Combines the STANDARD_RIGHTS_REQUIRED, KEY_QUERY_VALUE, KEY_SET_VALUE, KEY_CREATE_SUB_KEY, KEY_ENUMERATE_SUB_KEYS, KEY_NOTIFY, and KEY_CREATE_LINK access rights.") \
	X(CREATE_LINK, "Reserved for system use.") \
	X(CREATE_SUB_KEY, "Required to create a subkey of a registry key.") \
	X(ENUMERATE_SUB_KEYS, "Required to enumerate the subkeys of a registry key.") \
	X(EXECUTE, "Equivalent to KEY_READ.") \
	X(NOTIFY, "Required to request change notifications for a registry key or for subkeys of a registry key.") \
	X(QUERY_VALUE, "Required to query the values of a registry key.") \
	X(READ, "Combines the STANDARD_RIGHTS_READ, KEY_QUERY_VALUE, KEY_ENUMERATE_SUB_KEYS, and KEY_NOTIFY values.") \
	X(SET_VALUE, "Required to create, delete, or set a registry value.") \
	X(WOW64_32KEY, "Indicates that an application on 64-bit Windows should operate on the 32-bit registry view. This flag is ignored by 32-bit Windows. For more information, see Accessing an Alternate Registry View.") \
	X(WOW64_64KEY, "Indicates that an application on 64-bit Windows should operate on the 64-bit registry view. This flag is ignored by 32-bit Windows. For more information, see Accessing an Alternate Registry View.") \
	X(WRITE, "Combines the STANDARD_RIGHTS_WRITE, KEY_SET_VALUE, and KEY_CREATE_SUB_KEY access rights.") \

#define REG_SAM_ENUM(a,b) a = KEY_##a,
	enum KEY {
		REG_SAM(REG_SAM_ENUM)
	};
#undef REG_SAM_ENUM

#define REG_TYPES(X) \
	X(NONE,      void,     "No defined value type.") \
	X(BINARY,    PBYTE,    "Binary data in any form.") \
	X(DWORD,     DWORD,    "A 32-bit number.") \
	X(EXPAND_SZ, PCTSTR,   "A null-terminated string that contains unexpanded references to environment variables") \
	X(LINK,      PCTSTR,   "A null-terminated Unicode string that contains the target path of a symbolic link that was created by calling the RegCreateKeyEx function with REG_OPTION_CREATE_LINK.") \
	X(MULTI_SZ,  PCTSTR,   "A sequence of null-terminated strings, terminated by an empty string.") \
	X(QWORD,     LONGLONG, "A 64-bit number.") \
	X(SZ,        PCTSTR,   "A null-terminated string. This will be either a Unicode or an ANSI string, depending on whether you use the Unicode or ANSI functions.") \

#define REG_TYPE_ENUM(a, b, c) a = REG_##a,
	enum class REG_TYPE {
		REG_TYPES(REG_TYPE_ENUM)
	};
#undef REG_TYPE_ENUM

	/*
	template<REG_TYPE T> 
	struct reg_traits { };
#define REG_TYPE_TRAITS(a, b, c) template<> struct reg_traits<REG_##a> { typedef b type; };
	REG_TYPE(REG_TYPE_TRAITS);
#undef REG_TYPE_TRAITS
	*/

	// Registry value variant type
	struct Value {
		std::basic_string<TCHAR> name;
		DWORD index;
		DWORD type;
		std::vector<BYTE> data;

		// Defaults to default value.
		Value(PCTSTR name = TEXT(""), DWORD type = REG_NONE, PBYTE data = nullptr, DWORD len = 0)
			: name(name), index((DWORD)-1), type(type), data(data, data + len)
		{ }
		Value(const Value&) = default;
		Value& operator=(const Value&) = default;
		Value& operator=(DWORD dw)
		{
			type = REG_DWORD;
			data.resize(sizeof(DWORD));
			CopyMemory(data.data(), &dw, data.size());

			return *this;
		}
		Value& operator=(PCTSTR sz)
		{
			type = REG_SZ;
			data.resize((1 + _tcsclen(sz)) * sizeof(TCHAR));
			CopyMemory(data.data(), sz, data.size());

			return *this;
		}
		~Value()
		{ }

		auto operator<=>(const Value&) const = default;

		Value& Query(HKEY hKey)
		{
			LSTATUS status;
			DWORD len;
			status = RegQueryValueEx(hKey, name.c_str(), NULL, &type, NULL, &len);
			if (ERROR_SUCCESS != status) {
				throw std::runtime_error(GetFormatMessage(status));
			}
			data.resize(len);
			status = RegQueryValueEx(hKey, name.c_str(), NULL, &type, data.data(), &len);
			if (ERROR_SUCCESS != status) {
				throw std::runtime_error(GetFormatMessage(status));
			}

			return *this;
		}

		Value& Set(HKEY hKey)
		{
			LSTATUS status = RegSetValueEx(hKey, name.c_str(), 0, type, data.data(), (DWORD)data.size());
			if (ERROR_SUCCESS != status) {
				throw std::runtime_error(GetFormatMessage(status));
			}

			return *this;
		}

		operator const DWORD() const
		{
			if (type != REG_DWORD) {
				throw std::runtime_error("Reg::Value: wrong type for DWORD");
			}

			return static_cast<DWORD>(*data.data());
		}
		operator const TCHAR*() const
		{
			if (type != REG_SZ) { // multi,...
				throw std::runtime_error("Reg::Value: wrong type for REG_SZ");
			}

			return (const TCHAR*)(data.data());
		}
		// QWORD, BINARY, ...
	};

	// create or open key
	class Key {
		HKEY hkey = nullptr;
		DWORD disp = 0;
	public:
		Key() noexcept
			: hkey(nullptr), disp(0)
		{ }
		Key(HKEY hKey, PCTSTR lpSubKey, REGSAM sam = KEY_ALL_ACCESS | KEY_WOW64_64KEY)
			: disp(0)
		{
			LSTATUS status = RegCreateKeyEx(hKey, lpSubKey, 0, 0, 0, sam, 0, &hkey, &disp);
			if (status != ERROR_SUCCESS) {
				throw std::runtime_error(GetFormatMessage(status));
			}
		}
		Key(const Key&) = delete;
		Key& operator=(const Key&) = delete;
		Key(Key&& h) noexcept
			: hkey(std::exchange(h.hkey, nullptr)), disp(std::exchange(h.disp, 0))
		{ }
		Key& operator=(Key&& h) noexcept
		{
			if (this != &h) {
				hkey = std::exchange(h.hkey, nullptr);
				disp = std::exchange(h.disp, 0);
			}

			return *this;
		}
		~Key()
		{
			if (hkey and disp) {
				RegCloseKey(hkey);
			}
		}
		explicit operator bool() const
		{
			return hkey != nullptr;
		}
		operator HKEY() const
		{
			return hkey;
		}
		DWORD disposition() const
		{
			return disp;
		}

		struct Proxy {
			Key& key;
			PCTSTR value;
			Proxy(Key& key, PCTSTR value)
				: key(key), value(value)
			{ }
			template<class T>
			Key& operator=(const T& t)
			{
				Value val(value);
				val = t;
				val.Set(key);

				return key;
			}
			template<class T>
			operator T() const
			{
				Value val(value);

				return val.Query(key);
			}
		};
		Proxy operator[](PCTSTR value)
		{
			return Proxy(*this, value);
		}

		// iterator over key names
		// ??? may want to use RegNotifyChangeKeyValue 
		class KeyIterator {
			DWORD index;
			HKEY hkey;
			TCHAR name[MAX_PATH] = TEXT(""); // name of current index
			DWORD namelen = MAX_PATH;
		public:
			using iterator_category = std::forward_iterator_tag;
			using value_type = PCTSTR;

			KeyIterator(HKEY hkey, bool end = false)
				: index((DWORD)-1), hkey(hkey)
			{	
				if (end) {
					name[0] = 0;
				}
				else {
					next();
				}
			}
			KeyIterator begin() const
			{
				return KeyIterator(hkey);
			}
			KeyIterator end() const
			{
				return KeyIterator(hkey, true);
			}
			explicit operator bool() const
			{
				return index != -1;
			}
			bool operator==(const KeyIterator& k) const
			{
				if (hkey != k.hkey) {
					return false;
				}
				if (index == -1 or k.index == -1) {
					return index == k.index; // ends compare equal
				}

				return 0 == std::lexicographical_compare(name, name + namelen, k.name, k.name + k.namelen);
			}
			// key name
			value_type operator*() const
			{
				return name;
			}
			DWORD length() const
			{
				return namelen;
			}
			
			KeyIterator& operator++()
			{
				return next();
			};
		private:
			KeyIterator& next()
			{
				++index;
				name[0] = 0;
				namelen = MAX_PATH;
				LSTATUS status = RegEnumKeyEx(hkey, index, name, &namelen, NULL, NULL, NULL, NULL);
				if (ERROR_NO_MORE_ITEMS == status) {
					index = (DWORD)-1;
				}
				else if (ERROR_SUCCESS != status) {
					throw std::runtime_error(GetFormatMessage(status));
				}

				return *this;
			}
		};

		KeyIterator Keys()
		{
			return KeyIterator(*this);
		};

		// iterator over key values
		class ValueIterator {
			HKEY hkey;
			Value val;
		public:
			using iterator_category = std::forward_iterator_tag;
			using value_type = Value;

			ValueIterator(HKEY hkey)
				: hkey(hkey)
			{
				val.index = 0;
				EnumValue();
			}
			ValueIterator(const ValueIterator&) = default;
			ValueIterator& operator=(const ValueIterator&) = default;
			~ValueIterator()
			{ }

			auto operator<=>(const ValueIterator&) const = default;

			ValueIterator begin() const
			{
				return *this;
			}
			ValueIterator end() const
			{
				ValueIterator v(*this);
				v.val.index = (DWORD)-1;

				return v;
			}

			explicit operator bool() const
			{
				return val.index != -1;
			}

			const value_type& operator*() const
			{
				return val;
			}

			ValueIterator& operator++()
			{
				++val.index;
				EnumValue();

				return *this;
			}
		private:
			LSTATUS EnumValue()
			{
				LSTATUS status;
				
				DWORD datalen = 0;
				DWORD namelen = 0x3ff;
				val.name.resize(0x3ff);
				status = RegEnumValue(hkey, val.index, val.name.data(), &namelen, 0, &val.type, NULL, &datalen);
				val.name.resize(namelen);
				if (ERROR_NO_MORE_ITEMS == status) {
					val.index = (DWORD)-1;
				}
				else {
					val.data.resize(datalen);
					status = RegEnumValue(hkey, val.index, val.name.data(), &namelen, NULL, &val.type, val.data.data(), &datalen);
				}

				return status;
			}
		};

		ValueIterator Values() const
		{
			return ValueIterator(*this);
		}

	};
}
