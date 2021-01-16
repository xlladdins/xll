// registry.h - Windows registry wrappers
#include <algorithm>
#include <iterator>
#include <stdexcept>
#include <string>
#include <type_traits>
#include <vector>
#include <Windows.h>


namespace Reg {

	using SZ = std::basic_string<TCHAR>;

	// remove characters including and trailing c
	inline SZ chop(SZ& s, TCHAR c = TEXT('\0'))
	{
		auto i = s.find(c);
		if (i != s.npos) {
			s.erase(i);
		}

		return s;
	}

	// add trailing character if missing
	inline SZ tack(SZ& s, TCHAR c = TEXT('\\'))
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

#define REG_TYPE(X) \
	X(BINARY,    PBYTE,    "Binary data in any form.") \
	X(DWORD,     DWORD,    "A 32-bit number.") \
	X(EXPAND_SZ, PCTSTR,   "A null-terminated string that contains unexpanded references to environment variables") \
	X(LINK,      PCTSTR,   "A null-terminated Unicode string that contains the target path of a symbolic link that was created by calling the RegCreateKeyEx function with REG_OPTION_CREATE_LINK.") \
	X(MULTI_SZ,  PCTSTR,   "A sequence of null-terminated strings, terminated by an empty string.") \
	X(NONE,      void,     "No defined value type.") \
	X(QWORD,     LONGLONG, "A 64-bit number.") \
	X(SZ,        PCTSTR,   "A null-terminated string. This will be either a Unicode or an ANSI string, depending on whether you use the Unicode or ANSI functions.") \

	template<DWORD T> 
	struct reg_traits { };
#define REG_TYPE_TRAITS(a, b, c) template<> struct reg_traits<REG_##a> { typedef b type; };
	REG_TYPE(REG_TYPE_TRAITS)
#undef REG_TYPE_TRAITS

	// get value using subkey
	template<class T>
	inline typename T GetValue(HKEY hKey, PCTSTR lpSubKey, PCTSTR lpValue = 0);
	// set value for key opened with subkey specified
	template<class T>
	inline void SetValue(const T& t, HKEY hKey, PCTSTR lpValue = 0);

	template<>
	inline DWORD GetValue<DWORD>(HKEY hKey, PCTSTR lpSubKey, PCTSTR lpValue)
	{
		DWORD value;
		DWORD type = REG_DWORD;

		LSTATUS status = RegGetValue(hKey, lpSubKey, lpValue, RRF_RT_REG_DWORD, &type, (LPBYTE)&value, 0);
		if (ERROR_SUCCESS != status) {
			throw std::runtime_error(GetFormatMessage(status));
		}

		return value;
	}
	template<>
	inline void SetValue<DWORD>(const DWORD& dw, HKEY hKey, PCTSTR lpValue)
	{
		LSTATUS status = RegSetValueEx(hKey, lpValue, 0, REG_DWORD, (const BYTE*)&dw, sizeof(DWORD));
		if (ERROR_SUCCESS != status) {
			throw std::runtime_error(GetFormatMessage(status));
		}
	}
	// enable if convertible to DWORD???
	template<>
	inline void SetValue<int>(const int& i, HKEY hKey, PCTSTR lpValue)
	{
		SetValue(static_cast<DWORD>(i), hKey, lpValue);
	}

	template<>
	inline SZ GetValue<SZ>(HKEY hKey, PCTSTR lpSubKey, PCTSTR lpValue)
	{
		SZ value;
		DWORD type = REG_SZ;
		DWORD data = 0;

		LSTATUS status = RegGetValue(hKey, lpSubKey, lpValue, RRF_RT_REG_SZ, &type, NULL, &data);
		if (ERROR_SUCCESS != status) {
			throw std::runtime_error(GetFormatMessage(status));
		}
		size_t n = data / sizeof(TCHAR);
		value.resize(n);
		status = RegGetValue(hKey, lpSubKey, lpValue, RRF_RT_REG_SZ, &type, (LPBYTE)value.data(), &data);
		if (ERROR_SUCCESS != status) {
			throw std::runtime_error(GetFormatMessage(status));
		}

		return chop(value);
	}
	template<>
	inline void SetValue<SZ>(const SZ& sz, HKEY hKey, PCTSTR lpValue)
	{
		LSTATUS status = RegSetValueEx(hKey, lpValue, 0, REG_DWORD, (const BYTE*)sz.data(), (DWORD)sz.size());
		if (ERROR_SUCCESS != status) {
			throw std::runtime_error(GetFormatMessage(status));
		}
	}

	//!!! REG_BINARY
	//!!! std::vector<SZ> REG_MULTI_SZ

	// create or open key if it already exists
	class Key {
		HKEY hkey;
		DWORD disp;
	public:
		Key() noexcept
			: hkey(nullptr), disp(0)
		{ }
		Key(HKEY hKey, PCTSTR lpSubKey, REGSAM sam = KEY_ALL_ACCESS | KEY_WOW64_64KEY, bool open = false)
		{
			SZ subKey(lpSubKey);
			LSTATUS status;
			if (open) {
				status = RegOpenKeyEx(hKey, tack(subKey).c_str(), 0, sam, &hkey);
			}
			else {
				status = RegCreateKeyEx(hKey, tack(subKey).c_str(), 0, 0, 0, sam, 0, &hkey, &disp);
			}
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

		// get value given its name
		template<class T>
		T QueryValue(PCTSTR value);
		template<>
		DWORD QueryValue<DWORD>(PCTSTR value)
		{
			DWORD dw;
			DWORD type = REG_DWORD;
			DWORD size = sizeof(DWORD);

			LSTATUS status = RegQueryValueEx(hkey, value, 0, &type, (LPBYTE)&dw, &size);
			if (type != REG_DWORD) {
				throw std::runtime_error("Reg::Key::QueryVaue<DWORD>: type mismatch");
			}
			if (ERROR_SUCCESS != status)
			{
				throw std::runtime_error(GetFormatMessage(status));
			}

			return dw;
		}
		template<>
		int QueryValue<int>(PCTSTR value)
		{
			return static_cast<int>(QueryValue<DWORD>(value));
		}

		template<class T>
		void SetValue(const T& t, PCTSTR value)
		{
			Reg::SetValue(t, hkey, value);
		}
		template<>
		SZ QueryValue<SZ>(PCTSTR value)
		{
			SZ tstr;
			DWORD type = REG_SZ;
			DWORD size = 0;

			LSTATUS status = RegQueryValueEx(hkey, value, 0, &type, NULL, &size);
			if (type != REG_SZ) {
				throw std::runtime_error("Reg::Key::QueryVaue<SZ>: type mismatch");
			}
			if (ERROR_SUCCESS != status) {
				throw std::runtime_error(GetFormatMessage(status));
			}
			size_t n = size / sizeof(TCHAR);
			tstr.resize(n);
			status = RegQueryValueEx(hkey, value, 0, &type, (LPBYTE)tstr.data(), &size);
			if (ERROR_SUCCESS != status) {
				throw std::runtime_error(GetFormatMessage(status));
			}

			return chop(tstr);
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
				key.SetValue(t, value);

				return key;
			}
			template<class T>
			operator T() const
			{
				return key.QueryValue<T>(value);
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
			DWORD index;
			DWORD type;
			DWORD len;
			DWORD namelen = 0x3FF;
			TCHAR name[0x3FF] = TEXT(""); // name of current index
		public:
			using iterator_category = std::forward_iterator_tag;
			using value_type = std::tuple<PCTSTR,DWORD,DWORD,DWORD>; // name, index, type, len

			ValueIterator(HKEY hkey, DWORD index = 0)
				: hkey(hkey), index(index)
			{
				if (index != -1) {
					LSTATUS status = RegEnumValue(hkey, index, name, &namelen, NULL, &type, NULL, &len);
					if (ERROR_NO_MORE_ITEMS == status) {
						index = (DWORD)-1;
					}
					else if (ERROR_SUCCESS != status) {
						throw std::runtime_error(GetFormatMessage(status));
					}
				}
			}
			ValueIterator& operator=(const ValueIterator&) = default;
			ValueIterator(ValueIterator&&) = default;
			ValueIterator& operator=(ValueIterator&&) = default;
			~ValueIterator()
			{ }

			ValueIterator begin() const
			{
				return ValueIterator(hkey, 0);
			}
			ValueIterator end() const
			{
				return ValueIterator(hkey, (DWORD)-1);
			}

			explicit operator bool() const
			{
				return index != -1;
			}

			bool operator==(const ValueIterator& k) const
			{
				if (hkey != k.hkey) {
					return false;
				}
				if (index == -1 or k.index == -1) {
					return index == k.index; // ends compare equal
				}
				// same name, same value
				return 0 == std::lexicographical_compare(name, name + namelen, k.name, k.name + k.namelen);
			}

			value_type operator*() const
			{
				return std::tuple(name,index,type,len);
			}

			ValueIterator& operator++()
			{
				if (index != -1) {
					++index;
					namelen = 0x3FF;
					LSTATUS status = RegEnumValue(hkey, index, name, &namelen, NULL, &type, NULL, &len);
					if (ERROR_NO_MORE_ITEMS == status) {
						index = (DWORD)-1;
					}
					else if (ERROR_SUCCESS != status) {
						throw std::runtime_error(GetFormatMessage(status));
					}
				}

				return *this;
			}

			// stuff bits into data
			LSTATUS Value(PBYTE data)
			{
				if (index == (DWORD)-1) {
					return ERROR_NO_MORE_ITEMS;
				}

				namelen = 0x3FF;

				return RegEnumValue(hkey, index, name, &namelen, NULL, &type, data, &len);
			}
		};

		ValueIterator Values() const
		{
			return ValueIterator(*this);
		}

	};
}
