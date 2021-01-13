// registry.h - Windows registry wrappers
#include <algorithm>
#include <iterator>
#include <stdexcept>
#include <string>
#include <type_traits>
#include <vector>
#include <Windows.h>

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
		if (s.back() != c) {
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

#define REG_SAM_ENUM(a,b) a = KEY_##a,
	enum KEY {
		REG_SAM(REG_SAM_ENUM)
	};
#undef REG_SAM_ENUM

	template<DWORD T> 
	struct reg_traits { };
	template<> 
	struct reg_traits<REG_DWORD> { typedef DWORD type;  };
	template<>
	struct reg_traits<REG_QWORD> { typedef ULONGLONG type; };
	template<>
	struct reg_traits<REG_SZ> { typedef PCTSTR type; };
	template<>
	struct reg_traits<REG_BINARY> { typedef LPBYTE type; };
	//!!!...

	// get value using subkey for HKxx keys
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
		Key(HKEY hKey, PCTSTR lpSubKey, REGSAM samDesired = KEY_ALL_ACCESS | KEY_WOW64_64KEY)
		{
			SZ subKey(lpSubKey);
			LSTATUS status = RegCreateKeyEx(hKey, tack(subKey).c_str(), 0, 0, 0, samDesired, 0, &hkey, &disp);
			if (status != ERROR_SUCCESS) {
				throw std::runtime_error(GetFormatMessage(status));
			}
		}
		Key(const Key&) = delete;
		Key& operator=(const Key&) = delete;
		Key(Key&& h) noexcept
			: hkey(std::exchange(h.hkey, nullptr)), disp(std::exchange(h.disp, 0))
		{
		}
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

		// get value if key opened with specified subkey
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

		KeyIterator begin() const
		{
			return KeyIterator(*this);
		}
		KeyIterator end() const
		{
			return KeyIterator(*this, true);
		}

		// iterator over key values
		class ValueIterator {
			DWORD index;
			HKEY hkey;
			TCHAR name[0x3FF] = TEXT(""); // name of current index
			DWORD namelen = 0x3FF;
			DWORD type;
			std::vector<BYTE> buf;
		public:
			using iterator_category = std::forward_iterator_tag;
			using value_type = PCTSTR;

			ValueIterator(HKEY hkey, bool end = false)
				: index((DWORD)-1), hkey(hkey)
			{
				if (end) {
					name[0] = 0;
				}
				else {
					next();
				}
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
			// key name
			value_type operator*() const
			{
				return name;
			}
			DWORD length() const
			{
				return namelen;
			}
			DWORD Type() const
			{
				return type;
			}

			ValueIterator& operator++()
			{
				return next();
			};
			BYTE* data()
			{
				return &buf[0];
			}
			/*
			template<class T>
			const T& Value() const
			{
				static_assert(std::is_same_v<T, reg_traits<T>>::type);

				return *static_cast<const T*>(&data[0]);
			}
			*/
		private:
			ValueIterator& next()
			{
				++index;
				name[0] = 0;
				namelen = 0x3FF;
				DWORD len = 0;
				LSTATUS status = RegEnumValue(hkey, index, name, &namelen, NULL, &type, NULL, &len);
				if (ERROR_NO_MORE_ITEMS == status) {
					index = (DWORD)-1;

					return *this;
				}
				else if (ERROR_SUCCESS != status) {
					throw std::runtime_error(GetFormatMessage(status));
				}
				buf.resize(len);
				status = RegEnumValue(hkey, index, name, &namelen, NULL, &type, buf.data(), &len);
				if (ERROR_SUCCESS != status) {
					throw std::runtime_error(GetFormatMessage(status));
				}

				return *this;
			}
		};

	};
}
