// registry.h - Windows registry wrappers
#include <Windows.h>
#include <stdexcept>
#include <string>

namespace Reg {

	using tstring = std::basic_string<TCHAR>;

	// remove trailing characters
	inline tstring erase0(tstring& s, TCHAR c = TEXT('\0'))
	{
		s.erase(s.find(c));

		return s;
	}

	// add trailing character if missing
	inline tstring slash(tstring& s, TCHAR c = TEXT('\\'))
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

	// get value using subkey and value for HKxx keys
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
	inline tstring GetValue<tstring>(HKEY hKey, PCTSTR lpSubKey, PCTSTR lpValue)
	{
		tstring value;
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

		return erase0(value);
	}
	// SetValue
	// std::vector<tstring> REG_MULTI_SZ


	// create or open key if it already exists
	class CreateKey {
		HKEY hkey;
		DWORD disp;
	public:
		CreateKey() noexcept
			: hkey(nullptr), disp(0)
		{ }
		CreateKey(HKEY hKey, PCTSTR lpSubKey, REGSAM samDesired = KEY_ALL_ACCESS | KEY_WOW64_64KEY)
		{
			LSTATUS status = RegCreateKeyEx(hKey, lpSubKey, 0, 0, 0, samDesired, 0, &hkey, &disp);
			if (status != ERROR_SUCCESS) {
				throw std::runtime_error(GetFormatMessage(status));
			}
		}
		CreateKey(const CreateKey&) = delete;
		CreateKey& operator=(const CreateKey&) = delete;
		CreateKey(CreateKey&& h) noexcept
			: hkey(std::exchange(h.hkey, nullptr)), disp(std::exchange(h.disp, 0))
		{
		}
		CreateKey& operator=(CreateKey&& h) noexcept
		{
			if (this != &h) {
				hkey = std::exchange(h.hkey, nullptr);
				disp = std::exchange(h.disp, 0);
			}

			return *this;
		}
		~CreateKey()
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
				throw std::runtime_error("Reg::CreateKey::QueryVaue<DWORD>: type mismatch");
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
		
		struct Proxy {
			CreateKey& key;
			PCTSTR value;
			Proxy(CreateKey& key, PCTSTR value)
				: key(key), value(value)
			{ }
			template<class T>
			CreateKey& operator=(const T& t)
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
	};

}
