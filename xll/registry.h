// registry.h - Windows registry wrappers
#include <Windows.h>
#include <stdexcept>
#include <string>

namespace Win {

	inline std::string GetFormatMessage(DWORD id)
	{
		char buf[1024] = "unknown error";
		DWORD size = 1024;
		FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM, 0, id, 0, buf, size, 0);

		return std::string(buf);
	}

}

namespace Reg {
	class CreateKey {
		HKEY hkey;
		DWORD disp;
	public:
		CreateKey(HKEY hKey, PCTSTR lpSubKey)
		{
			LSTATUS status = RegCreateKeyEx(hKey, lpSubKey, 0, 0, 0, KEY_ALL_ACCESS | KEY_WOW64_64KEY, 0, &hkey, &disp);
			if (status != ERROR_SUCCESS) {
				throw std::runtime_error("CreateKey: RegCreateKeyEx failed");
			}
		}
		CreateKey(const CreateKey&) = delete;
		CreateKey& operator=(const CreateKey&) = delete;
		~CreateKey()
		{
			RegCloseKey(hkey);
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
			CreateKey& key;
			PCTSTR value;
			Proxy(CreateKey& key, PCTSTR value)
				: key(key), value(value)
			{ }
			CreateKey& operator=(DWORD dword)
			{
				LSTATUS status = RegSetValueEx(key, value, 0, REG_DWORD, (const BYTE*)&dword, sizeof(DWORD));
				if (ERROR_SUCCESS != status)
				{
					throw std::runtime_error("RegSetValueEx failed");
				}

				return key;
			}
			operator DWORD() const
			{
				DWORD dword, size = sizeof(DWORD);

				LSTATUS status = RegGetValue(key, 0, value, RRF_RT_REG_DWORD, 0, &dword, &size);
				if (ERROR_SUCCESS != status)
				{
					throw std::runtime_error("RegGetValue failed");
				}

				return dword;
			}
			// operator=(string) etc
		};
		Proxy operator[](PCTSTR value)
		{
			return Proxy(*this, value);
		}
	};

	template<class T>
	inline typename T QueryValue(HKEY hKey, PCTSTR lpValue);
	
	template<>
	inline DWORD QueryValue<DWORD>(HKEY hKey, PCTSTR lpValue)
	{
		DWORD value;
		DWORD type = REG_DWORD;

		LSTATUS status = RegQueryValueEx(hKey, lpValue, 0, &type, (LPBYTE)&value, 0);
		if (ERROR_SUCCESS != status) {
			throw Win::GetFormatMessage(status);
		}

		return value;
	}
	
	template<>
	inline std::basic_string<TCHAR> QueryValue<std::basic_string<TCHAR>>(HKEY hKey, PCTSTR lpValue)
	{
		std::basic_string<TCHAR> value;
		DWORD type = REG_SZ;
		DWORD data = 0;

		LSTATUS status = RegQueryValueEx(hKey, lpValue, NULL, &type, NULL, &data);
		if (ERROR_SUCCESS != status) {
			throw Win::GetFormatMessage(status);
		}
		size_t n = data / sizeof(TCHAR);
		value.resize(n);
		status = RegQueryValueEx(hKey, lpValue, NULL, &type, (LPBYTE)value.data(), &data);
		if (ERROR_SUCCESS != status) {
			throw Win::GetFormatMessage(status);
		}

		return value.erase(value.find('\0'));
	}

	template<class T>
	inline typename T GetValue(HKEY hKey, PCTSTR lpSubKey, PCTSTR lpValue = 0);

	template<>
	inline DWORD GetValue<DWORD>(HKEY hKey, PCTSTR lpSubKey, PCTSTR lpValue)
	{
		DWORD value;
		DWORD type = REG_DWORD;

		LSTATUS status = RegGetValue(hKey, lpSubKey, lpValue, RRF_RT_REG_DWORD, &type, (LPBYTE)&value, 0);
		if (ERROR_SUCCESS != status) {
			throw Win::GetFormatMessage(status);
		}

		return value;
	}

	template<>
	inline std::basic_string<TCHAR> GetValue<std::basic_string<TCHAR>>(HKEY hKey, PCTSTR lpSubKey, PCTSTR lpValue)
	{
		std::basic_string<TCHAR> value;
		DWORD type = REG_SZ;
		DWORD data = 0;

		LSTATUS status = RegGetValue(hKey, lpSubKey, lpValue, RRF_RT_REG_SZ, &type, NULL, &data);
		if (ERROR_SUCCESS != status) {
			throw Win::GetFormatMessage(status);
		}
		size_t n = data / sizeof(TCHAR);
		value.resize(n);
		status = RegGetValue(hKey, lpSubKey, lpValue, RRF_RT_REG_SZ, &type, (LPBYTE)value.data(), &data);
		if (ERROR_SUCCESS != status) {
			throw Win::GetFormatMessage(status);
		}

		return value.erase(value.find('\0'));
	}
}
