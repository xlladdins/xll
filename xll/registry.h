// registry.h - Windows registry wrappers
#include <Windows.h>
#include <stdexcept>

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
}
