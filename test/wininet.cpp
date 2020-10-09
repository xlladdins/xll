// wininet.cpp -
// https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetopena
//#define XLOPERX XLOPER
#include "../xll/xll.h"
#include <WinInet.h>

#pragma comment(lib, "wininet.lib")

class Internet {
	class Open {
		HINTERNET h;
	public:
		Open()
			: h{ InternetOpen(TEXT("XLL"), INTERNET_OPEN_TYPE_PRECONFIG, nullptr, nullptr, 0) }
		{ }
		~Open()
		{
			InternetCloseHandle(h);
		}
		operator HINTERNET()
		{
			return h;
		}
	};
	public:
	class OpenUrl {
		static inline Open hOpen;
		HINTERNET hUrl;
	public:
		OpenUrl(LPCTSTR url = nullptr, LPCTSTR headers = nullptr, DWORD len = 0L, DWORD flags = 0, DWORD_PTR context = 0)
		: hUrl{ InternetOpenUrl(hOpen, url, headers, len, flags, context) }
		{
			DWORD err;
			if (!url) {
				err = GetLastError();
			}
		}
		OpenUrl(const OpenUrl&) = delete;
		OpenUrl& operator=(const OpenUrl&) = delete;
		~OpenUrl()
		{
			InternetCloseHandle(hUrl);
		}
		BOOL ReadFile(LPVOID buf, DWORD n, DWORD* pn)
		{
			return InternetReadFile(hUrl, buf, n, pn);
		}
		std::vector<unsigned char> Read()
		{
			std::vector<unsigned char> buf;
			DWORD n = BUFSIZ;
			DWORD n_ = 0;
			buf.resize(n);
			DWORD off = 0;
			//BOOL b;
			while (ReadFile(&buf[off], n, &n_)) {
				if (n_ < n) {
					buf.resize(off + n_);
					
					break;
				}
				off += n;
				buf.resize(off + n);
			}
			//!!! read one more time and assert n_ == 0 ???

			return buf;
		}
	};
};

using namespace xll;

AddIn xai_read_url(
	Function(XLL_CSTRING4, "xll_read_url", "XLL.READ.URL")
	.Args({
		Arg(XLL_CSTRING, "url", "read url")
		})
	.Category("XLL")
	.FunctionHelp("Read a url")
);
LPCSTR WINAPI xll_read_url(traits<XLOPERX>::xcstr url)
{
#pragma XLLEXPORT
	static std::vector<unsigned char> v;
	auto io = Internet::OpenUrl(url);
	v = io.Read();

	return (LPCSTR)v.data();
}