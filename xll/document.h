// document.h - Generate HTML documentation for an add-in
#pragma once
#include <fileapi.h>
//#include <format>
#include <string_view>
#include "addin.h"
#include "error.h"
#include "xllio.h"

namespace xll {

	inline const std::string documentation_html = R"xyzyx(
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <style>
        body {
            background-color: #fff;
            color: #363636;
            font-family: 'Segoe UI Light',Arial,'Helvetica Neue',Verdana,Helvetica,Sans-Serif;
            margin: 0.5in;
            padding: 0;
        }
    </style>
    <title>{{FunctionText}}</title>
</head>
<body>
    <h1>{{FunctionText}}{{MacroType}}</h1>
    <p>
        This article describes the formula syntax and usage of the {FunctionText}{MacroType} in Microsoft Excel.
    </p>
    <h2>Description</h2>
    <p>
        {{FunctionHelp}}
    </p>
    <h2>Syntax</h2>
    <p>
        {{FunctionText}}({{ArgumentHelp}})
    </p>
    <p>
        The {{FunctionText}} function syntax has the following arguments:
    </p>
    <ul>
[[
		<b>{{ArgumentName}}.</b> {{ArgumentHelp}}
]]
    </ul>
    <h2>Remarks</h2>
    <p>
    {{Documentation}}
    </p>
</body>
</html>
)xyzyx";

	class File {
		HANDLE h;
	public:
		File(LPCTSTR name, DWORD mode = GENERIC_READ | GENERIC_WRITE, DWORD disp = CREATE_ALWAYS, DWORD attr = FILE_ATTRIBUTE_NORMAL)
			: h(CreateFile(name, mode, 0, NULL, disp, attr, NULL))
		{ }
		File(File&) = delete;
		File& operator=(File&) = delete;
		~File()
		{ 
			CloseHandle(h);
		}
		BOOL Write(LPCVOID buf, DWORD len, LPDWORD plen)
		{
			return WriteFile(h, buf, len, plen, NULL);
		}
		BOOL Write(LPCVOID buf, DWORD len)
		{
			DWORD plen;

			return WriteFile(h, buf, len, &plen, NULL);
		}
		BOOL Write(const OPER& o)
		{
			return Write(o.val.str + 1, o.val.str[0] * sizeof(TCHAR));
		}

	};

	// replace first occurence of old by new in string
	inline size_t replace(std::string& s, const std::string& o, const std::string& n, size_t pos = 0)
	{
		pos = s.find(o, pos);

		if (pos == std::string::npos) {
			return pos;
		}

		s.replace(pos, pos + o.length(), n);
		
		return pos + o.length() - n.length() + 1;
	}
	// replace all occurences
	inline size_t replace_all(std::string& s, const std::string& o, const std::string& n, size_t pos = 0)
	{
		while (pos != std::string::npos) {
			pos = replace(s, o, n, pos);
		}

		return pos;
	}
	using on = std::map<std::string, std::string>;
	// replace all occurences
	inline size_t replace_all(std::string& s, const on& m, size_t pos = 0)
	{
		for (const auto& [o,n] : m) {
			pos = replace_all(s, o, n, pos);
		} 		

		return pos;
	}
	// vector replace first occurence
	inline size_t replace(std::string& s, const std::vector<on>& m, size_t pos = 0)
	{
		// [[...]]
		pos = s.find("[[", pos);
		if (pos != std::string::npos) {
			size_t pos_ = s.find("]]", pos + 2);
			std::string_view t(s.c_str() + pos + 2, pos_ - pos - 2);
			std::string n;
			for (size_t i = 0; i < m.size(); ++i) {
				std::string ti(t);
				replace_all(ti, m[i]);
				n.append(ti);
			}
			s.replace(pos, pos_ - pos + 2, n);
			pos += n.size() - (pos_ - pos + 1);
		}

		return pos;
	}
	inline size_t replace_all(std::string& s, const std::vector<on>& m, size_t pos = 0)
	{
		while (pos != std::string::npos) {
			pos = replace(s, m, pos);
		} 		

		return pos;
	}

	inline std::string to_str(const OPER4& s)
	{
		return std::string(s.val.str + 1, s.val.str[0]);
	}
	inline std::string to_str(const OPER12& s)
	{
		return utf8::wcstostring(s.val.str + 1, s.val.str[0]);
	}
	inline std::string to_str(const std::string& s)
	{
		return s;
	}
	inline std::string to_str(const std::wstring& s)
	{
		return utf8::wcstostring(s.c_str(), s.length());
	}

	// Write html documentation given Excel function text.
	inline void Document(const OPER& ft)
	{
		try {
			std::string html(documentation_html);
			const Args& arg = AddIn::Args(ft);
			const OPER& mt = arg.MacroType();
			std::string type = (mt == 1 ? " function" : mt == 2 ? " macro" : mt == 3 ? " hidden" : "!!!unknown!!");
			const on m = {
				{"{{FunctionText}}", to_str(arg.FunctionText()) },
				{"{{MacroType}}", type },
				{"{{FunctionHelp}}", to_str(arg.FunctionHelp()) },
				{"{{Documentation}}", to_str(arg.Documentation()) },
			};
			if (mt == 1) { // function
				std::vector<on> nh;
				const auto& name = arg.ArgumentName();
				const auto& help = arg.ArgumentHelp();
				for (size_t i = 0; i < name.size(); ++i) {
					on nhi = {
						{"{{ArgumentName}}", to_str(name[i])},
						{"{{ArgumentHelp}}", to_str(help[i])},
					};
					nh.push_back(nhi);
				}
				replace_all(html, nh);
			}
			replace_all(html, m);

			OPER dir = dirname(Excel(xlGetName));
			OPER file = dir & ft & OPER(".html");
			traits<XLOPERX>::xstring name(file.val.str + 1, file.val.str[0]);
			File doc(name.c_str());
			doc.Write(html.c_str(), (DWORD)html.length());

		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());
		}
	}
}
