// document.h - Generate HTML documentation for an add-in
#pragma once
#include <fileapi.h>
#include <compare>
//#include <format>
#include <string_view>
#include "error.h"
#include "addin.h"
#include "auto.h"
#include "excel.h"
//#include "xllio.h"

namespace xll {

	inline std::string dirname(const std::string& path)
	{
		char drive[_MAX_DRIVE];
		char dir[_MAX_DIR];
		char fname[_MAX_FNAME];
		char ext[_MAX_EXT];

		errno_t err = _splitpath_s(path.c_str(), drive, dir, fname, ext);
		if (err != 0) {
			throw std::runtime_error("_splitpath_s failed");
		}

		return std::string(drive).append(dir);
	}

	using string = std::string;
	using map = std::map<std::string, std::string>;

	inline const string style_css = R"(<style>
    body {
        background-color: #fff;
        color: #363636;
        font-family: 'Segoe UI Light',Arial,'Helvetica Neue',Verdana,Helvetica,Sans-Serif;
        margin: 0.5in;
        padding: 0;
    }
	table {
		text-align: left;
		border-padding: 5px;
		border-collapse: collapse;
	}
	th:first-child,td:first-child {
		padding-right: 10px;
		text-align: right;
		font-weight: bold;
	}
	tbody>tr:nth-child(odd) {
		background-color: #f2f2f2;
	}
	td:first-child {
		background-color: #fff;
	}
	.katex { 
		font-size: 1em !important;
		font-weight: 200 !important;
	}
</style>
)";

	inline const string table_html = R"xyzyx(
	<h2>Category {Category}</h2>
	<table>
	<tbody>
[[
		<tr>
			<td><a href="{FunctionText}.html">{FunctionText}</a></td>
			<td>{FunctionHelp}</td>
		</tr>
]]
	<tbody>
	</table>
)xyzyx";

	inline const string index_html = R"xyzyx(<!DOCTYPE html>
<head>
    <meta charset="UTF-8" />
	{Style}
     <title>{Category}</title>
</head>
<body>
	<h1>{Category}</h1>
	<p>
		Functions and macros of the {Category} add-in.
	</p>
	{Description}
[[
	{Table}
]]
</body>
)xyzyx";

	inline const string documentation_html = R"xyzyx(<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
	{Style}
     <title>{FunctionText}</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.12.0/dist/katex.min.css" 
		integrity="sha384-AfEj0r4/OFrOo5t7NnNe46zW/tFgW6x/bCJG8FqQCEo3+Aro6EYUG4+cU+KJWu/X" crossorigin="anonymous">
    <script defer src="https://cdn.jsdelivr.net/npm/katex@0.12.0/dist/katex.min.js" 
		integrity="sha384-g7c+Jr9ZivxKLnZTDUhnkOnsh30B4H0rpLUpJ4jAIKs4fnJI+sEnkvrMWph2EDg4" crossorigin="anonymous"></script>
    <script defer src="https://cdn.jsdelivr.net/npm/katex@0.12.0/dist/contrib/auto-render.min.js" 
		integrity="sha384-mll67QQFJfxn0IYznZYonOWZ644AWYC+Pt2cHqMaRhXVrursRwvLnLaebdGIlYNa" crossorigin="anonymous"
        onload="renderMathInElement(document.body);"></script>
</head>
<body>
    <h1>{FunctionText}{MacroType}</h1>
    <p>
        This article describes the formula syntax of the {FunctionText}{MacroType}.
    </p>
    <h2>Description</h2>
    <p>
        {FunctionHelp}
    </p>
    <h2>Syntax</h2>
    <p>
        {FunctionText}({ArgumentText})
    </p>
<blockquote>
    <table>
	<tbody>
[[

	<tr>
		<td>{ArgumentName}</td>
		<td>{ArgumentHelp}</td>
	</tr>
]]
	</tbody>
    </table>
</blockquote>
    <p>
    {Documentation}
    </p>
	<footer>
		Return to <a href="index.html">index</a>.
	</footer>
</body>
</html>
)xyzyx";

	class File {
		HANDLE h;
	public:
		File(const char* name, DWORD mode = GENERIC_READ | GENERIC_WRITE, DWORD disp = CREATE_ALWAYS, DWORD attr = FILE_ATTRIBUTE_NORMAL)
			: h(CreateFileA(name, mode, 0, NULL, disp, attr, NULL))
		{ }
		File(const string& name, DWORD mode = GENERIC_READ | GENERIC_WRITE, DWORD disp = CREATE_ALWAYS, DWORD attr = FILE_ATTRIBUTE_NORMAL)
			: File(name.c_str(), mode, disp, attr)
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
		BOOL Write(const std::string& str)
		{
			return Write(str.c_str(), (DWORD)str.length());
		}
	};

	// replace first occurence of old after pos by new in string
	inline size_t replace(string& s, const string& o, const string& n, size_t pos = 0)
	{
		pos = s.find(o, pos);

		if (pos == string::npos) {
			return pos;
		}

		s.replace(pos, o.length(), n);
		
		return pos + o.length() - n.length() + 1;
	}
	// replace all occurences
	inline size_t replace_all(string& s, const string& o, const string& n, size_t pos = 0)
	{
		while (pos != string::npos) {
			pos = replace(s, o, n, pos);
		}

		return pos;
	}
	// replace all occurences
	inline void replace_all(string& s, const map& m)
	{
		for (const auto& [o,n] : m) {
			replace_all(s, o, n);
		} 		
	}

	// vector replace first occurence after pos
	inline size_t replace(string& s, const std::vector<map>& m, size_t pos = 0)
	{
		// [[...]]
		pos = s.find("[[", pos);
		if (pos != string::npos) {
			size_t pos_ = s.find("]]", pos + 2);
			if (pos_ == string::npos) {
				throw std::runtime_error("no matching ]] for [[");
			}
			std::string_view t(s.c_str() + pos + 2, pos_ - pos - 2);
			string n;
			for (size_t i = 0; i < m.size(); ++i) {
				string ti(t);
				replace_all(ti, m[i]);
				n.append(ti);
			}
			s.replace(pos, pos_ - pos + 2, n);
			pos += n.size() - (pos_ - pos + 1);
		}

		return pos;
	}
	inline size_t replace_all(string& s, const std::vector<map>& m, size_t pos = 0)
	{
		while (pos != string::npos) {
			pos = replace(s, m, pos);
		} 		

		return pos;
	}

	inline string to_str(const OPER4& s)
	{
		return s ? string(s.val.str + 1, s.val.str[0]) : "";
	}
	inline string to_str(const OPER12& s)
	{
		return s ? utf8::wcstostring(s.val.str + 1, s.val.str[0]) : "";
	}
	inline string to_str(const string& s)
	{
		return s;
	}
	inline string to_str(const std::wstring& s)
	{
		return utf8::wcstostring(s.c_str(), s.length());
	}

	template<class X>
	inline std::vector<map> NameHelp(const XArgs<X>& arg)
	{
		std::vector<map> nh;
		const auto& name = arg.ArgumentName();
		const auto& help = arg.ArgumentHelp();
		
		for (size_t i = 0; i < name.size(); ++i) {
			nh.push_back({
				{"{ArgumentName}", to_str(name[i])},
				{"{ArgumentHelp}", to_str(help[i])},
			});
		}

		return nh;
	}

	// Write html documentation given Excel function text.
	template<class X>
	inline XOPER<X> Document(const XOPER<X>& ft)
	{
		XOPER<X> file;
		try {
			const XArgs<X>& arg = XAddIn<X>::Args(ft);
			const XOPER<X>& mt = arg.MacroType();

			string type;
			if (mt == 1)
				type = " function";
			else if (mt == 2)
				type = " macro";
			else if (mt == 3)
				type = " hidden";
			else
				throw std::runtime_error("unknown MacroType");

			string html(documentation_html);

			// replace vectors first
			replace_all(html, NameHelp(arg));

			replace_all(html, {
				{"{Style}",         style_css },
				{"{FunctionText}",  to_str(arg.FunctionText()) },
				{"{MacroType}",     type },
				{"{FunctionHelp}",  to_str(arg.FunctionHelp()) },
				{"{ArgumentText}",  to_str(arg.ArgumentText()) },
				{"{Documentation}", to_str(arg.Documentation()) },
			});

			File doc(dirname(to_str(Excel4(xlGetName))).append(to_str(ft)).append(".html"));
			doc.Write(html);
		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());
		}

		return file;
	}

	// Generate documentation for add-ins;
	inline bool Document(const string& category, const string& description = "")
	{
		// sort by Category then by FunctionText
		std::map<string,map> cat_text; // Category -> FunctionText -> FunctionHelp

		try {
			for (auto& [key, arg] : AddIn4::Map) {
				cat_text[to_str(arg.Category())][to_str(arg.FunctionText())] = to_str(arg.FunctionHelp());
				Document(key);
			}
			for (auto& [key, arg] : AddIn12::Map) {
				cat_text[to_str(arg.Category())][to_str(arg.FunctionText())] = to_str(arg.FunctionHelp());
				Document(key);
			}

			std::vector<map> tables;
			for (const auto& [cat, text_help] : cat_text) {
				std::vector<map> rows;
				for (const auto& [text, help] : text_help) {
					rows.push_back({
						{"{FunctionText}", text},
						{"{FunctionHelp}", help},
					});
				}
				string table(table_html);
				replace_all(table, "{Category}", cat);
				replace_all(table, rows);

				tables.push_back({{ "{Table}", table }});
			}

			string index(index_html);
			replace(index, "{Style}", style_css);
			replace_all(index, "{Category}", category);
			replace(index, "{Description}", description);
			replace(index, tables);
			
			File doc(dirname(to_str(Excel4(xlGetName))).append("index.html"));
			doc.Write(index);

		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());

			return false;
		}

		return true;
	}

	inline int Documentation(const char* category, const char* description = "")
	{
#ifdef _DEBUG
		Auto<OpenAfter> aoa_document([category, description]() { return Document(category, description); });
#endif		
		return 1;
	}

}
