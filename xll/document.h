// document.h - Generate HTML documentation for an add-in
#pragma once
#include <fileapi.h>
#include <compare>
//#include <format>
#include <string_view>
#include "error.h"
#include "addin.h"
#include "xllio.h"

namespace xll {

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
</style>
)";

	inline const string table_html = R"xyzyx(
	<h2>Category {Category}</h2>
	<table>
	<thead>
		<tr>
			<th>Name</th>
			<th>Description</th>
		</tr>
	</thead>
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
</head>
<body>
    <h1>{FunctionText}{MacroType}</h1>
    <p>
        This article describes the formula syntax and usage of the {FunctionText}{MacroType} in Microsoft Excel.
    </p>
    <h2>Description</h2>
    <p>
        {FunctionHelp}
    </p>
    <h2>Syntax</h2>
    <p>
        {FunctionText}({ArgumentText})
    </p>
    <p>
        The {FunctionText} function syntax has the following arguments:
    </p>
    <table>
	<thead>
	<tr>
		<th>Name</th>
		<th>Description</th>
	</tr>
	</thead>
	<tbody>
[[

	<tr>
		<td>{ArgumentName}</td>
		<td>{ArgumentHelp}</td>
	</tr>
]]
	</tbody>
    </table>
    <h2>Remarks</h2>
    <p>
    {Documentation}
    </p>
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
		if (!s)
			return "";

		return string(s.val.str + 1, s.val.str[0]);
	}
	inline string to_str(const OPER12& s)
	{
		if (!s)
			return "";

		return utf8::wcstostring(s.val.str + 1, s.val.str[0]);
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
			map nhi = {
				{"{ArgumentName}", to_str(name[i])},
				{"{ArgumentHelp}", to_str(help[i])},
			};
			nh.push_back(nhi);
		}

		return nh;
	}

	// Write html documentation given Excel function text.
	template<class X>
	inline XOPER<X> Document(const XOPER<X>& ft)
	{
		XOPER<X> file;
		try {
			string html(documentation_html);
			const XArgs<X>& arg = XAddIn<X>::Args(ft);
			const XOPER<X>& mt = arg.MacroType();

			replace_all(html, "{Style}", style_css);

			string type;
			if (mt == 1)
				type = " function";
			else if (mt == 2)
				type = " macro";
			else if (mt == 3)
				type = " hidden";
			else
				throw std::runtime_error("unknown MacroType");

			if (mt == 1) { // function
				// replace vectors first
				replace_all(html, NameHelp(arg));
			}

			const map m = {
				{"{FunctionText}", to_str(arg.FunctionText()) },
				{"{MacroType}", type },
				{"{FunctionHelp}", to_str(arg.FunctionHelp()) },
				{"{ArgumentText}", to_str(arg.ArgumentText()) },
				{"{Documentation}", to_str(arg.Documentation()) },
			};
			replace_all(html, m);

			XOPER<X> dir = dirname(XExcel<X>(xlGetName));
			file = dir & ft & XOPER<X>(".html");
			File doc(to_str(file));
			doc.Write(html);
		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());
		}

		return file;
	}

	template<class X>
	inline map TextHelp(const XAddIn<X>& a)
	{
		map m;

		for (const auto& [_, v] : a.Map()) {
			m[to_str(v.FunctionText())] = to_str(v.FunctionHelp());
		}

		return m;
	}
	template<class X>
	inline map CategoryText(const XAddIn<X>& a)
	{
		map m;

		for (const auto& [_, v] : a.Map()) {
			m[to_str(v.Category())] = to_str(v.FunctionText());
		}

		return m;
	}

	// Index page for documentaton.
	inline bool Document(const string& category)
	{
		// sort by Category then by FunctionText
		std::map<string,map> cat_text; // Category -> FunctionText -> FunctionHelp

		try {
			for (auto& [key, arg] : AddIn4::Map) {
				cat_text[to_str(arg.Category())][to_str(arg.FunctionText())] = to_str(arg.FunctionHelp());
			}
			for (auto& [key, arg] : AddIn12::Map) {
				cat_text[to_str(arg.Category())][to_str(arg.FunctionText())] = to_str(arg.FunctionHelp());
			}

			std::vector<map> table;
			for (const auto& [cat, text_help] : cat_text) {
				std::vector<map> rows;
				for (const auto& [text, help] : text_help) {
					const map td = {
						{"{FunctionText}", text},
						{"{FunctionHelp}", help},
					};
					rows.push_back(td);
				}
				string t(table_html);
				replace(t, "{Category}", cat);
				replace(t, rows);

				const map r = {
					{ "{Table}", t },
				};
				table.push_back(r);
			}

			string index(index_html);
			replace(index, "{Style}", style_css);
			replace_all(index, "{Category}", category);
			replace(index, table);
			
			OPER4 dir = dirname(Excel4(xlGetName));
			string file = to_str(dir & OPER4("index.html"));
			File doc(file);
			doc.Write(index);

		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());

			return false;
		}

		return true;
	}
}
