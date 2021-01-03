// document.h - Generate HTML documentation for an add-in
#pragma once
#include <fileapi.h>
#include <compare>
//#include <format>
#include <fstream>
#include <string_view>
#include "error.h"
#include "addin.h"
#include "auto.h"
#include "excel.h"

// TODO use ostream instead of replace.

namespace xll {

	struct splitpath {
		char drive[_MAX_DRIVE];
		char dir[_MAX_DIR];
		char fname[_MAX_FNAME];
		char ext[_MAX_EXT];
		splitpath(const char* path)
		{
			errno_t err = _splitpath_s(path, drive, dir, fname, ext);
			if (err != 0) {
				throw std::runtime_error("_splitpath_s failed");
			}
		}
		std::string dirname() const
		{
			return std::string(drive) + std::string(dir);
		}
	};

	/*
	using string = std::string;
	using map = std::map<std::string, std::string>;

	template<class X>
	inline std::vector<map> NameHelp(const XArgs<X>& arg)
	{
		std::vector<map> nh;
		const auto& name = arg.ArgumentName();
		const auto& help = arg.ArgumentHelp();

		for (size_t i = 0; i < name.size(); ++i) {
			nh.push_back({
				{"{ArgumentName}", name[i].to_string()},
				{"{ArgumentHelp}", help[i].to_string()},
				});
		}

		return nh;
	}


	inline const string style_css = R"(<style>
    body {
        background-color: #fff;
        color: #363636;
        font-family: 'Segoe UI',Calibri,Arial,'Helvetica Neue',Verdana,Helvetica,Sans-Serif;
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
	*/

	// Write html documentation given Excel function text.
	template<class X>
	inline XOPER<X> Document(const XOPER<X>& functionText)
	{
		XOPER<X> file;
		try {
			// factor out!!!
			splitpath sp(Excel4(xlGetName).to_string().c_str());
			std::string ofile(sp.dirname() + sp.fname + ".html");
			std::ofstream ofs(ofile, std::ios::out);

			const XArgs<X>& arg = XAddIn<X>::Args(functionText);
			size_t foo;
			foo = arg.ArgumentCount();

			ofs << R"xyzyx(<!DOCTYPE html>
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
			/*
			// replace vectors first
			replace_all(html, NameHelp(arg));

			replace_all(html, {
				{"{Style}",         style_css },
				{"{FunctionText}",  arg.FunctionText().to_string() },
				{"{MacroType}",     type },
				{"{FunctionHelp}",  arg.FunctionHelp().to_string() },
				{"{ArgumentText}",  arg.ArgumentText().to_string() },
				{"{Documentation}", arg.Documentation().to_string() },
			});
			*/
		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());
		}

		return file;
	}

	// Generate documentation for add-ins;
	inline bool Document(const char* category="", const char* description = "")
	{
		char c;
		c = *category;
		c = *description;
		// sort by Category then by FunctionText
		std::map<OPER,std::map<OPER,OPER>> cat_text; // Category -> FunctionText -> FunctionHelp

		try {
			for (auto& [key, arg] : AddIn::Map) {
				cat_text[arg.Category()][arg.FunctionText()] = arg.FunctionHelp();
				Document(key);
			}
/*
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
			
			splitpath sp(Excel4(xlGetName)).c_str().to_string();
			string ofile(sp.dirname() + "index.html");
			std::ofstream ofs(ofile, std::ios::out);
			ofs.write(index.c_str(), index.length());
			*/
		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());

			return false;
		}

		return true;
	}

	inline int Documentation([[maybe_unused]] const char* category, [[maybe_unused]] const char* description = "")
	{
#ifdef _DEBUG
		Auto<OpenAfter> aoa_document([category, description]() { return Document(category, description); });
#endif		
		return 1;
	}

}
