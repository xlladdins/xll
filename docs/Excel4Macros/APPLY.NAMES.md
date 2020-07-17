APPLY.NAMES

Equivalent to clicking the Apply command on the Name submenu on the
Insert menu. Replaces definitions with their respective names. If no
names are defined in the current selection, APPLY.NAMES returns the
\#VALUE\! error value. Use APPLY.NAMES to replace references and values
in formulas with names.

**Syntax**

**APPLY.NAMES**(**name\_array**, ignore, use\_rowcol, omit\_col,
omit\_row, order\_num, append\_last)

**APPLY.NAMES**?(name\_array, ignore, use\_rowcol, omit\_col, omit\_row,
order\_num, append\_last)

Name\_array    is the name or names to apply as text elements in an
array.

  - > To give more than one name as the argument, you must use an array.
    > For example:

  - > APPLY.NAMES({"DataRange", "CriteriaRange"})

  - > If the names indicated by the argument name\_array have already
    > replaced all of the appropriate references or values, the
    > \#VALUE\! error value is returned.

>  

The next four arguments correspond to check boxes and options in the
Apply Names dialog box. Arguments that correspond to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Ignore    corresponds to the Ignore Relative/Absolute check box.

Use\_rowcol    corresponds to the Use Row And Column Names check box. If
use\_rowcol is FALSE, the next three arguments are ignored.

Omit\_col    corresponds to the Omit Column Name If Same Column check
box.

Omit\_row    corresponds to the Omit Row Name If Same Row check box.

Order\_num    determines which range name is listed first when a cell
reference is replaced by a row-oriented and a column-oriented range
name, as shown in the following table.

|                |                          |
| -------------- | ------------------------ |
| **Order\_num** | **Order of range names** |
| 1              | Row Column               |
| 2              | Column Row               |

Append\_last    determines whether the names most recently defined are
also replaced.

  - > If append\_last is TRUE, Microsoft Excel replaces the definitions
    > of the names in name\_array and also replaces the definitions of
    > the last names defined.

  - > If append\_last is FALSE or omitted, Microsoft Excel replaces the
    > definitions of the names in name\_array only.

>  

**Related Functions**

[CREATE.NAMES](CREATE.NAMES.md)   Creates names automatically from text labels on a sheet

[DEFINE.NAME](DEFINE.NAME.md)   Defines a name in the active workbook

[LIST.NAMES](LIST.NAMES.md)   Lists names and their associated information


