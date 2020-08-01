# CONSOLIDATE

Equivalent to clicking the Consolidate command on the Data menu.
Consolidates data from multiple ranges on multiple worksheets into a
single range on a single worksheet.

**Syntax**

**CONSOLIDATE**(source\_refs, function\_num, top\_row, left\_col,
create\_links)

**CONSOLIDATE**?(source\_refs, function\_num, top\_row, left\_col,
create\_links)

Source\_refs&nbsp;&nbsp;&nbsp;&nbsp;are references to areas that contain
data to be consolidated on the destination worksheet. Source\_refs must
be in text form and include the full path of the file and the cell
reference or named ranges in the workbook to be consolidated.
Source\_refs are usually external references and must be given as an
array, for example: {"SHEET1\!IncomeOne", "SHEET2\!IncomeTwo"}.

To add or delete source\_refs from an existing consolidation on a
worksheet, reuse the CONSOLIDATE function, specifying the new
source\_refs.

Function\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 11 that
specifies one of the 11 functions you can use to consolidate data. If
function\_num is omitted, the SUM function, number 9, is used. The
functions and their corresponding numbers are listed in the following
table.

|                   |              |
| ----------------- | ------------ |
| **Function\_num** | **Function** |
| 1                 | AVERAGE      |
| 2                 | COUNT        |
| 3                 | COUNTA       |
| 4                 | MAX          |
| 5                 | MIN          |
| 6                 | PRODUCT      |
| 7                 | STDEV        |
| 8                 | STDEVP       |
| 9                 | SUM          |
| 10                | VAR          |
| 11                | VARP         |

The following arguments correspond to text boxes and check boxes in the
Consolidate dialog box. Arguments that correspond to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Top\_row&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Top Row check box.
The default is FALSE.

Left\_col&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Left Column check
box. The default is FALSE.

If top\_row and left\_col are both FALSE or omitted, the data is
consolidated by position.

Create\_links&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Create Links To
Source Data check box.

**Remarks**

  - > If you use the CONSOLIDATE function with no arguments and there is
    > a consolidation on the active worksheet, Microsoft Excel
    > reconsolidates, using the sources, function, and position
    > attributes used to create the existing consolidation.

  - > If you use the CONSOLIDATE function with no arguments and there is
    > no consolidation on the active worksheet, the function returns the
    > \#VALUE\! error value.


**Related Functions**

[CHANGE.LINK](CHANGE.LINK.md)&nbsp;&nbsp;&nbsp;Changes supporting workbook links

[LINKS](LINKS.md)&nbsp;&nbsp;&nbsp;Returns the names of all linked workbooks

[OPEN.LINKS](OPEN.LINKS.md)&nbsp;&nbsp;&nbsp;Opens specified supporting workbooks

[UPDATE.LINK](UPDATE.LINK.md)&nbsp;&nbsp;&nbsp;Updates a link to another workbook



Return to [README](README.md#C)

# CONSOLIDATE

Equivalent to clicking the Consolidate command on the Data menu.
Consolidates data from multiple ranges on multiple worksheets into a
single range on a single worksheet.

**Syntax**

**CONSOLIDATE**(source\_refs, function\_num, top\_row, left\_col,
create\_links)

**CONSOLIDATE**?(source\_refs, function\_num, top\_row, left\_col,
create\_links)

Source\_refs&nbsp;&nbsp;&nbsp;&nbsp;are references to areas that contain
data to be consolidated on the destination worksheet. Source\_refs must
be in text form and include the full path of the file and the cell
reference or named ranges in the workbook to be consolidated.
Source\_refs are usually external references and must be given as an
array, for example: {"SHEET1\!IncomeOne", "SHEET2\!IncomeTwo"}.

To add or delete source\_refs from an existing consolidation on a
worksheet, reuse the CONSOLIDATE function, specifying the new
source\_refs.

Function\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 11 that
specifies one of the 11 functions you can use to consolidate data. If
function\_num is omitted, the SUM function, number 9, is used. The
functions and their corresponding numbers are listed in the following
table.

|                   |              |
| ----------------- | ------------ |
| **Function\_num** | **Function** |
| 1                 | AVERAGE      |
| 2                 | COUNT        |
| 3                 | COUNTA       |
| 4                 | MAX          |
| 5                 | MIN          |
| 6                 | PRODUCT      |
| 7                 | STDEV        |
| 8                 | STDEVP       |
| 9                 | SUM          |
| 10                | VAR          |
| 11                | VARP         |

The following arguments correspond to text boxes and check boxes in the
Consolidate dialog box. Arguments that correspond to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Top\_row&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Top Row check box.
The default is FALSE.

Left\_col&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Left Column check
box. The default is FALSE.

If top\_row and left\_col are both FALSE or omitted, the data is
consolidated by position.

Create\_links&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Create Links To
Source Data check box.

**Remarks**

  - > If you use the CONSOLIDATE function with no arguments and there is
    > a consolidation on the active worksheet, Microsoft Excel
    > reconsolidates, using the sources, function, and position
    > attributes used to create the existing consolidation.

  - > If you use the CONSOLIDATE function with no arguments and there is
    > no consolidation on the active worksheet, the function returns the
    > \#VALUE\! error value.


**Related Functions**

[CHANGE.LINK](CHANGE.LINK.md)&nbsp;&nbsp;&nbsp;Changes supporting workbook links

[LINKS](LINKS.md)&nbsp;&nbsp;&nbsp;Returns the names of all linked workbooks

[OPEN.LINKS](OPEN.LINKS.md)&nbsp;&nbsp;&nbsp;Opens specified supporting workbooks

[UPDATE.LINK](UPDATE.LINK.md)&nbsp;&nbsp;&nbsp;Updates a link to another workbook



Return to [README](README.md#C)

