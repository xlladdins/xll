REPORT.GET
==========

Returns information about reports defined for the active workbook. Use
REPORT.GET to return information you can use in other macro commands
that manipulate reports.

If this function is not available, you must install the Report Manager
add-in.

**Syntax**

**REPORT.GET**(**type\_num**, report\_name)

Type\_num    is a number from 1 to 3 specifying the type of information
to return, as shown in the following table.

+-----------------+---------------------------------------------------+
| > **Type\_num** | > **Returns**                                     |
+-----------------+---------------------------------------------------+
| > 1             | > An array of reports from all sheets in the      |
|                 | > active workbook or the \#N/A error value if     |
|                 | > none are specified                              |
+-----------------+---------------------------------------------------+
| > 2             | > An array of views, scenarios, and sheet names   |
|                 | > for the specified report in the active          |
|                 | > workbook. REPORT.GET returns the \#N/A error    |
|                 | > value if the scenario check box is not          |
|                 | > selected. Returns the \#VALUE! error value if   |
|                 | > name is invalid or the workbook is protected.   |
+-----------------+---------------------------------------------------+
| > 3             | > If continuous page numbers are used, returns    |
|                 | > TRUE. If page numbers start at 1 for each       |
|                 | > section, returns FALSE. Returns the \#VALUE!    |
|                 | > error value if report\_name is invalid or the   |
|                 | > workbook is protected.                          |
+-----------------+---------------------------------------------------+

Report\_name    specifies the name of a report in the active workbook.

**Remarks**

Report\_name is required if type\_num is 2 or 3.

**Example**

The following macro formula returns an array of reports from the active
workbook.

REPORT.GET(1)

**Related Functions**

REPORT.DEFINE   Creates a report

REPORT.DELETE   Removes a report from the active workbook

REPORT.PRINT   Prints a report

Return to [top](#Q)

REPORT.PRINT
