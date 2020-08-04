# REPORT.PRINT

Equivalent to clicking the Print button in the Report Manager dialog
box. Prints a report.

If this function is not available, you must install the Report Manager
add-in.

**Syntax**

**REPORT.PRINT**(**report\_name**, copies\_num,
show\_print\_dlg\_logical)

**REPORT.PRINT**?(report\_name, copies\_num)

Report\_name&nbsp;&nbsp;&nbsp;&nbsp;specifies the name of a report in
the active workbook.

Copies\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number of copies you want to
print. If omitted, the default is 1.

Show\_print\_dlg\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value
that, if TRUE, displays a dialog box asking how many copies to print,
or, if FALSE or omitted, prints the report immediately using existing
print settings.

**Remarks**

REPORT.PRINT returns the \#VALUE\! error value if report\_name is
invalid or if the workbook is protected.

**Related Functions**

[REPORT.DEFINE](REPORT.DEFINE.md)&nbsp;&nbsp;&nbsp;Creates a report

[REPORT.DELETE](REPORT.DELETE.md)&nbsp;&nbsp;&nbsp;Removes a report from the active workbook



Return to [README](README.md#R)

