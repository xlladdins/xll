WORKGROUP

Equivalent to clicking the Group Edit command on the Options menu in
Microsoft Excel version 4.0. Creates a group. This function is provided
for compatibility only. In Microsoft Excel version 5.0 and later
versions, you can create a group by using the WORKBOOK.SELECT function.

**Syntax**

**WORKGROUP**(name\_array)

**WORKGROUP**?(name\_array)

Name\_array    is the list of workbooks or sheets in workbooks that you
want grouped.

  - > If name\_array is omitted, the most recently created group is
    > recreated.

  - > If no group has been created during the current Microsoft Excel
    > session, all open, unhidden worksheets are created as a group.

  - > If you specify just the name of a workbook, WORKGROUP adds the
    > first sheet of the workbook to the group.

>  

**Remarks**

WORKGROUP returns the \#VALUE\! error value and interrupts the macro if
it can't find any of the sheets in name\_array or if any of the sheets
is a chart or module.

**Related Functions**

[FILL.GROUP](FILL.GROUP.md)   Fills the contents of the active worksheet's selection to
the same area on all other worksheets in the group

[WORKBOOK.SELECT](WORKBOOK.SELECT.md)   Selects one or more sheets in a workbook



Return to [README.md](README.md)

WORKGROUP

Equivalent to clicking the Group Edit command on the Options menu in
Microsoft Excel version 4.0. Creates a group. This function is provided
for compatibility only. In Microsoft Excel version 5.0 and later
versions, you can create a group by using the WORKBOOK.SELECT function.

**Syntax**

**WORKGROUP**(name\_array)

**WORKGROUP**?(name\_array)

Name\_array    is the list of workbooks or sheets in workbooks that you
want grouped.

  - > If name\_array is omitted, the most recently created group is
    > recreated.

  - > If no group has been created during the current Microsoft Excel
    > session, all open, unhidden worksheets are created as a group.

  - > If you specify just the name of a workbook, WORKGROUP adds the
    > first sheet of the workbook to the group.

>  

**Remarks**

WORKGROUP returns the \#VALUE\! error value and interrupts the macro if
it can't find any of the sheets in name\_array or if any of the sheets
is a chart or module.

**Related Functions**

[FILL.GROUP](FILL.GROUP.md)   Fills the contents of the active worksheet's selection to
the same area on all other worksheets in the group

[WORKBOOK.SELECT](WORKBOOK.SELECT.md)   Selects one or more sheets in a workbook



Return to [README.md](README.md)

