WORKBOOK.SELECT
===============

Equivalent to selecting a sheet or group of sheets in the active
workbook. If you select a group of sheets, subsequent commands effect
all the sheets in the group.

**Syntax**

**WORKBOOK.SELECT**(name\_array, active\_name, replace)

Name\_array    is a horizontal array of text names of sheets you want to
select. If name\_array is omitted, no sheets are selected.

Active\_name    is the name of a single sheet in the workbook that you
want to be the active sheet. If active\_name is omitted, the first sheet
in name\_array is made the active sheet.

Replace    specifies whether the currently selected sheets or macro
sheets are to be replaced by name\_array. If TRUE or omitted, then the
current sheet selection is replaced by name\_array. If FALSE, then
name\_array will be appended to the current sheet.

**Related Functions**

GET.WORKBOOK   Returns information about a workbook

SELECT   Selects a cell, worksheet object, or chart item

Return to [top](#T)

WORKBOOK.TAB.SPLIT
