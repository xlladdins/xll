CHANGE.LINK
===========

Equivalent to clicking the Change Source button in the Links dialog box,
which appears when you click the Links command on the Edit menu. Changes
a link from one supporting workbook to another.

**Syntax**

**CHANGE.LINK**(**old\_text, new\_text**, type\_of\_link)

**CHANGE.LINK**?(old\_text, new\_text, type\_of\_link)

Old\_text    is the path of the link from the active dependent workbook
you want to change.

New\_text    is the path of the link you want to change to.

Type\_of\_link    is the number 1 or 2 specifying what type of link you
want to change.

  -------------------- ------------------------
  **Type\_of\_link**   **Link document type**
  1 or omitted         Microsoft Excel link
  2                    DDE/OLE link
  -------------------- ------------------------

**Remarks**

The workbook whose links you want to change must be active when this
function is calculated.

**Related Functions**

GET.LINK.INFO   Returns information about a link

LINKS   Returns the name of all linked workbooks

OPEN.LINKS   Opens specified supporting workbooks

SET.UPDATE.STATUS   Controls the update status of a link

UPDATE.LINK   Updates a link to another workbook

Return to [top](#A)

CHART.ADD.DATA
