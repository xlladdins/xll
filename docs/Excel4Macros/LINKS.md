LINKS

Returns, as a horizontal array of text values, the names of all
workbooks referred to by external references in the workbook specified.
Use LINKS with OPEN.LINKS to open supporting workbooks.

**Syntax**

**LINKS**(document\_text, type\_num)

Document\_text    is the name of a workbook, including its path. If
document\_text is omitted, LINKS operates on the active workbook. If the
workbook specified by document\_text is not open, LINKS returns the
\#N/A error value.

Type\_num    is a number from 1 to 6 specifying the type of linked
workbooks to return.

|               |                                                |
| ------------- | ---------------------------------------------- |
| **Type\_num** | **Returns**                                    |
| 1 or omitted  | Microsoft Excel link                           |
| 2             | DDE/OLE link (Microsoft Excel for Windows)     |
| 3             | Reserved                                       |
| 4             | Not applicable                                 |
| 5             | Publisher (Microsoft Excel for the Macintosh)  |
| 6             | Subscriber (Microsoft Excel for the Macintosh) |

**Remarks**

  - > If the active workbook contains no external references, LINKS
    > returns the \#N/A error value.

  - > With the INDEX function, you can select individual workbook names
    > from the array for use in other functions that take workbook names
    > as arguments.

  - > The names of the workbook are always returned in alphabetic order.
    > If supporting workbooks are open, LINKS returns the names of the
    > workbooks; if supporting workbooks are closed, LINKS includes the
    > full path of each workbook.

  - > If type\_num is 5 or 6, LINKS returns a two-row array in which the
    > first row contains the edition name and the second row contains
    > the reference.

>  

**Examples**

If a chart named Chart1 is open and contains links to workbook Data1 and
Data2, and the LINKS function shown below is entered as an array into a
two-cell horizontal range:

LINKS("Chart1") equals "Data1" in the first cell of the range and
"Data2" in the second cell.

In Microsoft Excel for Windows, if the chart named VARIANCE.XLS is open
and contains data series that refer to workbook named BUDGET.XLS and
ACTUAL.XLS, then:

OPEN.LINKS(LINKS("VARIANCE.XLS")) opens BUDGET.XLS and ACTUAL.XLS.

In Microsoft Excel for the Macintosh, if the workbook named SALES 1991
is open and contains references to the workbook WEST SALES, SOUTH SALES,
and EAST SALES, then:

OPEN.LINKS(LINKS("SALES 1991")) opens WEST SALES, SOUTH SALES, and EAST
SALES.

**Related Functions**

[CHANGE.LINK](CHANGE.LINK.md)   Changes supporting workbook links

[GET.LINK.INFO](GET.LINK.INFO.md)   Returns information about a link

[OPEN.LINKS](OPEN.LINKS.md)   Opens specified supporting workbook

[UPDATE.LINK](UPDATE.LINK.md)   Updates a link to another workbook



Return to [README.md](README.md)

LINKS

Returns, as a horizontal array of text values, the names of all
workbooks referred to by external references in the workbook specified.
Use LINKS with OPEN.LINKS to open supporting workbooks.

**Syntax**

**LINKS**(document\_text, type\_num)

Document\_text    is the name of a workbook, including its path. If
document\_text is omitted, LINKS operates on the active workbook. If the
workbook specified by document\_text is not open, LINKS returns the
\#N/A error value.

Type\_num    is a number from 1 to 6 specifying the type of linked
workbooks to return.

|               |                                                |
| ------------- | ---------------------------------------------- |
| **Type\_num** | **Returns**                                    |
| 1 or omitted  | Microsoft Excel link                           |
| 2             | DDE/OLE link (Microsoft Excel for Windows)     |
| 3             | Reserved                                       |
| 4             | Not applicable                                 |
| 5             | Publisher (Microsoft Excel for the Macintosh)  |
| 6             | Subscriber (Microsoft Excel for the Macintosh) |

**Remarks**

  - > If the active workbook contains no external references, LINKS
    > returns the \#N/A error value.

  - > With the INDEX function, you can select individual workbook names
    > from the array for use in other functions that take workbook names
    > as arguments.

  - > The names of the workbook are always returned in alphabetic order.
    > If supporting workbooks are open, LINKS returns the names of the
    > workbooks; if supporting workbooks are closed, LINKS includes the
    > full path of each workbook.

  - > If type\_num is 5 or 6, LINKS returns a two-row array in which the
    > first row contains the edition name and the second row contains
    > the reference.

>  

**Examples**

If a chart named Chart1 is open and contains links to workbook Data1 and
Data2, and the LINKS function shown below is entered as an array into a
two-cell horizontal range:

LINKS("Chart1") equals "Data1" in the first cell of the range and
"Data2" in the second cell.

In Microsoft Excel for Windows, if the chart named VARIANCE.XLS is open
and contains data series that refer to workbook named BUDGET.XLS and
ACTUAL.XLS, then:

OPEN.LINKS(LINKS("VARIANCE.XLS")) opens BUDGET.XLS and ACTUAL.XLS.

In Microsoft Excel for the Macintosh, if the workbook named SALES 1991
is open and contains references to the workbook WEST SALES, SOUTH SALES,
and EAST SALES, then:

OPEN.LINKS(LINKS("SALES 1991")) opens WEST SALES, SOUTH SALES, and EAST
SALES.

**Related Functions**

[CHANGE.LINK](CHANGE.LINK.md)   Changes supporting workbook links

[GET.LINK.INFO](GET.LINK.INFO.md)   Returns information about a link

[OPEN.LINKS](OPEN.LINKS.md)   Opens specified supporting workbook

[UPDATE.LINK](UPDATE.LINK.md)   Updates a link to another workbook



Return to [README.md](README.md)

