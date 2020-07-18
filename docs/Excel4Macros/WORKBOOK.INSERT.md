# WORKBOOK.INSERT

Equivalent to clicking the Worksheet, Chart, or Macro commands on the
Insert menu. Inserts one or more new sheets into the current workbook.

**Syntax**

**WORKBOOK.INSERT**(**type\_num**)

**WORKBOOK.INSERT**?(**type\_num**)

**Type\_num**&nbsp;&nbsp;&nbsp;&nbsp;specifies the type of sheet to
insert.

|               |                                               |
| ------------- | --------------------------------------------- |
| **Type\_num** | **Type of sheet**                             |
| 1             | Worksheet                                     |
| 2             | Chart                                         |
| 3             | Microsoft Excel 4.0 Macro Sheet               |
| 4             | Microsoft Excel 4.0 International Macro Sheet |
| 5             | (Reserved)                                    |
| 6             | Microsoft Excel Visual Basic Module           |
| 7             | Dialog                                        |
| Quoted text   | Template                                      |

If omitted, the type of the active sheet is used.

If the current selection contains one sheet, then only one sheet is
inserted. If the selection contains more than one sheet and the active
sheet is a worksheet, then an equal number of sheets is inserted to the
left of the selected group of sheets.

**Remarks**

  - > The new sheets are always inserted to the left of the current
    > selection.

  - > If the workbook structure is protected, you cannot insert new
    > sheets.



Return to [README](README.md)

