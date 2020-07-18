GET.NAME

Returns the definition of a name as it appears in the Refers To box of
the Define Name dialog box, which appears when you choose the Define
command from the Name submenu on the Insert menu. If the definition
contains references, they are given as R1C1-style references. Use
GET.NAME to check the value defined by a name. To get the name
corresponding to a definition, use GET.DEF.

**Syntax**

**GET.NAME**(**name\_text**, info\_type)

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;can be a name defined on the macro
sheet; an external reference to a name defined on the active workbook,
for example, "\!Sales"; or an external reference to a name defined on a
particular open workbook, for example, "\[Book1\]SHEET1\!Sales".
Name\_text can also be a hidden name.

Info\_type&nbsp;&nbsp;&nbsp;&nbsp; specifies the type of information to
return about the name. If 1 or omitted, the definition is returned. If
2, returns TRUE if the name is defined for just the sheet, FALSE if the
name is defined for the entire workbook.

**Remarks**

If the Contents check box has been selected in the Protect Sheet dialog
box to protect the workbook containing the name, GET.NAME returns the
\#N/A error value. To see the Protect Sheet dialog box, choose the
Protect Sheet command on the Protection submenu from the Tools menu.

**Examples**

If the name Sales on a macro sheet is defined as the number 523, then:

GET.NAME("Sales") equals "=523"

If the name Profit on the active sheet is defined as the formula
=Sales-Costs, then:

GET.NAME("\!Profit") equals "=Sales-Costs"

If the name Database on the active sheet is defined as the range
A1:F500, then:

GET.NAME("\!Database") equals "=R1C1:R500C6"

**Related Functions**

[DEFINE.NAME](DEFINE.NAME.md)&nbsp;&nbsp;&nbsp;Defines a name on the active or macro sheet

[GET.CELL](GET.CELL.md)&nbsp;&nbsp;&nbsp;Returns information about the specified cell

[GET.DEF](GET.DEF.md)&nbsp;&nbsp;&nbsp;Returns a name matching a definition

[NAMES](NAMES.md)&nbsp;&nbsp;&nbsp;Returns the names defined in a workbook

[SET.NAME](SET.NAME.md)&nbsp;&nbsp;&nbsp;Defines a name as a value



Return to [README](README.md)

