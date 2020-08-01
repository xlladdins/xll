# GET.DEF

Returns the name, as text, that is defined for a particular area, value,
or formula in a workbook. Use GET.DEF to get the name corresponding to a
definition. To get the definition of a name, use GET.NAME.

**Syntax**

**GET.DEF**(**def\_text**, document\_text, type\_num)

Def\_text&nbsp;&nbsp;&nbsp;&nbsp;can be anything you can define a name
to refer to, including a reference, a value, an object, or a formula.

  - > References must be given in R1C1 style, such as "R3C5".

  - > If def\_text is a value or formula, it is not necessary to include
    > the equal sign that is displayed in the Refers To box in the
    > Define Name dialog box, which appears when you choose the Name
    > command from the Define submenu on the Insert Menu.

  - > If there is more than one name for def\_text, GET.DEF returns the
    > first name. If no name matches def\_text, GET.DEF returns the
    > \#NAME? error value.

Document\_text&nbsp;&nbsp;&nbsp;&nbsp;specifies the sheet or macro sheet
that def\_text is on. If document\_text is omitted, it is assumed to be
the active macro sheet.

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 specifying
which types of names are returned.

|               |                   |
| ------------- | ----------------- |
| **Type\_num** | **Returns**       |
| 1 or omitted  | Normal names only |
| 2             | Hidden names only |
| 3             | All names         |

**Examples**

If the specified range in Sheet4 is named Sales, the following macro
formula returns "Sales":

GET.DEF("R2C2:R9C6", "Sheet4")

If the value 100 in Sheet4 is defined as Constant, the following macro
formula returns "Constant":

GET.DEF("100", "Sheet4")

If the specified formula in Sheet4 is named SumTotal, the following
macro formula returns "SumTotal":

GET.DEF("SUM(R1C1:R10C1)", "Sheet4")

If 3 is defined as the hidden name Counter on the active macro sheet,
the following macro formula returns "Counter":

GET.DEF("3", , 2)

**Related Functions**

[GET.CELL](GET.CELL.md)&nbsp;&nbsp;&nbsp;Returns information about the specified cell

[GET.NAME](GET.NAME.md)&nbsp;&nbsp;&nbsp;Returns the definition of a name

[GET.NOTE](GET.NOTE.md)&nbsp;&nbsp;&nbsp;Returns characters from a note

[NAMES](NAMES.md)&nbsp;&nbsp;&nbsp;Returns the names defined on a workbook



Return to [README](README.md#G)

# GET.DEF

Returns the name, as text, that is defined for a particular area, value,
or formula in a workbook. Use GET.DEF to get the name corresponding to a
definition. To get the definition of a name, use GET.NAME.

**Syntax**

**GET.DEF**(**def\_text**, document\_text, type\_num)

Def\_text&nbsp;&nbsp;&nbsp;&nbsp;can be anything you can define a name
to refer to, including a reference, a value, an object, or a formula.

  - > References must be given in R1C1 style, such as "R3C5".

  - > If def\_text is a value or formula, it is not necessary to include
    > the equal sign that is displayed in the Refers To box in the
    > Define Name dialog box, which appears when you choose the Name
    > command from the Define submenu on the Insert Menu.

  - > If there is more than one name for def\_text, GET.DEF returns the
    > first name. If no name matches def\_text, GET.DEF returns the
    > \#NAME? error value.

Document\_text&nbsp;&nbsp;&nbsp;&nbsp;specifies the sheet or macro sheet
that def\_text is on. If document\_text is omitted, it is assumed to be
the active macro sheet.

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 specifying
which types of names are returned.

|               |                   |
| ------------- | ----------------- |
| **Type\_num** | **Returns**       |
| 1 or omitted  | Normal names only |
| 2             | Hidden names only |
| 3             | All names         |

**Examples**

If the specified range in Sheet4 is named Sales, the following macro
formula returns "Sales":

GET.DEF("R2C2:R9C6", "Sheet4")

If the value 100 in Sheet4 is defined as Constant, the following macro
formula returns "Constant":

GET.DEF("100", "Sheet4")

If the specified formula in Sheet4 is named SumTotal, the following
macro formula returns "SumTotal":

GET.DEF("SUM(R1C1:R10C1)", "Sheet4")

If 3 is defined as the hidden name Counter on the active macro sheet,
the following macro formula returns "Counter":

GET.DEF("3", , 2)

**Related Functions**

[GET.CELL](GET.CELL.md)&nbsp;&nbsp;&nbsp;Returns information about the specified cell

[GET.NAME](GET.NAME.md)&nbsp;&nbsp;&nbsp;Returns the definition of a name

[GET.NOTE](GET.NOTE.md)&nbsp;&nbsp;&nbsp;Returns characters from a note

[NAMES](NAMES.md)&nbsp;&nbsp;&nbsp;Returns the names defined on a workbook



Return to [README](README.md#G)

