GET.FORMULA

Returns the contents of a cell as they would appear in the formula bar.
The contents are given as text, for example, "=2\*PI()/360". If the
formula contains references, they are returned as R1C1-style references,
such as "=RC\[1\]\*(1+R1C1)". Use GET.FORMULA to get a formula from a
cell in order to edit its arguments. Use GET.CELL(6) to get a formula in
either A1 or R1C1 format, depending on the workspace setting.

**Syntax**

**GET.FORMULA**(**reference**)

Reference**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a cell or range of cells on a sheet
or macro sheet.

  - > If a range of cells is selected, GET.FORMULA returns the contents
    > of the upper-left cell in reference.

  - > Reference can be an external reference.

  - > Reference can be the object identifier of a picture created by the
    > camera tool.

  - > Reference can also be a reference to a chart series in the form
    > "Sn" where n is the number of the series. When a chart series is
    > specified, GET.FORMULA returns the series formula using R1C1-style
    > references.

**Tip****&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;If you want to get the formula in the active
cell, use the ACTIVE.CELL function as the reference argument.

**Examples**

If cell A3 on the active sheet contains the number 523, then:

GET.FORMULA(\!$A$3) equals "523"

If cell C2 on the active sheet contains the formula =B2\*(1+$A$1), then:

GET.FORMULA(\!$C$2) equals "=RC\[-1\]\*(1+R1C1)"

The following macro formula returns the contents of the active cell on
the active sheet:

GET.FORMULA(ACTIVE.CELL())

**Related Functions**

[GET.CELL](GET.CELL.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns information about the specified cell

[GET.DEF](GET.DEF.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns a name matching a definition

[GET.NAME](GET.NAME.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns the definition of a name

[GET.NOTE](GET.NOTE.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns characters from a comment



Return to [README](README.md)

