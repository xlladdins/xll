COLUMN.WIDTH

Equivalent to choosing the Width command on the Column submenu of the
Format menu. Changes the width of the columns in the specified
reference.

**Syntax**

**COLUMN.WIDTH**(width\_num, reference, standard, type\_num,
standard\_num)

**COLUMN.WIDTH**?(width\_num, reference, standard, type\_num,
standard\_num)

Width\_num    specifies how wide you want the columns to be in units of
one character of the font corresponding to the Normal cell style.
Width\_num is ignored if standard is TRUE or if type\_num is provided.

Reference    specifies the columns for which you want to change the
width.

  - > If reference is specified, it must be either an external reference
    > to the active worksheet, such as \!$A:$C or \!Database, or an
    > R1C1-style reference in the form of text, such as "C1:C3",
    > "C\[-4\]:C\[-2\]", or "Database".

  - > If reference is a relative R1C1-style reference in the form of
    > text, it is assumed to be relative to the active cell.

  - > If reference is omitted, it is assumed to be the current
    > selection.

>  

Standard\_num    is a logical value corresponding to the Standard Width
command from the Column submenu on the Format menu.

  - > If standard is TRUE, Microsoft Excel sets the column width to the
    > currently defined standard (default) width and ignores width\_num.

  - > If standard is FALSE or omitted, Microsoft Excel sets the width
    > according to width\_num or type\_num.

>  

Type\_num    is a number from 1 to 3 corresponding to the Hide, Unhide,
or AutoFit Selection commands, respectively, on the Column submenu of
the Format menu.

|               |                                                                                                                                                     |
| ------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Action taken**                                                                                                                                    |
| 1             | Hides the column selection by setting the column width to 0                                                                                         |
| 2             | Unhides the column selection by setting the column width to the value set before the selection was hidden                                           |
| 3             | Sets the column selection to a best-fit width, which varies from column to column depending on the length of the longest data string in each column |

Standard\_num    specifies how wide the standard width is, and is
measured in points. If standard\_num is omitted, the standard width
setting remains unchanged.

**Remarks**

  - > Changing the value of standard\_num changes the width of all
    > columns except those that have been set to a custom value.

  - > If any of the argument settings conflict, such as when standard is
    > TRUE and type\_num is 3, Microsoft Excel uses the type\_num
    > argument and ignores any arguments that conflict with type\_num.

  - > If you are recording a macro while using a mouse and you change
    > column widths by dragging the column border, Microsoft Excel
    > records the references of the columns using R1C1-style references
    > in the form of text.

>  

**Related Function**

[ROW.HEIGHT](ROW.HEIGHT.md)   Changes the heights of rows



Return to [README](README.md)

