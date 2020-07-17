CALLER
======

Returns information about the cell, range of cells, command on a menu,
tool on a toolbar, or object that called the macro that is currently
running. Use CALLER in a subroutine or custom function whose behavior
depends on the location, size, name, or other attribute of the caller.

**Syntax**

**CALLER**( )

**Remarks**

-   If the custom function is entered in a single cell, CALLER returns
    > the reference of that cell.

-   If the custom function was part of an array formula entered in a
    > range of cells, CALLER returns the reference of the range.

-   If CALLER appears in a macro called by an Auto\_Open, Auto\_Close,
    > Auto\_Activate, or Auto\_Deactivate macro, it returns the name of
    > the calling sheet.

-   If CALLER appears in a macro called by a command on a menu, it
    > returns a horizontal array of three elements including the
    > command\'s position number, the menu number, and the menu bar
    > number.

-   If CALLER appears in a macro called by an assigned-to-object macro,
    > it returns the object identifier.

-   If CALLER appears in a macro called by a tool on a toolbar, it
    > returns a horizontal array containing the position number and the
    > toolbar name.

-   If CALLER appears in a macro called by an ON.DOUBLECLICK or ON.ENTRY
    > function, CALLER returns the name of the chart object identifier
    > or cell reference, if applicable, to which the ON.DOUBLECLICK or
    > ON.ENTRY macro applies.

-   If CALLER appears in a macro that was run manually, or for any
    > reason not described above, it returns the \#REF! error value.

>  

**Examples**

If the custom function MACROS!VALUEONE is entered in cell B3 on a sheet
named SALES, the nested CALLER function returns the following values.

  ----------------------- --------------
  **Nested function**     **Returns**
  COLUMN(CALLER())        2
  COLUMNS(CALLER())       1
  GET.CELL(1, CALLER())   SALES!\$B\$3
  ROW(CALLER())           3
  ROWS(CALLER())          1
  ----------------------- --------------

If the same custom function was entered into an array in cells B2:C3,
the following values would be returned.

  --------------------- -------------
  **Nested function**   **Returns**
  COLUMN(CALLER())      2
  COLUMNS(CALLER())     2
  ROW(CALLER())         2
  ROWS(CALLER())        2
  --------------------- -------------

**Related Functions**

GET.BAR   Returns the name or position number of menu bars, menus, and
commands

GET.CELL   Returns information about the specified cell

Return to [top](#A)

CANCEL.COPY
