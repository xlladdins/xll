LIST.NAMES

Equivalent to clicking the Paste command on the Name submenu of the
Insert menu and selecting the Paste List option button. Lists all names
(except hidden names) defined in your workbook. LIST.NAMES also lists
the cells to which the names refer; whether a macro corresponding to a
particular name is a command macro or a custom function; the shortcut
key for each command macro; and the category of the custom functions.

**Syntax**

**LIST.NAMES**( )

**Remarks**

  - > If the current selection is a single cell or five or more columns
    > wide, LIST.NAMES pastes all five types of information about
    > worksheet names into five columns. The first column contains cell
    > names. The second column contains the corresponding cell
    > references. The third column contains the number 1 if the name
    > refers to a custom function, the number 2 if it refers to a
    > command macro, or 0 if it refers to anything else. The fourth
    > column lists the shortcut keys for command macros. The fifth
    > column contains a category name for custom functions or the number
    > of the built-in category.

  - > If the selection includes fewer than five columns, LIST.NAMES
    > omits the information that would have been pasted into the other
    > columns.

  - > When you use LIST.NAMES, Microsoft Excel completely replaces the
    > contents of the cells it pastes into.

>  

**Related Functions**

GET.DEF   Returns a name matching a definition

GET.NAME   Returns the definition of a name

NAMES   Returns the names defined in a workbook


