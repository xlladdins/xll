DEFINE.NAME
===========

Equivalent to clicking the Define command on the Name submenu of the
Insert menu. Defines a name on the active sheet or macro sheet. Use
DEFINE.NAME instead of SET.NAME when you want to define a name on the
active sheet.

**Syntax**

**DEFINE.NAME**(**name\_text**, refers\_to, macro\_type, shortcut\_text,
hidden, category, local)

**DEFINE.NAME**?(name\_text, refers\_to, macro\_type, shortcut\_text,
hidden, category, local)

Name\_text    is the text you want to use as the name. Names cannot
include spaces, and cannot look like cell references.

Refers\_to    describes what name\_text should refer to, and can be any
of the following values.

  ------------------------------------------------------------------------------------------------------------------------------------------- -------------------------------------------
  **If refers\_to is**                                                                                                                        **Then name\_text is**

  A number, text, or logical value                                                                                                            Defined to refer to that value

  An external reference, such as !\$A\$1 or SALES!\$A\$1:\$C\$3                                                                               Defined to refer to those cells

  A formula in the form of text, such as \"=2\*PI()/360\" (if the formula contains references, they must be R1C1-style references, such as\   Defined to refer to that formula
  \"=R2C2\*(1+RC\[-1\])\")                                                                                                                    

  Omitted                                                                                                                                     Defined to refer to the current selection
  ------------------------------------------------------------------------------------------------------------------------------------------- -------------------------------------------

The next two arguments, macro\_type and shortcut\_text, apply only if
the sheet in the active window is a macro sheet.

Macro\_type    is a number from 1 to 3 that indicates the type of macro.

  ----------------- ------------------------------------------------------
  **Macro\_type**   **Type of macro**
  1                 Custom function (also known as a function macro)
  2                 Command macro.
  3 or omitted      None (that is, name\_text does not refer to a macro)
  ----------------- ------------------------------------------------------

Shortcut\_text    is a text value that specifies the macro shortcut key.
Shortcut\_text must be a single letter, such as \"z\" or \"Z\".

Hidden    is a logical value specifying whether to define the name as a
hidden name. If hidden is TRUE, Microsoft Excel defines the name as a
hidden name; if FALSE or omitted, Microsoft Excel defines the name
normally.

Category    is a number or text identifying the category of a custom
function and corresponds to categories in the Function Category list
box.

-   Categories are numbered starting with 1, the first category in the
    > list.

-   If category is text but is not one of the existing function types,
    > Microsoft Excel creates a new category and assigns your custom
    > function to it.

Local    is a logical value which, if TRUE, defines the name on just the
current sheet or macro sheet. If FALSE or omitted, defines the name for
all sheets in the workbook.

**Remarks**

-   You can use hidden names to define values that you want to prevent
    > the user from seeing or changing; they do not appear in the Define
    > Name, Paste Name, or Goto dialog boxes. Hidden names can only be
    > created with the DEFINE.NAME macro function.

-   If you are recording a macro and you define a name to refer to a
    > formula, Microsoft Excel converts A1-style references to
    > R1C1-style references. For example, if the active cell is C2, and
    > you define the name Previous to refer to =B2, Microsoft Excel
    > records that command as DEFINE.NAME(\"Previous\",\"=RC\[-1\]\").

-   In DEFINE.NAME?, the dialog-box form of the function, if refers\_to
    > is not specified, the current selection is proposed in the Refers
    > To box. Also, if a name is not specified, text in the active cell
    > is proposed as the name.

>  

**Related Functions**

DELETE.NAME   Deletes a name

GET.DEF   Returns a name matching a definition

GET.NAME   Returns the definition of a name

NAMES   Returns the names defined in a workbook

SET.NAME   Defines a name as a value

Return to [top](#A)

DEFINE.STYLE
