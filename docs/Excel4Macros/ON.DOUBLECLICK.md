# ON.DOUBLECLICK

Runs a macro when you double-click any cell or object on the specified
sheet or macro sheet or double-click any item on the specified chart.

**Syntax**

**ON.DOUBLECLICK**(sheet\_text, macro\_text)

Sheet\_text&nbsp;&nbsp;&nbsp;&nbsp;is a text value specifying the name
of a sheet in a workbook. If sheet\_text is omitted, the macro is run
whenever you double-click any sheet not specified by a previous
ON.DOUBLECLICK formula. Sheet\_text must be in the form
"\[book1\]sheet1".

Macro\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of, or an R1C1-style
reference to, a macro you want to run when you double-click the sheet
specified by sheet\_text. The name or reference must be in text form. If
macro\_text is omitted, double-clicking reverts to its normal behavior,
and any macros assigned by previous ON.DOUBLECLICK functions are turned
off.

**Remarks**

  - > ON.DOUBLECLICK overrides Microsoft Excel's normal double-click
    > behavior, such as displaying cell notes, displaying the Patterns
    > dialog box, or allowing editing cell text directly in cells.

  - > To determine what cell, object, or chart item has been
    > double-clicked, use a CALLER function in the macro specified by
    > macro\_text.

  - > ON.DOUBLECLICK does not affect objects to which ASSIGN.TO.OBJECT
    > macros have already been assigned. Use ON.DOUBLECLICK (TRUE) to
    > make Microsoft Excel carry out the action that would normally
    > occur if you double-click on the current selection.


**Related Functions**

[ASSIGN.TO.OBJECT](ASSIGN.TO.OBJECT.md)&nbsp;&nbsp;&nbsp;Assigns a macro to an object

[ON.WINDOW](ON.WINDOW.md)&nbsp;&nbsp;&nbsp;Runs a macro when you switch to a window



Return to [README](README.md)

