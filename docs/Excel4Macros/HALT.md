HALT

Stops all macros from running. Use HALT instead of RETURN to prevent a
macro from returning to the macro that called it.

**Syntax**

**HALT**(cancel\_close)

Cancel\_close&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether a macro sheet, when encountering the HALT function in an
Auto\_Close macro, is closed.

  - > If cancel\_close is TRUE, Microsoft Excel halts the macro and
    > prevents the workbook from being closed.

  - > If cancel\_close is FALSE or omitted, Microsoft Excel halts the
    > macro and allows the workbook to be closed.

  - > If cancel\_close is specified in a macro that is not an
    > Auto\_Close macro, it is ignored and the HALT function simply
    > stops the current macro.


**Remarks**

You can prevent an Auto\_Close or Auto\_Open macro from running by
holding down the SHIFT key while opening or closing the workbook.

**Examples**

If A1 contains the \#N/A error value, then when the following macro
formula is calculated, the macro halts:

IF(ISERROR(A1), HALT(), GOTO(D4))

The following macro formula at the end of an Auto\_Close macro ends the
macro and prevents the workbook from being closed:

HALT(TRUE)

**Related Functions**

[BREAK](BREAK.md)&nbsp;&nbsp;&nbsp;Interrupts a FOR-NEXT, FOR.CELL-NEXT, or
[WHILE](WHILE.md)-NEXT loop

[RETURN](RETURN.md)&nbsp;&nbsp;&nbsp;Ends the currently running macro



Return to [README](README.md)

