RETURN

Ends the currently running macro. If the currently running macro is a
subroutine macro that was called by another macro, control is returned
to the calling macro. If the currently running macro is a custom
function, control is returned to the formula that called the custom
function. If the currently running macro is a command macro started by
the user with the Run button in the Macro dialog box or a shortcut key
or by clicking an object, control is returned to the user.

**Syntax**

**RETURN**(value)

Value**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies what to return.

  - > If the macro is a custom function or a subroutine, value specifies
    > what value to return. However, not all subroutines return values;
    > the last line in macros that do not return values is =RETURN().

  - > If the macro is a command macro run by the user, value should be
    > omitted.


**Remarks**

RETURN signals the end of a macro. Every macro must end with a RETURN or
HALT function, but not every macro returns values.

**Example**

The following function returns the sum of the range B1:B10:

RETURN(SUM(B1:B10))

**Related Functions**

[BREAK](BREAK.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Interrupts a FOR-NEXT, FOR.CELL-NEXT, or
[WHILE](WHILE.md)-NEXT loop

[HALT](HALT.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Stops all macros from running

[RESULT](RESULT.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Specifies the data type a custom function
returns



Return to [README](README.md)

