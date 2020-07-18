ERROR

Specifies what action to take if an error is encountered while a macro
is running. Use ERROR to control whether Microsoft Excel error messages
are displayed, or to run your own macro when an error is encountered.

**Syntax**

**ERROR**(**enable\_logical**, macro\_ref)

Enable\_logical**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value or number that
selects or clears error-checking.

  - > If enable\_logical is FALSE or 0, all error-checking is cleared.
    > If error-checking is cleared and an error is encountered while a
    > macro is running, Microsoft Excel ignores it and continues.
    > Error-checking is selected again by an ERROR(TRUE) statement, or
    > when the macro stops running.

  - > If enable\_logical is TRUE or 1, you can either select normal
    > error-checking (by omitting the other argument) or specify a macro
    > to run when an error is encountered by using the macro\_ref
    > argument. When normal error-checking is active, the Macro Error
    > dialog box is displayed when an error is encountered. You can halt
    > the macro, start single-stepping through the macro, continue
    > running the macro normally, or go to the macro cell where the
    > error occurred.

  - > If enable\_logical is 2 and macro\_ref is omitted, error-checking
    > is normal except that if the user clicks the Cancel button in an
    > alert message, ERROR returns FALSE and the macro is not
    > interrupted.

  - > If enable\_logical is 2 and macro\_ref is given, the macro goes to
    > that macro\_ref when an error is encountered. If the user clicks
    > the Cancel button in an alert message, FALSE is returned and the
    > macro is not interrupted.


Macro\_ref**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies a macro to run if
enable\_logical is TRUE, 1, or 2 and an error is encountered. It can be
either the name of the macro or a cell reference. If enable\_logical is
FALSE or 0, macro\_ref is ignored.

**Important**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;**Both ERROR(FALSE) and ERROR(TRUE,
macro\_ref ) keep Microsoft Excel from displaying any messages at all,
including the message asking whether to save changes when you close an
unsaved workbook. If you want alert messages but not error messages to
be displayed, use ERROR(2, macro\_ref ).

**Remarks**

You can use GET.WORKSPACE to determine whether error-checking is on or
off.

**Examples**

ERROR(FALSE) clears error-checking.

ERROR(TRUE, Recover) selects error-checking and runs the macro named
Recover when an error is encountered.

The following macro runs the macro ForceMenus if an error occurs in the
current macro:

\=ERROR(TRUE, ForceMenus)

**Related Functions**

[CANCEL.KEY](CANCEL.KEY.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Disables macro interruption

[LAST.ERROR](LAST.ERROR.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns the reference of the cell where the
last error occurred

[ON.KEY](ON.KEY.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Runs a macro when a specified key is pressed



Return to [README](README.md)

