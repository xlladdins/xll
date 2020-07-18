SEND.KEYS

Sends keystrokes to the active application just as if they were typed at
the keyboard. Use SEND.KEYS to send keystrokes that perform actions and
execute commands to applications you are running with Microsoft Excel's
other dynamic data exchange (DDE) functions.

**Syntax**

**SEND.KEYS**(**key\_text**, wait\_logical)

**Note&nbsp;&nbsp;&nbsp;**This function is available only in Microsoft
Excel for Windows.

Key\_text&nbsp;&nbsp;&nbsp;&nbsp;is the key or key combination you want
to send to another application. The format for key\_text is described in
the ON.KEY function.

Wait\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that determines
whether the macro continues before the actions caused by key\_text are
carried out.

  - > If wait\_logical is TRUE, Microsoft Excel waits for the keys to be
    > processed before returning control to the macro.

  - > If wait\_logical is FALSE or omitted, the macro continues running
    > without waiting for the keys to be processed.

> &nbsp;

**Remarks**

If Microsoft Excel is the active application, wait\_logical is assumed
to be FALSE, even if you enter wait\_logical as TRUE. This is because if
wait\_logical is TRUE, Microsoft Excel waits for the keys to be
processed in the other application before returning control to the
macro. Microsoft Excel doesn't process keys while a macro is running.

**Example**

The following macro uses the Calculator application in Microsoft Excel
for Windows to multiply some numbers, and then cuts the result and
pastes it into Microsoft Excel.

\=EXEC("CALC.EXE", 1)

\=SEND.KEYS("10\*30", TRUE)

\=SEND.KEYS("\~", TRUE)

\=SEND.KEYS("%ec", TRUE)

\=APP.ACTIVATE(, FALSE)

\=SELECT(\!B1)

\=PASTE()

\=RETURN()

**Related Functions**

[APP.ACTIVATE](APP.ACTIVATE.md)&nbsp;&nbsp;&nbsp;Switches to an application

[EXECUTE](EXECUTE.md)&nbsp;&nbsp;&nbsp;Carries out a command in another application

[ON.KEY](ON.KEY.md)&nbsp;&nbsp;&nbsp;Runs a macro when a specified key is pressed



Return to [README](README.md)

