RESTART

Removes a number of RETURN statements from the stack. When one macro
calls another, the RETURN statement at the end of the second macro
returns control to the calling macro. You can use the RESTART function
to determine which macro regains control.

**Syntax**

**RESTART**(level\_num)

Level\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the number of
previous RETURN statements you want to be ignored. If level\_num is
omitted, the next RETURN statement will halt macro execution.

For example, if the currently running macro has two "ancestors" (the
current macro was called by one macro that, in turn, was called by
another macro), using RESTART(1) in the third macro returns control to
the first calling macro when the RETURN statement is encountered. The
RESTART(1) formula removes one level of RETURN statements from Microsoft
Excel's memory so that the second macro is skipped.

**Remarks**

RESTART is particularly useful if you frequently use macros to call
other macros that in turn call other macros. Use RESTART in combination
with IF statements to prevent macro execution from returning to macros
that called, either directly or indirectly, the currently running macro.

**Related Functions**

[HALT](HALT.md)&nbsp;&nbsp;&nbsp;Stops all macros from running

[RETURN](RETURN.md)&nbsp;&nbsp;&nbsp;Ends the currently running macro



Return to [README](README.md)

