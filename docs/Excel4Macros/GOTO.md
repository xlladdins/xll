GOTO

Directs a macro to continue running at the upper-left cell of reference.
Use GOTO to direct macro execution to another cell or a named range.

**Syntax**

**GOTO**(**reference**)

Reference&nbsp;&nbsp;&nbsp;&nbsp;is a cell reference or a name that is
defined as a reference. Reference can be an external reference to
another macro sheet. If that macro sheet is not open, GOTO displays a
message.

**Tip**&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;It's often preferable to use IF, ELSE, ELSE.IF,
and END.IF instead of GOTO when you want to perform multiple actions
based on a condition because the IF method makes your macros more
structured.

**Examples**

If A1 contains the \#N/A error value, then when the following formula is
calculated, the macro branches to C3:

IF(ISERROR($A$1), GOTO($C$3))

You can also use macro names with GOTO statements. The following macro
formula branches macro execution to a macro named Compile:

GOTO(Compile)

Because Compile is a named range, it should not be enclosed in quotation
marks.

**Related Function**

[FORMULA.GOTO](FORMULA.GOTO.md)&nbsp;&nbsp;&nbsp;Selects a named area or reference on any
open workbook



Return to [README](README.md)

