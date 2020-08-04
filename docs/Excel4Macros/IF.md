# IF

Used with ELSE, ELSE.IF, and END.IF to control which formulas in a macro
are executed. There are two syntax forms of the IF function. The
following syntax can be used on macro sheets only; use it when you want
your macro to branch to a particular set of functions based on the
outcome of a logical test. The worksheet form of this function can be
used on worksheets and macro sheets.

**Syntax**

**IF**(**logical\_test**)

Logical\_test&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that IF uses to
determine which functions to carry out next&mdash;that is, where to
branch.

  - > If logical\_test is TRUE, Microsoft Excel carries out the
    > functions between the IF function and the next ELSE, ELSE.IF, or
    > END.IF function. Instructions between ELSE.IF or ELSE and END.IF
    > are not carried out.

  - > If logical\_test is FALSE, Microsoft Excel immediately branches to
    > the next ELSE.IF, ELSE, or END.IF function.

  - > If logical\_test produces an error, the macro halts.

**Tips**

  - > Use IF with ELSE, ELSE.IF, and END.IF when you want to perform
    > multiple actions based on a condition. This method is preferable
    > to using GOTO because it makes your macros more structured.

  - > If your macro ends with an error at a cell containing this form of
    > the IF function, make sure there is a corresponding END.IF
    > function.

**Example**

The following macro runs the macro CompleteEntry if the user clicks OK:

IF(ALERT("Are you done with this entry?", 1), CompleteEntry(), )

**Tip**&nbsp;&nbsp;&nbsp;You can indent formulas in a macro. To indent a
formula, type as many spaces as you want between the equal sign and the
first letter of the formula.

**Related Functions**

[ELSE](ELSE.md)&nbsp;&nbsp;&nbsp;Specifies an action to take if an IF function
returns FALSE

[ELSE.IF](ELSE.IF.md)&nbsp;&nbsp;&nbsp;Specifies an action to take if an IF or another
[ELSE.IF](ELSE.IF.md) function returns FALSE

[END.IF](END.IF.md)&nbsp;&nbsp;&nbsp;Ends a group of macro functions started with an
[IF](IF.md) statement

[ERROR](ERROR.md)&nbsp;&nbsp;&nbsp;Specifies what action to take if an error occurs
while a macro is running



Return to [README](README.md#I)

