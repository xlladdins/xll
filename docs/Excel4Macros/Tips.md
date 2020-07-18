Tips

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



Return to [README](README.md)

