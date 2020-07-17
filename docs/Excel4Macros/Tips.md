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

**Tip**   You can indent formulas in a macro. To indent a formula, type
as many spaces as you want between the equal sign and the first letter
of the formula.

**Related Functions**

[ELSE](ELSE.md)   Specifies an action to take if an IF function returns FALSE

[ELSE.IF](ELSE.IF.md)   Specifies an action to take if an IF or another ELSE.IF
function returns FALSE

[END.IF](END.IF.md)   Ends a group of macro functions started with an IF statement

[ERROR](ERROR.md)   Specifies what action to take if an error occurs while a macro
is running



Return to [README.md](README.md)

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

**Tip**   You can indent formulas in a macro. To indent a formula, type
as many spaces as you want between the equal sign and the first letter
of the formula.

**Related Functions**

[ELSE](ELSE.md)   Specifies an action to take if an IF function returns FALSE

[ELSE.IF](ELSE.IF.md)   Specifies an action to take if an IF or another ELSE.IF
function returns FALSE

[END.IF](END.IF.md)   Ends a group of macro functions started with an IF statement

[ERROR](ERROR.md)   Specifies what action to take if an error occurs while a macro
is running



Return to [README.md](README.md)

