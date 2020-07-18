# END.IF

Ends a block of functions associated with the preceding IF function. You
must include one and only one END.IF function for each macro-sheets-only
syntax form (syntax 2) of the IF function in a macro. Syntax 1 of the IF
function, which can be used on both worksheets and macro sheets, does
not require an END.IF function. Use END.IF with IF, ELSE, and ELSE.IF
when you want to perform multiple actions based on a condition. This
method is preferable to using GOTO because it makes your macros more
structured.

**Syntax**

**END.IF**( )

**Remarks**

  - > If you accidentally omit an END.IF function, your macro will end
    > with an error at the cell containing the first IF function that
    > does not have a corresponding END.IF function.

  - > END.IF must be entered in a cell by itself.

  - > For more information about ELSE, ELSE.IF, END.IF, and IF, and for
    > examples of these functions, see form 2 of the IF function.


**Related Functions**

[ELSE](ELSE.md)&nbsp;&nbsp;&nbsp;Specifies an action to take if an IF function
returns FALSE

[ELSE.IF](ELSE.IF.md)&nbsp;&nbsp;&nbsp;Specifies an action to take if an IF or another
[ELSE.IF](ELSE.IF.md) function returns FALSE

[IF](IF.md)&nbsp;&nbsp;&nbsp;Specifies an action to take if a logical test is
[TRUE](TRUE.md)



Return to [README](README.md)

