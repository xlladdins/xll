ELSE.IF

Used with IF, ELSE, and END.IF to control which functions are carried
out in a macro. ELSE.IF signals the beginning of a group of formulas in
a macro sheet that will be carried out if the preceding IF or ELSE.IF
function returns FALSE and if logical\_test is TRUE. Use ELSE.IF with
IF, ELSE, and END.IF when you want to perform multiple actions based on
a condition. This method is preferable to using GOTO because it makes
your macros more structured.

**Syntax**

**ELSE.IF**(**logical\_test**)

Logical\_test    is a logical value that ELSE.IF uses to determine what
functions to carry out next—that is, where to branch.

  - > If logical\_test is TRUE, Microsoft Excel carries out the
    > functions between the ELSE.IF function and the next ELSE.IF, ELSE,
    > or END.IF function.

  - > If logical\_test is FALSE, Microsoft Excel immediately branches to
    > the next ELSE.IF, ELSE, or END.IF function.

>  

**Remarks**

  - > ELSE.IF must be entered in a cell by itself.

  - > Logical\_test will always be evaluated, even if the ELSE.IF
    > section is not reached (due to a previous IF or ELSE.IF
    > logical\_test evaluating to TRUE). For this reason, you should not
    > use formulas that carry out actions for logical\_test. If you need
    > to base the ELSE.IF condition on the return value of a formula
    > that carries out an action, use the form "ELSE, IF(logical\_test),
    > and END.IF" in place of "ELSE.IF(logical\_test)."

>  

For more information about ELSE, ELSE.IF, END.IF, and IF, and for
examples of these functions, see form 2 of the IF function.

**Related Functions**

[ELSE](ELSE.md)   Specifies an action to take if an IF function returns FALSE

[END.IF](END.IF.md)   Ends a group of macro functions started with an IF statement

[IF](IF.md)   Specifies an action to take if a logical test is TRUE



Return to [README.md](README.md)

ELSE.IF

Used with IF, ELSE, and END.IF to control which functions are carried
out in a macro. ELSE.IF signals the beginning of a group of formulas in
a macro sheet that will be carried out if the preceding IF or ELSE.IF
function returns FALSE and if logical\_test is TRUE. Use ELSE.IF with
IF, ELSE, and END.IF when you want to perform multiple actions based on
a condition. This method is preferable to using GOTO because it makes
your macros more structured.

**Syntax**

**ELSE.IF**(**logical\_test**)

Logical\_test    is a logical value that ELSE.IF uses to determine what
functions to carry out next—that is, where to branch.

  - > If logical\_test is TRUE, Microsoft Excel carries out the
    > functions between the ELSE.IF function and the next ELSE.IF, ELSE,
    > or END.IF function.

  - > If logical\_test is FALSE, Microsoft Excel immediately branches to
    > the next ELSE.IF, ELSE, or END.IF function.

>  

**Remarks**

  - > ELSE.IF must be entered in a cell by itself.

  - > Logical\_test will always be evaluated, even if the ELSE.IF
    > section is not reached (due to a previous IF or ELSE.IF
    > logical\_test evaluating to TRUE). For this reason, you should not
    > use formulas that carry out actions for logical\_test. If you need
    > to base the ELSE.IF condition on the return value of a formula
    > that carries out an action, use the form "ELSE, IF(logical\_test),
    > and END.IF" in place of "ELSE.IF(logical\_test)."

>  

For more information about ELSE, ELSE.IF, END.IF, and IF, and for
examples of these functions, see form 2 of the IF function.

**Related Functions**

[ELSE](ELSE.md)   Specifies an action to take if an IF function returns FALSE

[END.IF](END.IF.md)   Ends a group of macro functions started with an IF statement

[IF](IF.md)   Specifies an action to take if a logical test is TRUE



Return to [README.md](README.md)

