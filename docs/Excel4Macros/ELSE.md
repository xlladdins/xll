ELSE
====

Used with IF, ELSE.IF, and END.IF to control which functions are carried
out in a macro. ELSE signals the beginning of a group of formulas in a
macro sheet that will be carried out if the results of all preceding
ELSE.IF statements and the preceding IF statement are FALSE. Use ELSE
with IF, ELSE.IF, and END.IF when you want to perform multiple actions
based on a condition. This method is preferable to using GOTO because it
makes your macros more structured.

**Syntax**

**ELSE**( )

**Remarks**

ELSE must be entered in a cell by itself. In other words, the cell can
contain only \"=ELSE()\".

For more information about ELSE, ELSE.IF, END.IF, and IF, and for
examples of these functions, see form 2 of the IF function.

**Related Functions**

ELSE.IF   Specifies an action to take if an IF or another ELSE.IF
function returns FALSE

END.IF   Ends a group of macro functions started with an IF statement

IF   Specifies an action to take if a logical test is TRUE

Return to [top](#E)

ELSE.IF
