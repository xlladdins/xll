WHILE

Carries out the statements between the WHILE function and the next NEXT
function until logical\_test is FALSE. Use WHILE-NEXT loops to carry out
a series of macro formulas while a certain condition remains TRUE.

**Syntax**

**WHILE**(**logical\_test**)

Logical\_test    is a value or formula that evaluates to TRUE or FALSE.
If logical\_test is FALSE the first time the WHILE function is reached,
the macro skips the loop and resumes running at the statement after the
next NEXT function. If there is no NEXT function in the same column,
WHILE displays an error message and interrupts the macro.

**Remarks**

If you know exactly how many times you'll need to carry out the
statements within a loop, in most cases you should use a FOR-NEXT loop.
Also, avoid creating an infinite loop by making sure logical\_test does
not always evaluate to TRUE.

**Example**

The following statement starts a loop that executes while the value in
the current cell is less than 5:

\=WHILE(TYPE(ACTIVE.CELL()\<5))

The following statement starts a loop that executes until the position
in the open file identified as FileNumber reaches the end of the file:

\=WHILE(FPOS(FileNumber)\<=FSIZE(FileNumber))

**Related Functions**

[FOR](FOR.md)   Starts a FOR-NEXT loop

[FOR.CELL](FOR.CELL.md)   Starts a FOR.CELL-NEXT loop

[IF](IF.md)   Specifies an action to take if a logical test is TRUE

[NEXT](NEXT.md)   Ends a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop



Return to [README.md](README.md)

WHILE

Carries out the statements between the WHILE function and the next NEXT
function until logical\_test is FALSE. Use WHILE-NEXT loops to carry out
a series of macro formulas while a certain condition remains TRUE.

**Syntax**

**WHILE**(**logical\_test**)

Logical\_test    is a value or formula that evaluates to TRUE or FALSE.
If logical\_test is FALSE the first time the WHILE function is reached,
the macro skips the loop and resumes running at the statement after the
next NEXT function. If there is no NEXT function in the same column,
WHILE displays an error message and interrupts the macro.

**Remarks**

If you know exactly how many times you'll need to carry out the
statements within a loop, in most cases you should use a FOR-NEXT loop.
Also, avoid creating an infinite loop by making sure logical\_test does
not always evaluate to TRUE.

**Example**

The following statement starts a loop that executes while the value in
the current cell is less than 5:

\=WHILE(TYPE(ACTIVE.CELL()\<5))

The following statement starts a loop that executes until the position
in the open file identified as FileNumber reaches the end of the file:

\=WHILE(FPOS(FileNumber)\<=FSIZE(FileNumber))

**Related Functions**

[FOR](FOR.md)   Starts a FOR-NEXT loop

[FOR.CELL](FOR.CELL.md)   Starts a FOR.CELL-NEXT loop

[IF](IF.md)   Specifies an action to take if a logical test is TRUE

[NEXT](NEXT.md)   Ends a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop



Return to [README.md](README.md)

