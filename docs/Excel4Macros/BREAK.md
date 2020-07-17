BREAK

Interrupts a FOR-NEXT, a FOR.CELL-NEXT, or a WHILE-NEXT loop. If BREAK
is encountered within a loop, that loop is terminated and the macro
proceeds to the statement following the NEXT statement at the end of the
current loop.

**Syntax**

**BREAK**( )

**Example**

Use BREAK to test for conditions not anticipated by the FOR or WHILE
statement. For example, use the BREAK nested in an IF statement to exit
a WHILE-NEXT loop when a certain value is encountered:

\=IF(Counter=8, BREAK())

**Related Functions**

[FOR](FOR.md)   Starts a FOR-NEXT loop

[FOR.CELL](FOR.CELL.md)   Starts a FOR.CELL-NEXT loop

[NEXT](NEXT.md)   Ends a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

[WHILE](WHILE.md)   Starts a WHILE-NEXT loop



Return to [README](README.md)

