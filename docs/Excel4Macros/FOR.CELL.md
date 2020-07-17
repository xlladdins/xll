FOR.CELL
========

Starts a FOR.CELL-NEXT loop. This function is similar to FOR, except
that the instructions between FOR.CELL and NEXT are repeated over a
range of cells, one cell at a time, and there is no loop counter.

**Syntax**

**FOR.CELL**(**ref\_name**, area\_ref, skip\_blanks)

Ref\_name    is the name in the form of text that Microsoft Excel gives
to the one cell in the range that is currently being operated on;
ref\_name refers to a new cell during each loop.

Area\_ref    is the range of cells on which you want the FOR.CELL-NEXT
loop to operate and can be a multiple selection. If area\_ref is
omitted, it is assumed to be the current selection.

Skip\_blanks    is a logical value specifying whether Microsoft Excel
skips blank cells as it operates on the cells in area\_ref.

  ------------------ ------------------------------------
  **Skip\_blanks**   **Result**
  TRUE               Skips blank cells in area\_ref
  FALSE or omitted   Operates on all cells in area\_ref
  ------------------ ------------------------------------

**Remarks**

FOR.CELL operates on each cell in a row from left to right one area at a
time before moving to the next row in the selection.

**Example**

The following macro starts a FOR.CELL-NEXT loop and uses the name
CurrentCell to refer to the cell in the range that is currently being
operated on:

FOR.CELL(\"CurrentCell\", SELECTION(), TRUE)

**Related Functions**

BREAK   Interrupts a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

FOR   Starts a FOR-NEXT loop

NEXT   Ends a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

WHILE   Starts a WHILE-NEXT loop

Return to [top](#E)

FORMAT.AUTO
