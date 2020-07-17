FOR
===

Starts a FOR-NEXT loop. The instructions between FOR and NEXT are
repeated until the loop counter reaches a specified value. Use FOR when
you need to repeat instructions a specified number of times. Use
FOR.CELL when you need to repeat instructions over a range of cells.

**Syntax**

**FOR**(**counter\_text, start\_num, end\_num**, step\_num)

Counter\_text    is the name of the loop counter in the form of text.

Start\_num    is the value initially assigned to counter\_text.

End\_num    is the last value assigned to counter\_text.

Step\_num    is a value added to the loop counter after each iteration.
If step\_num is omitted, it is assumed to be 1.

**Remarks**

-   Microsoft Excel follows these steps as it executes a FOR-NEXT loop:

+----------+----------------------------------------------------------+
| **Step** | **Action**                                               |
+----------+----------------------------------------------------------+
| 1        | Sets counter\_text to the value start\_num.              |
+----------+----------------------------------------------------------+
| 2        | If counter\_text is greater than end\_num (or less than  |
|          | end\_num if step\_num is negative), the loop ends, and   |
|          | the macro continues with the function after the NEXT     |
|          | function.                                                |
|          |                                                          |
|          | If counter\_text is less than or equal to end\_num (or   |
|          | greater than or equal to end\_num if step\_num is        |
|          | negative), the macro continues in the loop.              |
+----------+----------------------------------------------------------+
| 3        | Carries out functions up to the following NEXT function. |
|          | The NEXT function must be below the FOR function and in  |
|          | the same column.                                         |
+----------+----------------------------------------------------------+
| 4        | Adds step\_num to the loop counter.                      |
+----------+----------------------------------------------------------+
| 5        | Returns to the FOR function and proceeds as described in |
|          | step 2.                                                  |
+----------+----------------------------------------------------------+

 

-   You can interrupt a FOR-NEXT loop by using the BREAK function.

>  

**Example**

The following macro starts a FOR-NEXT loop that is executed once for
every open window:

FOR(\"Counter\", 1, COLUMNS(WINDOWS()))

**Related Functions**

BREAK   Interrupts a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

FOR.CELL   Starts a FOR.CELL-NEXT loop

NEXT   Ends a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

WHILE   Starts a WHILE-NEXT loop

Return to [top](#E)

FOR.CELL
