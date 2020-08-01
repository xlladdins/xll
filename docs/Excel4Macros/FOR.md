# FOR

Starts a FOR-NEXT loop. The instructions between FOR and NEXT are
repeated until the loop counter reaches a specified value. Use FOR when
you need to repeat instructions a specified number of times. Use
FOR.CELL when you need to repeat instructions over a range of cells.

**Syntax**

**FOR**(**counter\_text, start\_num, end\_num**, step\_num)

Counter\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of the loop counter in
the form of text.

Start\_num&nbsp;&nbsp;&nbsp;&nbsp;is the value initially assigned to
counter\_text.

End\_num&nbsp;&nbsp;&nbsp;&nbsp;is the last value assigned to
counter\_text.

Step\_num&nbsp;&nbsp;&nbsp;&nbsp;is a value added to the loop counter
after each iteration. If step\_num is omitted, it is assumed to be 1.

**Remarks**

  - > Microsoft Excel follows these steps as it executes a FOR-NEXT
    > loop:

<table>
<tbody>
<tr class="odd">
<td><strong>Step</strong></td>
<td><strong>Action</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Sets counter_text to the value start_num.</td>
</tr>
<tr class="odd">
<td>2</td>
<td><p>If counter_text is greater than end_num (or less than end_num if step_num is negative), the loop ends, and the macro continues with the function after the NEXT function.</p>
<p>If counter_text is less than or equal to end_num (or greater than or equal to end_num if step_num is negative), the macro continues in the loop.</p></td>
</tr>
<tr class="even">
<td>3</td>
<td>Carries out functions up to the following NEXT function. The NEXT function must be below the FOR function and in the same column.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Adds step_num to the loop counter.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns to the FOR function and proceeds as described in step 2.</td>
</tr>
</tbody>
</table>

&nbsp;

  - > You can interrupt a FOR-NEXT loop by using the BREAK function.


**Example**

The following macro starts a FOR-NEXT loop that is executed once for
every open window:

FOR("Counter", 1, COLUMNS(WINDOWS()))

**Related Functions**

[BREAK](BREAK.md)&nbsp;&nbsp;&nbsp;Interrupts a FOR-NEXT, FOR.CELL-NEXT, or
[WHILE](WHILE.md)-NEXT loop

[FOR.CELL](FOR.CELL.md)&nbsp;&nbsp;&nbsp;Starts a FOR.CELL-NEXT loop

[NEXT](NEXT.md)&nbsp;&nbsp;&nbsp;Ends a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

[WHILE](WHILE.md)&nbsp;&nbsp;&nbsp;Starts a WHILE-NEXT loop



Return to [README](README.md#F)

# FOR

Starts a FOR-NEXT loop. The instructions between FOR and NEXT are
repeated until the loop counter reaches a specified value. Use FOR when
you need to repeat instructions a specified number of times. Use
FOR.CELL when you need to repeat instructions over a range of cells.

**Syntax**

**FOR**(**counter\_text, start\_num, end\_num**, step\_num)

Counter\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of the loop counter in
the form of text.

Start\_num&nbsp;&nbsp;&nbsp;&nbsp;is the value initially assigned to
counter\_text.

End\_num&nbsp;&nbsp;&nbsp;&nbsp;is the last value assigned to
counter\_text.

Step\_num&nbsp;&nbsp;&nbsp;&nbsp;is a value added to the loop counter
after each iteration. If step\_num is omitted, it is assumed to be 1.

**Remarks**

  - > Microsoft Excel follows these steps as it executes a FOR-NEXT
    > loop:

<table>
<tbody>
<tr class="odd">
<td><strong>Step</strong></td>
<td><strong>Action</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Sets counter_text to the value start_num.</td>
</tr>
<tr class="odd">
<td>2</td>
<td><p>If counter_text is greater than end_num (or less than end_num if step_num is negative), the loop ends, and the macro continues with the function after the NEXT function.</p>
<p>If counter_text is less than or equal to end_num (or greater than or equal to end_num if step_num is negative), the macro continues in the loop.</p></td>
</tr>
<tr class="even">
<td>3</td>
<td>Carries out functions up to the following NEXT function. The NEXT function must be below the FOR function and in the same column.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Adds step_num to the loop counter.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns to the FOR function and proceeds as described in step 2.</td>
</tr>
</tbody>
</table>

&nbsp;

  - > You can interrupt a FOR-NEXT loop by using the BREAK function.


**Example**

The following macro starts a FOR-NEXT loop that is executed once for
every open window:

FOR("Counter", 1, COLUMNS(WINDOWS()))

**Related Functions**

[BREAK](BREAK.md)&nbsp;&nbsp;&nbsp;Interrupts a FOR-NEXT, FOR.CELL-NEXT, or
[WHILE](WHILE.md)-NEXT loop

[FOR.CELL](FOR.CELL.md)&nbsp;&nbsp;&nbsp;Starts a FOR.CELL-NEXT loop

[NEXT](NEXT.md)&nbsp;&nbsp;&nbsp;Ends a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

[WHILE](WHILE.md)&nbsp;&nbsp;&nbsp;Starts a WHILE-NEXT loop



Return to [README](README.md#F)

