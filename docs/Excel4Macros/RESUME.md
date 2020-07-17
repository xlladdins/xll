RESUME

Equivalent to choosing the Resume button on the toolbar. Resumes a
paused macro. Returns TRUE if successful or the \#VALUE\! error value if
no macro is paused. A macro can be paused by using the PAUSE function or
choosing Pause from the Single Step dialog box, which appears when you
choose the Step Into button from the Macro dialog box.

**Syntax**

**RESUME**(type\_num)

Type\_num    is a number from 1 to 4 specifying how to resume.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>How Microsoft Excel resumes</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1 or omitted</p>
</blockquote></td>
<td><blockquote>
<p>If paused by a PAUSE function, continues running the macro. If paused from the Single Step dialog box, returns to that dialog box.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Halts the paused macro</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Continues running the macro</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Opens the Single Step dialog box</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Tip**   You can use Microsoft Excel's ON functions to resume based on
an event. For an example, see ENTER.DATA.

**Remarks**

  - > If one macro runs a second macro that pauses, and you need to halt
    > only the paused macro, use RESUME(2) instead of HALT. HALT halts
    > all macros and prevents resuming or returning to any macro.

  - > If the macro was paused from the Single Step dialog box, RESUME
    > returns to the Single Step dialog box.

>  

**Related Functions**

[HALT](HALT.md)   Stops all macros from running

[PAUSE](PAUSE.md)   Pauses a macro

[RETURN](RETURN.md)   Ends the currently running macro



Return to [README](README.md)

