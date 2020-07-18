RUN

Equivalent to choosing the Run button in the Macro dialog box, which
appears when you choose the Macros command on the Macro submenu of the
Tools menu. Runs a macro.

**Syntax**

**RUN**(reference, step)

**RUN**?(reference, step)

Reference**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a reference to the macro you want to
run or a number from 1 to 4 specifying an Auto macro to run.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>If reference is</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Specifies</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>All Auto_Open macros on the active workbook</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>All Auto_Close macros</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>All Auto_Activate macros</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>All Auto_Deactivate macros</p>
</blockquote></td>
</tr>
</tbody>
</table>

&nbsp;

  - > If reference is a range of cells, RUN begins with the macro
    > function in the upper-left cell of reference.

  - > If the macro sheet containing the macro is not in the active
    > workbook, reference can be an external reference to the name of
    > the macro, such as RUN(\[BOOK1\]Macro\!Months) or an external
    > R1C1-style reference to the location of the macro, such as
    > RUN("\[Book1\]Macro\!R2C3"). The reference must be in text form.

  - > If reference is omitted, the macro function in the active cell is
    > carried out, and macro execution continues down that column.


Step**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value specifying that the macro
is to be run in single-step mode. If step is TRUE, Microsoft Excel runs
the macro in single-step mode; if FALSE or omitted, Microsoft Excel runs
the macro normally.

**Remarks**

  - > RUN is recorded when you choose the Run button the Macro dialog
    > box while recording a macro. The reference you enter in the Run
    > dialog box is recorded as text, with A1-style references converted
    > to R1C1-style references.

  - > To run a macro from a macro sheet, you could alternatively enter
    > the name of the macro as a formula, followed by a set of
    > parentheses. For example, enter =\[Book1\]Macro\!Months() instead
    > of =RUN(\[Book1\]Macro\!Months).


**Related Function**

[GOTO](GOTO.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Directs macro execution to another cell



Return to [README](README.md)

