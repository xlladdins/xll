SOLVER.FINISH

Equivalent to clicking OK in the Solver Results dialog box that appears
when the solution process is complete. The dialog-box form displays the
dialog box with the arguments that you supply as defaults. This function
must be used if you supplied the value TRUE for the userfinish argument
to SOLVER.SOLVE.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.FINISH**(keep\_final, report\_array)

**SOLVER.FINISH**?(keep\_final, report\_array)

Keep\_final    is the number 1 or 2 and specifies whether to keep the
final solution. If keep\_final is 1 or omitted, the final solution
values are kept in the changing cells. If keep\_final is 2, the final
solution values are discarded and the former values of the changing
cells are restored.

Report\_array    is an array argument specifying what reports to create
when Solver is finished.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>If report_array is</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Microsoft Excel creates</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>{1}</p>
</blockquote></td>
<td><blockquote>
<p>An answer report</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>{2}</p>
</blockquote></td>
<td><blockquote>
<p>A sensitivity report</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>{3}</p>
</blockquote></td>
<td><blockquote>
<p>A limit report</p>
</blockquote></td>
</tr>
</tbody>
</table>

Any combination of these produces multiple reports. For example, if
report\_array is {1, 2}, Microsoft Excel creates an answer report and a
sensitivity report.

**Related Function**

SOLVER.SOLVE   Equivalent to clicking the Solver command on the Tools
menu and clicking the Solve button in the Solver Parameters dialog box


