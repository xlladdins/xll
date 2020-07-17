SOLVER.OK

Equivalent to clicking the Solver command on the Tools menu and
specifying options in the Solver Parameters dialog box. Specifies basic
Solver options, except that constraints are added via SOLVER.ADD.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.OK**(set\_cell, max\_min\_val, value\_of, **by\_changing**)

**SOLVER.OK**?(set\_cell, max\_min\_val, value\_of, by\_changing)

Set\_cell    corresponds to the Set Target Cell box in the Solver
Parameters dialog box.

  - > Set\_cell must be a reference to a cell on the active worksheet.

  - > If you enter a cell reference, you must also enter a value for
    > max\_min\_val. If you do not enter a cell, you must include three
    > commas before the by\_changing value.

>  

Max\_min\_val    corresponds to the options Max, Min, and Value Of in
the Solver Parameters dialog box. Use this option only if you entered a
reference for set\_cell.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Max_min_val</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Option specified</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Maximize</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Minimize</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Match specific value</p>
</blockquote></td>
</tr>
</tbody>
</table>

Value\_of    is a number that becomes the target for the cell in the Set
Target Cell box if max\_min\_val is 3. Value\_of is ignored if the cell
is being maximized or minimized.

By\_changing    indicates the changing cells, as entered in the By
Changing Cells box. By\_changing must refer to a cell or range of cells
on the active worksheet, and can be a multiple selection.

**Remarks**

The constraints in a Solver problem can refer to a maximum of 400 cells.

**Related Function**

[SOLVER.SOLVE](SOLVER.SOLVE.md)   Returns an integer value indicating the condition that
caused Solver to stop



Return to [README.md](README.md)

SOLVER.OK

Equivalent to clicking the Solver command on the Tools menu and
specifying options in the Solver Parameters dialog box. Specifies basic
Solver options, except that constraints are added via SOLVER.ADD.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.OK**(set\_cell, max\_min\_val, value\_of, **by\_changing**)

**SOLVER.OK**?(set\_cell, max\_min\_val, value\_of, by\_changing)

Set\_cell    corresponds to the Set Target Cell box in the Solver
Parameters dialog box.

  - > Set\_cell must be a reference to a cell on the active worksheet.

  - > If you enter a cell reference, you must also enter a value for
    > max\_min\_val. If you do not enter a cell, you must include three
    > commas before the by\_changing value.

>  

Max\_min\_val    corresponds to the options Max, Min, and Value Of in
the Solver Parameters dialog box. Use this option only if you entered a
reference for set\_cell.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Max_min_val</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Option specified</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Maximize</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Minimize</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Match specific value</p>
</blockquote></td>
</tr>
</tbody>
</table>

Value\_of    is a number that becomes the target for the cell in the Set
Target Cell box if max\_min\_val is 3. Value\_of is ignored if the cell
is being maximized or minimized.

By\_changing    indicates the changing cells, as entered in the By
Changing Cells box. By\_changing must refer to a cell or range of cells
on the active worksheet, and can be a multiple selection.

**Remarks**

The constraints in a Solver problem can refer to a maximum of 400 cells.

**Related Function**

[SOLVER.SOLVE](SOLVER.SOLVE.md)   Returns an integer value indicating the condition that
caused Solver to stop



Return to [README.md](README.md)

