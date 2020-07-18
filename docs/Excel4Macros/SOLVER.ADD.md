SOLVER.ADD

Equivalent to clicking the Solver command on the Tools menu and clicking
the Add button in the Solver Parameters dialog box. Adds a constraint to
the current problem. For an explanation of constraints, see "Remarks"
later in this topic.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.ADD**(**cell\_ref, relation**, formula)

Cell\_ref**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a reference to a cell or range of
cells on the active sheet and forms the left side of the constraint.

Relation**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies the arithmetic relationship
between the left and right sides, or whether cell\_ref must be an
integer.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Relation</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Arithmetic relationship</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>&lt;=</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>=</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>&gt;=</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Int (cell_ref is an integer)</p>
</blockquote></td>
</tr>
</tbody>
</table>

Formula**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the right side of the constraint and
will often be a single number, but it may be a formula (as text) or a
reference to a range of cells.

  - > If relation is 4, cell\_ref must be a subset of the references in
    > the By Changing cells text box.

  - > if relation is 4, formula must be either "=integer" or "integer".

  - > Any cell reference in a formula must use the R1C1 reference style.

  - > If formula is a reference to a range of cells, the number of cells
    > in the range usually matches the number of cells in cell\_ref,
    > although the shape of the areas need not be the same. For example,
    > cell\_ref could be a row and formula could refer to a column, as
    > long as the number of cells is the same. Formula can also be a
    > single reference, as in the following relationship:
    > &nbsp;A1:A4&nbsp;\<=&nbsp;B1.


**Remarks**

  - > The SOLVER.ADD, SOLVER.CHANGE, and SOLVER.DELETE functions
    > correspond to the Add, Change, and Delete buttons in the Solver
    > Parameters dialog box. You use these functions to define
    > constraints. For many macro applications, however, you may find it
    > more convenient to load the problem specifications from the sheet
    > in a single step using the SOLVER.LOAD function.

  - > Each constraint is uniquely identified by the combination of the
    > cell reference on the left and the relationship (\<=, =, or \>=)
    > between its left and right sides, or the cell reference may be
    > defined as an integer only. This takes the place of selecting the
    > appropriate constraint in the Tools Solver Parameters dialog box.
    > You can manipulate the constraints with SOLVER.CHANGE or
    > SOLVER.DELETE. The constraints in a Solver problem can refer to a
    > maximum of 400 cells.


**Related Function**

[SOLVER.DELETE](SOLVER.DELETE.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Deletes an existing constraint



Return to [README](README.md)

