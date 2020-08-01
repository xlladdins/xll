# SOLVER.SAVE

Equivalent to clicking the Solver command on the Tools menu, clicking
the Options button in the Solver Parameters dialog box, and clicking the
Save Model button in the Solver Options dialog box. Saves the Solver
problem specifications on the worksheet.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.SAVE** (**save\_area**)

Save\_area&nbsp;&nbsp;&nbsp;&nbsp;is a reference on the active sheet to
a range of cells or to the upper-left corner of a range of cells into
which you want to paste the current problem specification.

  - > If you specify only one cell for save\_area, the area is extended
    > downwards for as many cells as are required to hold the problem
    > specifications (3 plus the number of constraints).

  - > If you specify more than one cell and if the area is too small,
    > the last constraints (in alphabetic order by cell reference) or
    > options will be omitted and the function will return a nonzero
    > value.

  - > Save\_area must be on the active worksheet, but it need not be the
    > current selection.



Return to [README](README.md#S)

# SOLVER.SAVE

Equivalent to clicking the Solver command on the Tools menu, clicking
the Options button in the Solver Parameters dialog box, and clicking the
Save Model button in the Solver Options dialog box. Saves the Solver
problem specifications on the worksheet.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.SAVE** (**save\_area**)

Save\_area&nbsp;&nbsp;&nbsp;&nbsp;is a reference on the active sheet to
a range of cells or to the upper-left corner of a range of cells into
which you want to paste the current problem specification.

  - > If you specify only one cell for save\_area, the area is extended
    > downwards for as many cells as are required to hold the problem
    > specifications (3 plus the number of constraints).

  - > If you specify more than one cell and if the area is too small,
    > the last constraints (in alphabetic order by cell reference) or
    > options will be omitted and the function will return a nonzero
    > value.

  - > Save\_area must be on the active worksheet, but it need not be the
    > current selection.



Return to [README](README.md#S)

