SOLVER.CHANGE
=============

Equivalent to clicking the Solver command on the Tools menu and clicking
the Change button in the Solver Parameters dialog box. Changes the right
side of an existing constraint.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.CHANGE**(**cell\_ref, relation**, formula)

For an explanation of the arguments and constraints, see SOLVER.ADD.

**Remarks**

-   If the combination of cell\_ref and relation does not match any
    > existing constraint, the function returns the value 4 and no
    > action is taken.

-   To change the cell\_ref or relation of an existing constraint, use
    > SOLVER.DELETE to delete the old constraint and then use SOLVER.ADD
    > to add the constraint in the form you want.

>  

**Related Functions**

SOLVER.DELETE   Deletes an existing constraint

SOLVER.ADD   Adds a constraint to the current problem

Return to [top](#Q)

SOLVER.DELETE
