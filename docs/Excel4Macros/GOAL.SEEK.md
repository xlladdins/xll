GOAL.SEEK

Equivalent to clicking the Goal Seek command on the Tools menu.
Calculates the values necessary to achieve a specific goal. If the goal
is an amount returned by a formula, the GOAL.SEEK function calculates
values that, when supplied to your formula, cause your formula to return
the amount you want.

**Syntax**

**GOAL.SEEK**(**target\_cell, target\_value, variable\_cell**)

**GOAL.SEEK**?(target\_cell, target\_value, variable\_cell)

Target\_cell**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Set Cell box in
the Goal Seek dialog box and is a reference to the cell containing the
formula. If target\_cell does not contain a formula, Microsoft Excel
displays an error message.

Target\_value**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the To Value box in
the Goal Seek dialog box and is the value you want the formula in
target\_cell to return. This value is called a goal.

Variable\_cell**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the By Changing
Cell box in the Goal Seek dialog box and is the single cell that you
want Microsoft Excel to change so that the formula in target\_cell
returns target\_value. Target\_cell must depend on variable\_cell; if it
does not, Microsoft Excel will not be able to find a solution.

**Remarks**

The max\_num and max\_change values set with the CALCULATION function
can be used to change the solution process. Max\_num sets the number of
iterations; max\_change determines the precision of the solution.

**Tip****&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;You can also use Microsoft Excel Solver to help
solve your math equations for optimal values.

**Related Functions**

[R](R.md)elated functions include the SOLVER functions, such as SOLVER.OPTIONS,
[SOLVER.SOLVE](SOLVER.SOLVE.md), and so on.



Return to [README](README.md)

