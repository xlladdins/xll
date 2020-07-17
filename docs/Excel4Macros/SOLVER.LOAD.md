SOLVER.LOAD

Equivalent to clicking the Solver command on the Tools menu, clicking
the Options button in the Solver Parameters dialog box, and clicking the
Load Model button in the Solver Options dialog box. Loads Solver problem
specifications that you have previously saved on the worksheet.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.LOAD**(**load\_area**)

Load\_area    is a reference on the active sheet to a range of cells
from which you want to load a complete problem specification.

  - > The first cell in load\_area contains a formula for the Set Cell
    > box; the second cell contains a formula for the changing cells;
    > subsequent cells contain constraints in the form of logical
    > formulas. The last cell optionally contains an array of Solver
    > option values. The order of the Solver option values is the same
    > as the top-to-bottom order in the Solver Options dialog box.

  - > Although load\_area must be on the active sheet, it need not be
    > the current selection.



Return to [README](README.md)

