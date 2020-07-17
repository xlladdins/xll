SCENARIO.CELLS

Equivalent to clicking the Scenarios command on the Tools menu and then
editing the Changing Cells box. Defines the changing cells for a model
on your worksheet. Changing cells are the cells into which values will
be entered when you display a scenario. If you have only one set of
changing cells on your sheet, SCENARIO.CELLS will change the changing
cells for all scenarios. If your sheet has scenarios defined with
multiple sets of changing cell, this function returns an error and the
macro is halted. This function is provided for compatibility with
Microsoft Excel version 4.0. Use SCENARIO.EDIT with the changing\_ref
argument instead of SCENARIO.CELLS if you want to change the changing
cells of a scenario.

**Syntax**

**SCENARIO.CELLS**(**changing\_ref**)

**SCENARIO.CELLS**? (changing\_ref)

Changing\_ref    is a reference to the cells you want to define as
changing cells for the model. If changing\_ref contains nonadjacent
references, you must separate the reference areas by commas and enclose
changing\_ref in an extra set of parentheses.

**Related Function**

[SCENARIO.EDIT](SCENARIO.EDIT.md)   Equivalent to clicking the Scenarios command on the
[T](T.md)ools menu and then clicking the Edit button


