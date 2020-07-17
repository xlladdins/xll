TRACER.NAVIGATE

Equivalent to double-clicking on a displayed tracer arrow. Moves the
selection from one end of a tracer arrow to the other. If it is an error
tracer arrow, then the selection goes to the end of the branch.

**Syntax**

**TRACER.NAVIGATE**(direction, arrow\_num, ref\_num)

Direction    is a logical value which, if TRUE, moves the selection to
the arrow endpoint in the precedents direction. If FALSE, moves the
selection to the arrow endpoint in the dependents direction.

Arrow\_num    is a number specifying which reference a tracer arrow will
follow. For example, a 1 indicates that the arrow will follow the first
reference in the formula. 1 is the default.

Ref\_num    If the arrow is an external reference arrow with multiple
links, this argument tells which of the links to follow. Refer to the
Links dialog, which is displayed with the Links command from the Edit
menu. If ref\_num is 1, the link in the first reference in the Links
dialog box will be followed. The default is 1.

**Remarks**

  - > Returns TRUE if successful. Returns FALSE if arrow\_num exceeds
    > the number of tracer arrows or if there are no tracer arrows.

  - > Returns FALSE if ref\_num exceeds the number of links.

  - > Returns the \#VALUE\! error value if not available; for example,
    > if the selection is something other than a worksheet, or the
    > active cell does not contain an arrow.

**Related Function**

[TRACER.DISPLAY](TRACER.DISPLAY.md)   Allows tracer arrow to be displayed showing which cells
formulas in other cells depend on



Return to [README.md](README.md)

TRACER.NAVIGATE

Equivalent to double-clicking on a displayed tracer arrow. Moves the
selection from one end of a tracer arrow to the other. If it is an error
tracer arrow, then the selection goes to the end of the branch.

**Syntax**

**TRACER.NAVIGATE**(direction, arrow\_num, ref\_num)

Direction    is a logical value which, if TRUE, moves the selection to
the arrow endpoint in the precedents direction. If FALSE, moves the
selection to the arrow endpoint in the dependents direction.

Arrow\_num    is a number specifying which reference a tracer arrow will
follow. For example, a 1 indicates that the arrow will follow the first
reference in the formula. 1 is the default.

Ref\_num    If the arrow is an external reference arrow with multiple
links, this argument tells which of the links to follow. Refer to the
Links dialog, which is displayed with the Links command from the Edit
menu. If ref\_num is 1, the link in the first reference in the Links
dialog box will be followed. The default is 1.

**Remarks**

  - > Returns TRUE if successful. Returns FALSE if arrow\_num exceeds
    > the number of tracer arrows or if there are no tracer arrows.

  - > Returns FALSE if ref\_num exceeds the number of links.

  - > Returns the \#VALUE\! error value if not available; for example,
    > if the selection is something other than a worksheet, or the
    > active cell does not contain an arrow.

**Related Function**

[TRACER.DISPLAY](TRACER.DISPLAY.md)   Allows tracer arrow to be displayed showing which cells
formulas in other cells depend on



Return to [README.md](README.md)

