CUT

Equivalent to choosing the Cut command from the Edit menu. Cuts or moves
data or objects.

**Syntax**

**CUT**(from\_reference, to\_reference)

From\_reference    is a reference to the cell or range of cells you want
to cut. If from\_reference is omitted, it is assumed to be the current
selection.

To\_reference    is a reference to the cell or range of cells where you
want to paste what you have cut.

  - > To\_reference should be a single cell or an enlarged multiple of
    > from\_reference. For example, if from\_reference is a 2 by 4
    > rectangle, to\_reference can be a 4 by 8 rectangle.

  - > To\_reference can be omitted so that you can paste from\_reference
    > later using the PASTE or PASTE.SPECIAL functions.

>  

**Remarks**

The following information may be helpful if you're having problems with
CUT updating references in unexpected ways. When you move cells using
CUT, formulas that referred to from\_reference will refer to
to\_reference, and formulas that referred to to\_reference may return
\#REF\! error values. However, if from\_reference or to\_reference
contains references that are calculated at runtime (for example,
CUT(ACTIVE.CELL(), \!B1)), then Microsoft Excel does not update those
references when the CUT function is run, so no error values are
returned.

**Related Functions**

COPY   Copies and pastes data or objects

PASTE   Pastes cut or copied data


