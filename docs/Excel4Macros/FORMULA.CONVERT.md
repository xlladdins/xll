FORMULA.CONVERT

Changes the style and type of references in a formula between A1 and
R1C1 and between relative and absolute. Use FORMULA.CONVERT to convert
references of one style or type to another style or type.

**Syntax**

**FORMULA.CONVERT**(**formula\_text, from\_a1**, to\_a1, to\_ref\_type,
rel\_to\_ref)

Formula\_text    is the formula, given as text, containing the
references you want to change. Formula\_text must be a valid formula,
and an equal sign must be included.

From\_a1    is a logical value specifying whether the references in
formula\_text are in A1 or R1C1 style. If from\_a1 is TRUE, references
are in A1 style; if FALSE, references are in R1C1 style.

To\_a1    is a logical value specifying the form for the references
FORMULA.CONVERT returns. If to\_a1 is TRUE, references are returned in
A1 style; if FALSE, references are returned in R1C1 style. If to\_a1 is
omitted, the reference style is not changed.

To\_ref\_type    is a number from 1 to 4 specifying the reference type
of the returned formula. If to\_ref\_type is omitted, the reference type
is not changed.

|                   |                               |
| ----------------- | ----------------------------- |
| **To\_ref\_type** | **Reference type returned**   |
| 1                 | Absolute                      |
| 2                 | Absolute row, relative column |
| 3                 | Relative row, absolute column |
| 4                 | Relative                      |

Rel\_to\_ref    is an absolute reference that specifies what cell the
relative references are or should be relative to.

**Examples**

Use FORMULA.CONVERT to convert relative references entered by the user
in an INPUT function or custom dialog box into absolute references. The
following macro formula converts the given formula to an absolute,
R1C1-style reference:

FORMULA.CONVERT("=A1:A10", TRUE, FALSE, 1) equals "=R1C1:R10C1"

The following macro formula converts the references in the given formula
to relative, A1-style references:

FORMULA.CONVERT("=SUM(R10C2:R15C2)", FALSE, TRUE, 4) equals
"=SUM(B10:B15)"

**Tip**   To put the converted formula into a cell or range of cells,
use the FORMULA.CONVERT function as the formula\_text argument to the
FORMULA function.

**Related Functions**

ABSREF   Returns the absolute reference of a range of cells to another
range

FORMULA   Enters values into a cell or range or onto a chart

RELREF   Returns a relative reference


