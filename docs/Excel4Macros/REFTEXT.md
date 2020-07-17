REFTEXT

Converts a reference to an absolute reference in the form of text. Use
REFTEXT when you need to manipulate references with text functions.
After manipulating the reference text, you can convert it back into a
normal reference by using TEXTREF.

**Syntax**

**REFTEXT**(**reference**, a1)

Reference    is the reference you want to convert.

A1    is a logical value specifying A1-style or R1C1-style references.

  - > If a1 is TRUE, REFTEXT returns an A1-style reference.

  - > If a1 is FALSE or omitted, REFTEXT returns an R1C1-style
    > reference.

>  

**Examples**

REFTEXT(C3, TRUE) equals "$C$3"

REFTEXT(B2:F2) equals "R2C2:R2C6"

If the active cell is B9 on the active sheet named SHEET1, then:

REFTEXT(ACTIVE.CELL()) equals "\[Book1\]SHEET1\!R9C2"

REFTEXT(ACTIVE.CELL(), TRUE) equals "\[Book1\]SHEET1\!$B$9"

**Related Functions**

[ABSREF](ABSREF.md)   Returns the absolute reference of a range of cells to another
range

[DEREF](DEREF.md)   Returns the values of cells in the reference

[RELREF](RELREF.md)   Returns a relative reference

[TEXTREF](TEXTREF.md)   Converts text to a reference



Return to [README](README.md)

