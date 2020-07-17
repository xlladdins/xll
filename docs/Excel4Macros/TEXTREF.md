TEXTREF

Converts text to an absolute reference in either A1- or R1C1-style. Use
TEXTREF to convert references stored as text to references so that you
can use them with other functions, such as OFFSET.

**Syntax**

**TEXTREF**(**text**, a1)

Text    is a reference in the form of text.

A1    is a logical value specifying the reference type of text. If a1 is
TRUE, text is assumed to be an A1-style reference; if FALSE or omitted,
text is assumed to be an R1C1-style reference.

**Remarks**

  - > If you use TEXTREF by itself in a cell, you will get the value
    > contained in the cell specified by text, not the reference itself,
    > because references are automatically converted into the contents
    > of the referenced cell.

  - > If you use TEXTREF as a reference argument to a function,
    > Microsoft Excel does not convert the reference to a value.

>  

**Tip   **You can convert a reference to text with REFTEXT, manipulate
it with the REPLACE and MID functions, and convert it back to a
reference with TEXTREF.

**Examples**

TEXTREF("B7", TRUE) equals the reference value $B$7

TEXTREF("R5C5", FALSE) equals the reference value R5C5

TEXTREF("B7", FALSE) equals the \#REF\! error value, because "B7" can't
be interpreted as an R1C1-style reference.

**Related Functions**

[DEREF](DEREF.md)   Returns the values of the cells in a reference

[REFTEXT](REFTEXT.md)   Converts a reference to text



Return to [README.md](README.md)

TEXTREF

Converts text to an absolute reference in either A1- or R1C1-style. Use
TEXTREF to convert references stored as text to references so that you
can use them with other functions, such as OFFSET.

**Syntax**

**TEXTREF**(**text**, a1)

Text    is a reference in the form of text.

A1    is a logical value specifying the reference type of text. If a1 is
TRUE, text is assumed to be an A1-style reference; if FALSE or omitted,
text is assumed to be an R1C1-style reference.

**Remarks**

  - > If you use TEXTREF by itself in a cell, you will get the value
    > contained in the cell specified by text, not the reference itself,
    > because references are automatically converted into the contents
    > of the referenced cell.

  - > If you use TEXTREF as a reference argument to a function,
    > Microsoft Excel does not convert the reference to a value.

>  

**Tip   **You can convert a reference to text with REFTEXT, manipulate
it with the REPLACE and MID functions, and convert it back to a
reference with TEXTREF.

**Examples**

TEXTREF("B7", TRUE) equals the reference value $B$7

TEXTREF("R5C5", FALSE) equals the reference value R5C5

TEXTREF("B7", FALSE) equals the \#REF\! error value, because "B7" can't
be interpreted as an R1C1-style reference.

**Related Functions**

[DEREF](DEREF.md)   Returns the values of the cells in a reference

[REFTEXT](REFTEXT.md)   Converts a reference to text



Return to [README.md](README.md)

