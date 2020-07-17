ABSREF

Returns the absolute reference of the cells that are offset from a
reference by a specified amount. You should generally use OFFSET instead
of ABSREF. This function is provided for users who prefer to supply an
absolute reference in text form.

**Syntax**

**ABSREF**(**ref\_text**, **reference**)

Ref\_text    specifies a position relative to reference. Think of
ref\_text as "directions" from one range of cells to another.

  - > Ref\_text must be an R1C1-style relative reference in the form of
    > text, such as "R\[1\]C\[1\]".

  - > Ref\_text is considered relative to the cell in the upper-left
    > corner of reference.

>  

Reference    is a cell or range of cells specifying a starting point
that ref\_text uses to locate another range of cells. Reference can be
an external reference.

**Remarks**

  - > If you use ABSREF in a function or operation, you will usually get
    > the values contained in the reference instead of the reference
    > itself because the reference is automatically converted to the
    > contents of the reference.

  - > If you use ABSREF in a function that requires a reference
    > argument, then Microsoft Excel does not convert the reference to a
    > value.

  - > If you want to work with the actual reference, use the REFTEXT
    > function to convert the active-cell reference to text, which you
    > can then store or manipulate (or convert back to a reference with
    > TEXTREF). See the third example following.

>  

**Examples**

ABSREF("R\[-2\]C\[-2\]", C3) equals $A$1

ABSREF(RELREF(A1, C3), D4) equals $B$2

REFTEXT(ABSREF("R\[-2\]C\[-2\]:R\[2\]C\[2\]", C3:G7), TRUE) is
equivalent to

REFTEXT(ABSREF("R\[-2\]C\[-2\]:R\[2\]C\[2\]", C3), TRUE), which equals
"$A$1:$E$5"

In Microsoft Excel for Windows ABSREF("R\[-2\]C\[-2\]",
\[FINANCE.XLS\]Sheet1\!C3) equals \[FINANCE.XLS\]Sheet1\!$A$1.

In Microsoft Excel for the Macintosh ABSREF("R\[-2\]C\[-2\]",
\[FINANCE\]Sheet1\!C3) equals \[FINANCE\]Sheet1\!$A$1

**Related Function**

[RELREF](RELREF.md)   Returns a relative reference



Return to [README](README.md)

