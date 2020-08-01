# ACTIVE.CELL

Returns the reference of the active cell in the selection as an external
reference.

**Syntax**

**ACTIVE.CELL**( )

**Remarks**

  - > If an object is selected, ACTIVE.CELL returns the \#N/A error
    > value.

  - > If a chart sheet is active, ACTIVE.CELL returns a zero value.

  - > If you use ACTIVE.CELL in a function or operation, you will
    > usually get the value contained in the active cell instead of its
    > reference, because the reference is automatically converted to the
    > contents of the reference. See the third example following.

  - > If you use ACTIVE.CELL in a function that requires a reference
    > argument, then Microsoft Excel does not convert the reference to a
    > value.

  - > If you want to work with the actual reference, use the REFTEXT XLM
    > function to convert the active-cell reference to text, which you
    > can then store or manipulate (or convert back to a reference with
    > TEXTREF). See the second example following.

**Tip**&nbsp;&nbsp;&nbsp;Use the following macro formula to verify that
the current selection is a cell or range of cells:

\=ISREF(ACTIVE.CELL( ))

**Examples**

The following macro formula assigns the name Sales to the active cell:

SET.NAME("Sales", ACTIVE.CELL())

In this example, note that "Sales" refers to a cell on the active
worksheet, but the name itself exists only in the macro sheet's list of
names. In other words, the preceding formula does not define a name on
the worksheet or in the workbook's global list of names.

The following macro formula puts the reference of the active cell into
the cell named Temp:

FORMULA("="\&REFTEXT(ACTIVE.CELL()), Temp)

The following macro formula checks the contents of the active cell. If
the cell contains only the letter "c" or "s", the macro branches to an
area named FinishRefresh:

IF(OR(ACTIVE.CELL()="c", ACTIVE.CELL()="s"), GOTO(FinishRefresh))

In Microsoft Excel for Windows, if the sheet in the active window is
named SALES and A1 is the active cell, then:

ACTIVE.CELL() equals SALES\!$A$1

In Microsoft Excel for the Macintosh, if the sheet in the active window
is named SALES 1 and A1 is the active cell, then:

ACTIVE.CELL() equals 'SALES&nbsp;1'\!$A$1

**Related Function**

[SELECT](SELECT.md)&nbsp;&nbsp;&nbsp;Selects a cell, worksheet object, or chart item



Return to [README](README.md#A)

# ACTIVE.CELL

Returns the reference of the active cell in the selection as an external
reference.

**Syntax**

**ACTIVE.CELL**( )

**Remarks**

  - > If an object is selected, ACTIVE.CELL returns the \#N/A error
    > value.

  - > If a chart sheet is active, ACTIVE.CELL returns a zero value.

  - > If you use ACTIVE.CELL in a function or operation, you will
    > usually get the value contained in the active cell instead of its
    > reference, because the reference is automatically converted to the
    > contents of the reference. See the third example following.

  - > If you use ACTIVE.CELL in a function that requires a reference
    > argument, then Microsoft Excel does not convert the reference to a
    > value.

  - > If you want to work with the actual reference, use the REFTEXT XLM
    > function to convert the active-cell reference to text, which you
    > can then store or manipulate (or convert back to a reference with
    > TEXTREF). See the second example following.

**Tip**&nbsp;&nbsp;&nbsp;Use the following macro formula to verify that
the current selection is a cell or range of cells:

\=ISREF(ACTIVE.CELL( ))

**Examples**

The following macro formula assigns the name Sales to the active cell:

SET.NAME("Sales", ACTIVE.CELL())

In this example, note that "Sales" refers to a cell on the active
worksheet, but the name itself exists only in the macro sheet's list of
names. In other words, the preceding formula does not define a name on
the worksheet or in the workbook's global list of names.

The following macro formula puts the reference of the active cell into
the cell named Temp:

FORMULA("="\&REFTEXT(ACTIVE.CELL()), Temp)

The following macro formula checks the contents of the active cell. If
the cell contains only the letter "c" or "s", the macro branches to an
area named FinishRefresh:

IF(OR(ACTIVE.CELL()="c", ACTIVE.CELL()="s"), GOTO(FinishRefresh))

In Microsoft Excel for Windows, if the sheet in the active window is
named SALES and A1 is the active cell, then:

ACTIVE.CELL() equals SALES\!$A$1

In Microsoft Excel for the Macintosh, if the sheet in the active window
is named SALES 1 and A1 is the active cell, then:

ACTIVE.CELL() equals 'SALES&nbsp;1'\!$A$1

**Related Function**

[SELECT](SELECT.md)&nbsp;&nbsp;&nbsp;Selects a cell, worksheet object, or chart item



Return to [README](README.md#A)

