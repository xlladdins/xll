VOLATILE

Specifies whether a custom worksheet function is volatile or
nonvolatile. A volatile custom function is recalculated every time a
calculation occurs on the worksheet.

**Syntax**

**VOLATILE**(logical)

Logical    is a logical value specifying whether the custom function is
volatile or nonvolatile. If logical is TRUE or omitted, the function is
volatile; if FALSE, nonvolatile.

**Remarks**

  - > VOLATILE must precede every other formula in the custom function
    > except RESULT and ARGUMENT.

  - > Normally, a worksheet recalculates a cell containing a nonvolatile
    > custom function only when any part of the complete formula in the
    > cell is recalculated. Use VOLATILE(TRUE) to recalculate the
    > function every time the worksheet is recalculated.

  - > Most custom functions are nonvolatile by default, but custom
    > functions with reference arguments are volatile by default. Use
    > VOLATILE(FALSE) to prevent these functions from being recalculated
    > unnecessarily often.

>  

**Related Function**

RESULT   Specifies the data type a custom function returns


