# DEREF

Returns the value of the cells in a reference.

**Syntax**

**DEREF**(**reference**)

Reference&nbsp;&nbsp;&nbsp;&nbsp;is the cell or cells from which you
want to obtain a value. If reference is the reference of a single cell,
DEREF returns the value of that cell. If reference is the reference of a
range of cells, DEREF returns the array of values in those cells. If
reference refers to the active sheet, it must be an absolute reference.
Relative references are converted to absolute references.

**Remarks**

In most formulas, there is no difference between using a value and using
the reference of a cell containing that value. The reference is
automatically converted to the value, as necessary. For example, if cell
A1 contains the value 2, then the formula =A1+1, like the formula =2+1,
returns the result 3, because the reference A1 is converted to the value
2. However, in a few functions, such as the SET.NAME function,
references are not automatically converted to values. Instead, those
functions behave differently depending on whether an argument is a
reference or a value.

**Example**

See the sixth example for SET.NAME.

**Related Function**

[SET.NAME](SET.NAME.md)&nbsp;&nbsp;&nbsp;Defines a names as a value



Return to [README](README.md#D)

# DEREF

Returns the value of the cells in a reference.

**Syntax**

**DEREF**(**reference**)

Reference&nbsp;&nbsp;&nbsp;&nbsp;is the cell or cells from which you
want to obtain a value. If reference is the reference of a single cell,
DEREF returns the value of that cell. If reference is the reference of a
range of cells, DEREF returns the array of values in those cells. If
reference refers to the active sheet, it must be an absolute reference.
Relative references are converted to absolute references.

**Remarks**

In most formulas, there is no difference between using a value and using
the reference of a cell containing that value. The reference is
automatically converted to the value, as necessary. For example, if cell
A1 contains the value 2, then the formula =A1+1, like the formula =2+1,
returns the result 3, because the reference A1 is converted to the value
2. However, in a few functions, such as the SET.NAME function,
references are not automatically converted to values. Instead, those
functions behave differently depending on whether an argument is a
reference or a value.

**Example**

See the sixth example for SET.NAME.

**Related Function**

[SET.NAME](SET.NAME.md)&nbsp;&nbsp;&nbsp;Defines a names as a value



Return to [README](README.md#D)

