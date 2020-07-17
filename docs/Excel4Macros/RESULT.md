RESULT
======

Specifies the type of data a macro or custom function returns. Use
RESULT to make sure your macros, custom functions, or subroutines return
values of the correct data type.

**Syntax**

**RESULT**(type\_num)

Type\_num    is a number specifying the data type.

+-----------------+-----------------------------+
| > **Type\_num** | > **Type of returned data** |
+-----------------+-----------------------------+
| > 1             | > Number                    |
+-----------------+-----------------------------+
| > 2             | > Text                      |
+-----------------+-----------------------------+
| > 4             | > Logical                   |
+-----------------+-----------------------------+
| > 8             | > Reference                 |
+-----------------+-----------------------------+
| > 16            | > Error                     |
+-----------------+-----------------------------+
| > 64            | > Array                     |
+-----------------+-----------------------------+

 

-   Type\_num can be the sum of the numbers in the preceding table to
    > allow for more than one possible result type. For example, if
    > type\_num is 12, which equals 4 + 8, the result can be a logical
    > or a reference value.

-   If you omit type\_num, it is assumed to be 7. Since 7 equals 1 + 2 +
    > 4, the value returned can be a number (1), text (2), or logical
    > value (4).

>  

**Examples**

The following function specifies that a custom function\'s return value
can be a number or a logical value (4+1=5):

RESULT(5)

**Related Functions**

ARGUMENT   Passes an argument to a macro

RETURN   Ends the currently running macro

Return to [top](#Q)

RESUME
