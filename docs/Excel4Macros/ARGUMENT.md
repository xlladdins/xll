ARGUMENT

Describes the arguments used in a custom function, which is a type of
macro, or in a subroutine. A custom function or subroutine must contain
one ARGUMENT function for each argument in the macro itself. There are
two forms of the ARGUMENT function. In the first form, only name\_text
is required; in the second form, only reference is required. Use the
first form if you want to store the argument as a name. Use the second
form if you want to store the argument in a specific cell or cells.

**Syntax 1**

For name storage

**ARGUMENT**(**name\_text**, data\_type\_num)

**Syntax 2**

For cell storage

**ARGUMENT**(name\_text, data\_type\_num, **reference**)

Name\_text    is the name of the argument or of the cells containing the
argument. Name\_text is required if you omit reference.

Data\_type\_num    is a number that determines what type of values
Microsoft Excel accepts for the argument. The following table lists the
possible data types.

|                     |                   |
| ------------------- | ----------------- |
| **Data\_type\_num** | **Type of value** |
| 1                   | Number            |
| 2                   | Text              |
| 4                   | Logical           |
| 8                   | Reference         |
| 16                  | Error             |
| 64                  | Array             |

 

  - > Data\_type\_num can be a sum of the preceding different numbers to
    > allow for more than one possible type of data. For example, if
    > data\_type\_num is 7, which is the sum of 1, 2, and 4, then the
    > value can be a number, text, or logical value.

  - > Data\_type\_num is an optional argument. If you omit
    > data\_type\_num, it is assumed to be 7.

  - > If the value that is passed to the function macro is not of the
    > type specified by data\_type\_num, Microsoft Excel first attempts
    > to convert it to the specified type. If the value cannot be
    > converted, the macro returns the \#VALUE\! error value.

>  

Reference    is the cell or cells in which you want to store the
argument's value.

  - > If you specify reference, the value that is passed to ARGUMENT is
    > entered as a constant in the specified cell, and name\_text
    > becomes an optional argument because you can refer to the cell
    > with either reference or name\_text.

  - > If you omit reference, name\_text is defined on the macro sheet
    > and refers to the value that is passed to ARGUMENT. Once
    > name\_text is defined, you can use it in formulas.

>  

**Remarks**

  - > Custom functions and subroutines can accept from 1 to 29
    > arguments.

  - > If a macro contains an ARGUMENT function and you omit the
    > corresponding argument in the function that starts the macro, the
    > macro uses the \#N/A error value as the value of the argument.

>  

**Examples**

To create a custom function that calculates profit, use the following
functions to specify arguments for cost, sales, and sales volume:

ARGUMENT("UnitsSold", 1)

ARGUMENT("UnitCost", 1)

ARGUMENT("UnitPrice", 1)

**Related Functions**

[RESULT](RESULT.md)   Specifies the data type a custom function returns

[VOLATILE](VOLATILE.md)   Makes custom functions recalculate automatically


