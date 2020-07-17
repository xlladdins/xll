SET.NAME

Defines a name on a macro sheet to refer to a value. The defined name
exists only on the macro sheet's list of names and does not appear in
the global list of names for the workbook. The SET.NAME function is
useful for storing values while the macro is calculating.

**Syntax**

**SET.NAME**(**name\_text**, value)

Name\_text    is the name in the form of text that refers to value.

Value    is the value you want to store in name\_text.

  - > If value is omitted, the name name\_text is deleted.

  - > If value is a reference, name\_text is defined to refer to that
    > reference.

>  

**Remarks**

  - > If you want to define a name as a constant value, you can use the
    > following syntax instead of SET.NAME:

> name\_text=value
> 
> See the first two examples following.

  - > SET.NAME defines names as absolute references, even if a relative
    > reference is specified. See the third and fourth examples
    > following.

  - > If you want name\_text to refer permanently to the value of a
    > referenced cell rather than to the reference itself, you must use
    > the DEREF function. Use of DEREF prevents name\_text from
    > referring to a new value every time the contents of the referenced
    > cell changes. See the last example following.

>  

**Tips**

  - > If you need to return an array to a macro sheet (for example, if
    > the macro needs a list of all open windows), assign a name to the
    > array instead of placing the array information in a range of
    > cells. For example:

> SET.NAME("OpenDocuments", WINDOWS()) or  
> SET.NAME("OpenDocuments", {"WORKSHEET1", "WORKSHEET2"})

  - > You can then use the INDEX function with the name you have defined
    > to access items in the array stored in the name.

  - > When you're debugging a macro and want to know the current value
    > assigned to a name created by SET.NAME, you can halt the macro,
    > click Define on the Name submenu of the Insert menu, and select
    > the name from the Define Name dialog box.

>  

**Examples**

Each of these formulas defines the name Counter to refer to the constant
number 1 on the macro sheet:

SET.NAME("Counter", 1)

Counter=1

Each of these formulas redefines Counter to refer to the current value
of Counter plus 1:

SET.NAME("Counter", Counter+1)

Counter=Counter+1

The following macro formula defines the name Reference to refer to cell
$A$1:

SET.NAME("Reference", A1)

The following macro formula defines the name Results to refer to the
cells $A$1:$C$3:

SET.NAME("Results", A1:C3)

The following macro formula defines the name Range as the current
selection:

SET.NAME("Range", SELECTION())

If $A$1 contains the value 2, the following macro formula defines the
name Index to refer to the constant value 2:

SET.NAME("Index", DEREF(A1))

**Related Functions**

[DEFINE.NAME](DEFINE.NAME.md)   Defines a name on the active worksheet or macro sheet

[SET.VALUE](SET.VALUE.md)   Sets the value of a cell on a macro sheet



Return to [README.md](README.md)

SET.NAME

Defines a name on a macro sheet to refer to a value. The defined name
exists only on the macro sheet's list of names and does not appear in
the global list of names for the workbook. The SET.NAME function is
useful for storing values while the macro is calculating.

**Syntax**

**SET.NAME**(**name\_text**, value)

Name\_text    is the name in the form of text that refers to value.

Value    is the value you want to store in name\_text.

  - > If value is omitted, the name name\_text is deleted.

  - > If value is a reference, name\_text is defined to refer to that
    > reference.

>  

**Remarks**

  - > If you want to define a name as a constant value, you can use the
    > following syntax instead of SET.NAME:

> name\_text=value
> 
> See the first two examples following.

  - > SET.NAME defines names as absolute references, even if a relative
    > reference is specified. See the third and fourth examples
    > following.

  - > If you want name\_text to refer permanently to the value of a
    > referenced cell rather than to the reference itself, you must use
    > the DEREF function. Use of DEREF prevents name\_text from
    > referring to a new value every time the contents of the referenced
    > cell changes. See the last example following.

>  

**Tips**

  - > If you need to return an array to a macro sheet (for example, if
    > the macro needs a list of all open windows), assign a name to the
    > array instead of placing the array information in a range of
    > cells. For example:

> SET.NAME("OpenDocuments", WINDOWS()) or  
> SET.NAME("OpenDocuments", {"WORKSHEET1", "WORKSHEET2"})

  - > You can then use the INDEX function with the name you have defined
    > to access items in the array stored in the name.

  - > When you're debugging a macro and want to know the current value
    > assigned to a name created by SET.NAME, you can halt the macro,
    > click Define on the Name submenu of the Insert menu, and select
    > the name from the Define Name dialog box.

>  

**Examples**

Each of these formulas defines the name Counter to refer to the constant
number 1 on the macro sheet:

SET.NAME("Counter", 1)

Counter=1

Each of these formulas redefines Counter to refer to the current value
of Counter plus 1:

SET.NAME("Counter", Counter+1)

Counter=Counter+1

The following macro formula defines the name Reference to refer to cell
$A$1:

SET.NAME("Reference", A1)

The following macro formula defines the name Results to refer to the
cells $A$1:$C$3:

SET.NAME("Results", A1:C3)

The following macro formula defines the name Range as the current
selection:

SET.NAME("Range", SELECTION())

If $A$1 contains the value 2, the following macro formula defines the
name Index to refer to the constant value 2:

SET.NAME("Index", DEREF(A1))

**Related Functions**

[DEFINE.NAME](DEFINE.NAME.md)   Defines a name on the active worksheet or macro sheet

[SET.VALUE](SET.VALUE.md)   Sets the value of a cell on a macro sheet



Return to [README.md](README.md)

