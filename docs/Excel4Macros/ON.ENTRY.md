ON.ENTRY

Runs a macro when you enter data into any cell on the specified sheet.

**Syntax**

**ON.ENTRY**(sheet\_text, macro\_text)

Sheet\_text    is a text value specifying the name of a sheet in a
workbook. If sheet\_text is omitted, the macro is run whenever you enter
data into any sheet or macro sheet.

Macro\_text    is the name of, or an R1C1-style reference to, a macro
you want to run when you enter data into the sheet specified by
sheet\_text. The name or reference must be in text form. If macro\_text
is omitted, entering data reverts to its normal behavior, and any macros
assigned by previous ON.ENTRY functions are turned off.

**Remarks**

  - > The macro is run only when you enter data in a cell, not when you
    > use edit commands or macro functions.

  - > To determine what cell had data entered into it, use a CALLER
    > function in the macro specified by macro\_text.

>  

**Related Functions**

[ENTER.DATA](ENTER.DATA.md)   Turns Data Entry mode on or off

[ON.RECALC](ON.RECALC.md)   Runs a macro when a workbook is recalculated



Return to [README](README.md)

