ON.RECALC

Runs a macro when a specific sheet is recalculated. Use ON.RECALC to
perform an operation on a sheet each time the sheet is recalculated,
such as checking that a certain condition is still being met.

**Syntax**

**ON.RECALC**(sheet\_text, macro\_text)

Sheet\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a sheet, given as
text. If sheet\_text is omitted, the macro is run whenever any open
sheet not specified by a previous ON.RECALC formula is recalculated.
Only one ON.RECALC formula can be run for each recalculation.

Macro\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of, or an R1C1-style
reference to, a macro you want to run when the sheet specified by
sheet\_text is recalculated. The name or reference must be in text form.
The macro will be run each time the sheet is recalculated until
ON.RECALC is cleared. If macro\_text is omitted, recalculating reverts
to its normal behavior, and any macros assigned by previous ON.RECALC
functions are turned off.

**Remarks**

A macro specified to be run by ON.RECALC is not run by actions taken by
other macros. For example, a macro specified by ON.RECALC will not be
run after the CALCULATE.NOW function is carried out, but will be run if
you change data in a sheet set to calculate automatically or choose the
Calc Now button in the Calculation panel of the Options dialog box,
which appears when you choose the Options command from the Tools menu.

**Examples**

In Microsoft Excel for Windows, the following macro formula specifies
that the macro Printer on the macro sheet AUTOREPT.XLS be run when the
worksheet named REPORT.XLS is recalculated:

ON.RECALC("REPORT.XLS", "AUTOREPT.XLS\!Printer")

In Microsoft Excel for the Macintosh, the following macro formula turns
off ON.RECALC for the workbook named SALES:

ON.RECALC("SALES")

**Related Functions**

[CALCULATE.DOCUMENT](CALCULATE.DOCUMENT.md)&nbsp;&nbsp;&nbsp;Calculates the active sheet only

[CALCULATE.NOW](CALCULATE.NOW.md)&nbsp;&nbsp;&nbsp;Calculates all open workbooks immediately

[CALCULATION](CALCULATION.md)&nbsp;&nbsp;&nbsp;Controls calculation settings

[ON.ENTRY](ON.ENTRY.md)&nbsp;&nbsp;&nbsp;Runs a macro when data is entered



Return to [README](README.md)

