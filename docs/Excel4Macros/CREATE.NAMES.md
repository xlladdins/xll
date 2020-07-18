CREATE.NAMES

Equivalent to clicking the Create command on the Name submenu of the
Insert menu. Use CREATE.NAMES to quickly create names from text labels
on a sheet.

Arguments are logical values corresponding to check boxes in the Create
Names dialog box. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE or omitted, Microsoft Excel clears the check box.

**Syntax**

**CREATE.NAMES**(top, left, bottom, right)

**CREATE.NAMES**?(top, left, bottom, right)

Top**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Top Row check box.

Left**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Left Column check box.

Bottom**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Bottom Row check box.

Right**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Right Column check box.

**Remarks**

The cell containing the label text that Microsoft Excel uses to create
the names is not included in the resulting named range.

**Related Functions**

[APPLY.NAMES](APPLY.NAMES.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Replaces references and values with their
corresponding names

[DEFINE.NAME](DEFINE.NAME.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Defines a name on the active sheet or macro
sheet

[DELETE.NAME](DELETE.NAME.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Deletes a name

[FORMULA.GOTO](FORMULA.GOTO.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Selects a named area or reference on any
open workbook



Return to [README](README.md)

