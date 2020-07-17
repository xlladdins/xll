CELL.PROTECTION
===============

Equivalent to choosing the Protection tab in the Format Cells dialog
box, which appears when you choose the Cells command from the Format
menu. Allows you to control cell protection and display.

Arguments are logical values corresponding to check boxes in the
Protection tab. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box. If an
argument is omitted and the setting has been previously changed from the
defaults, the setting is not changed.

**Syntax**

**CELL.PROTECTION**(locked, hidden)

**CELL.PROTECTION**?(locked, hidden)

Locked    corresponds to the Locked check box. The default is TRUE.

Hidden    corresponds to the Hidden check box. The default is FALSE.

**Remarks**

Options selected in the Protection tab of the Format Cells dialog box or
with the CELL.PROTECTION function are activated only when the Protect
Sheet command is chosen from the Protection submenu on the Tools menu or
the PROTECT.DOCUMENT function is used to select protection.

**Related Functions**

PROTECT.DOCUMENT   Controls protection for the active sheet

SAVE.AS   Saves a workbook and allows you to specify the name, file
type, password, backup file, and location of the workbook

Return to [top](#A)

CHANGE.LINK
