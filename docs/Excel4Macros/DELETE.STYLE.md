DELETE.STYLE
============

Equivalent to choosing the Delete button from the Style dialog box,
which appears when you choose the Style command from the Format menu.
Deletes a style from a workbook. Cells formatted with the deleted style
revert to the Normal style.

**Syntax**

**DELETE.STYLE**(style\_text)

Style\_text    is the name of a style to be deleted. If style\_text does
not exist, DELETE.STYLE returns the \#VALUE! error value and interrupts
the macro.

**Remarks**

You can only delete styles from the active workbook. External references
are not permitted as part of the style\_text argument.

**Related Functions**

APPLY.STYLE   Applies a style to the selection

DEFINE.STYLE   Creates or changes a cell style

MERGE.STYLES   Merges styles from another workbook into the active
workbook

Return to [top](#A)

DELETE.TOOL
