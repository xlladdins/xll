APPLY.STYLE
===========

Equivalent to clicking the Style command on the Format menu, selecting a
style, and clicking the OK button. Applies a previously defined style to
the current selection.

**Syntax**

**APPLY.STYLE**(style\_text)

**APPLY.STYLE**?(style\_text)

Style\_text    is the name, as text, of a previously defined style. If
style\_text is not defined, APPLY.STYLE returns the \#VALUE! error value
and interrupts the macro. If style\_text is omitted, the Normal style is
applied to the selection.

**Related Functions**

DEFINE.STYLE   Defines a cell style

DELETE.STYLE   Deletes a cell style

MERGE.STYLES   Imports styles from another workbook into the active
workbook

Return to [top](#A)

APP.MAXIMIZE
