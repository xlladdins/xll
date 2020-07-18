# COLOR.PALETTE

Copies a color palette from an open workbook to the active workbook. Use
COLOR.PALETTE to share color palettes between workbooks.

**Syntax**

**COLOR.PALETTE**(**file\_text**)

**COLOR.PALETTE**?(file\_text)

File\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a workbook, as a text
string, that you want to copy a color palette from. The workbook
specified by file\_text must be open, or COLOR.PALETTE returns the
\#VALUE\! error value and interrupts the macro. If file\_text is empty
text (""), then COLOR.PALETTE sets colors to the default values.

**Related Function**

[EDIT.COLOR](EDIT.COLOR.md)&nbsp;&nbsp;&nbsp;Defines a color on the color palette



Return to [README](README.md)

