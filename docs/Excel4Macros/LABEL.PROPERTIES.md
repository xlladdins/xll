LABEL.PROPERTIES

Sets the accelerator property of the label and group box controls on a
worksheet or dialog sheet.

**Syntax**

**LABEL.PROPERTIES**(accel\_text, accel\_text2, 3d\_shading)

**LABEL.PROPERTIES**?(accel\_text, accel\_text2, 3d\_shading)

Accel\_text&nbsp;&nbsp;&nbsp;&nbsp;is a text string containing the
character to use as the label's accelerator key on a dialog sheet. The
character is matched against the text of the control, and the first
matching character is underlined. When the user presses ALT+accel\_text
in Microsoft Excel for Windows or COMMAND+accel\_text in Microsoft Excel
for the Macintosh, the control is clicked. This argument is ignored for
controls on worksheets.

Accel\_text2&nbsp;&nbsp;&nbsp;&nbsp;is a text string containing the
second accelerator key on a dialog sheet. This argument is for only Far
East versions of Microsoft Excel.

3d\_shading&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether the list box appears as 3-D. If TRUE, the list box will appear
as 3-D. If FALSE or omitted, the list box will not be 3-D. This argument
is available for only worksheets.

**Related Functions**

[CHECKBOX.PROPERTIES](CHECKBOX.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Sets various properties of check
box and option box controls

[SCROLLBAR.PROPERTIES](SCROLLBAR.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Sets the properties of the scroll
bar and spinner controls

[PUSHBUTTON.PROPERTIES](PUSHBUTTON.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Sets the properties of the push
button control



Return to [README](README.md)

