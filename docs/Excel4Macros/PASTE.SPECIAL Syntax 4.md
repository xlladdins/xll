PASTE.SPECIAL Syntax 4
======================

Equivalent to clicking the Paste Special command on the Edit menu when
pasting data from another application into Microsoft Excel. Use syntax 4
when pasting information from another application.

**Syntax**

**PASTE.SPECIAL**(format\_text, pastelink\_logical,
display\_icon\_logical, icon\_file, icon\_number, icon\_label)

**PASTE.SPECIAL**?(format\_text, pastelink\_logical,
display\_icon\_logical, icon\_file, icon\_number, icon\_label)

Format\_text    is text specifying the type of data you want to paste
from the Clipboard.

-   The valid data types vary depending on the application from which
    > the data was copied. For example, if you\'re copying data from
    > Microsoft Word, some of the data types are \"Microsoft Document
    > Object\", \"Picture\", and \"Text\".

-   For more information about object classes, see your Microsoft
    > Windows or Apple system software documentation.

>  

Pastelink\_logical    is a logical value specifying whether to link the
pasted data to its source application.

-   If pastelink\_logical is TRUE, Microsoft Excel updates the pasted
    > information whenever it changes in the source application.

-   If pastelink\_logical is FALSE or omitted, the information is pasted
    > without a link.

-   If Microsoft Excel or the source application does not support
    > linking for the specified format\_text, then pastelink\_logical is
    > ignored.

Display\_icon\_logical    is a logical value that specifies whether you
want an application\'s icon to be displayed on the worksheet instead of
the linked data. Equivalent to the Display as Icon check box on the
Paste Special dialog box. If TRUE, the application\'s icon will be
displayed. If FALSE or omitted, the application\'s icon will not be
displayed.

Icon\_file    is the name of the file (with an .EXE or .DLL extension)
that contains the icon. If display\_icon\_logical is FALSE, this
argument is ignored.

Icon\_number    is the number associated with the icon and corresponds
to the icon\'s relative position within the Icon Drop Down list box on
the Change Icon Dialog box, which appears when you click the Change Icon
button in the Paste Special dialog box. If display\_icon\_logical is
FALSE, this argument is ignored.

Icon\_label    is the caption that you want to appear below the icon,
and is equivalent to the Caption text box on the Change Icon dialog box,
which appears when you click the Change Icon button in the Paste Special
dialog box. If display\_icon\_logical is FALSE, this argument is
ignored.

**Related Functions**

Syntax 1   Pasting into a sheet or macro sheet

Syntax 2   Copying from a sheet and pasting into a chart

Syntax 3   Copying and pasting between charts

Return to [top](#H)

PASTE.TOOL
