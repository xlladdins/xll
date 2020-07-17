OPEN.DIALOG

Displays the standard Microsoft Excel File Open dialog box with the
specified file filters. When the user presses the button specified by
button\_text, the filename the user picks or types in will be returned.

**Syntax**

**OPEN.DIALOG**(file\_filter, button\_text, title, filter\_index)

File\_filter    is the file filtering criteria to use, as text. For
Microsoft Excel for Windows, file\_filter consists of two parts, a
descriptive phrase denoting the file type followed by a comma and then
the MS-DOS wildcard file filter specification, as in "Text Files
(\*.TXT), \*.TXT, Add-in Files (\*.XLA), \*.XLA". Groups of filter
specifications are also separated by commas. Each separate pair is
listed in the file type drop-down list box. File\_filter can include an
asterisk (\*) to represent any sequence of characters and a question
mark (?) to represent any single character. For Microsoft Excel for the
Macintosh, file\_filter consists of file type codes separated by commas,
as in "TEXT,XLA ,XLS4". Spaces are significant and should not be
inserted before or after the comma separators unless they are part of
the file type code.

Button\_text    is the text used for the Open button in the dialog. If
omitted, "Open" will be used. Button\_text is ignored on Microsoft Excel
for Windows.

Title    specifies the dialog window title. If omitted, "File Open" will
be used as the default.

Filter\_index    specifies the index number of the default file
filtering criteria from 1 to the number of filters specified in
file\_filter. If omitted or greater than the number of filters present,
1 will be used as the starting index number. The argument is ignored on
Microsoft Excel for the Macintosh.

**Remarks**

  - > To use multiple MS-DOS wildcard expressions within file\_filter
    > for a single file filter type, separate the wildcard expressions
    > with semicolons, as in "VB Files (\*.bas; \*.txt), \*.bas;\*.txt"

  - > If file\_filter is omitted, "ALL Files (\*.\*), \*.\*" will be
    > used as the default in Microsoft Excel for Windows. The default
    > for Microsoft Excel for the Macintosh is all file types.

  - > If the user cancels the dialog box, FALSE is returned.

**Examples**

OPEN.DIALOG("Text Files (\*.TXT), \*.TXT, Add-in Files (\*.XLA),
\*.XLA,ALL FILES (\*.\*), \*.\*","FILE OPENER") opens a file open dialog
box titled "FILE OPENER" with three file filter criteria in the
drop-down list box.

**Related Function**

[SAVE.DIALOG](SAVE.DIALOG.md)   Displays the standard Microsoft Excel File Save As dialog
box and gets a file name from the user



Return to [README.md](README.md)

OPEN.DIALOG

Displays the standard Microsoft Excel File Open dialog box with the
specified file filters. When the user presses the button specified by
button\_text, the filename the user picks or types in will be returned.

**Syntax**

**OPEN.DIALOG**(file\_filter, button\_text, title, filter\_index)

File\_filter    is the file filtering criteria to use, as text. For
Microsoft Excel for Windows, file\_filter consists of two parts, a
descriptive phrase denoting the file type followed by a comma and then
the MS-DOS wildcard file filter specification, as in "Text Files
(\*.TXT), \*.TXT, Add-in Files (\*.XLA), \*.XLA". Groups of filter
specifications are also separated by commas. Each separate pair is
listed in the file type drop-down list box. File\_filter can include an
asterisk (\*) to represent any sequence of characters and a question
mark (?) to represent any single character. For Microsoft Excel for the
Macintosh, file\_filter consists of file type codes separated by commas,
as in "TEXT,XLA ,XLS4". Spaces are significant and should not be
inserted before or after the comma separators unless they are part of
the file type code.

Button\_text    is the text used for the Open button in the dialog. If
omitted, "Open" will be used. Button\_text is ignored on Microsoft Excel
for Windows.

Title    specifies the dialog window title. If omitted, "File Open" will
be used as the default.

Filter\_index    specifies the index number of the default file
filtering criteria from 1 to the number of filters specified in
file\_filter. If omitted or greater than the number of filters present,
1 will be used as the starting index number. The argument is ignored on
Microsoft Excel for the Macintosh.

**Remarks**

  - > To use multiple MS-DOS wildcard expressions within file\_filter
    > for a single file filter type, separate the wildcard expressions
    > with semicolons, as in "VB Files (\*.bas; \*.txt), \*.bas;\*.txt"

  - > If file\_filter is omitted, "ALL Files (\*.\*), \*.\*" will be
    > used as the default in Microsoft Excel for Windows. The default
    > for Microsoft Excel for the Macintosh is all file types.

  - > If the user cancels the dialog box, FALSE is returned.

**Examples**

OPEN.DIALOG("Text Files (\*.TXT), \*.TXT, Add-in Files (\*.XLA),
\*.XLA,ALL FILES (\*.\*), \*.\*","FILE OPENER") opens a file open dialog
box titled "FILE OPENER" with three file filter criteria in the
drop-down list box.

**Related Function**

[SAVE.DIALOG](SAVE.DIALOG.md)   Displays the standard Microsoft Excel File Save As dialog
box and gets a file name from the user



Return to [README.md](README.md)

