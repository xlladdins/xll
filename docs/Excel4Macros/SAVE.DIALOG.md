# SAVE.DIALOG

Displays the standard Microsoft Excel File Save As dialog box and gets a
file name from the user. This function returns the path and file name of
the file that has been saved. Use SAVE.AS to automatically save a file
with a particular format and other properties.

**Syntax**

**SAVE.DIALOG**(init\_filename, title, button\_text, file\_filter,
filter\_index)

Init\_filename&nbsp;&nbsp; Specifies the suggested filename for saving.
If omitted, the active workbook's name is used, as returned by the
GET.DOCUMENT(1) function.

Title&nbsp;&nbsp;&nbsp;&nbsp;Specifies the default window title on
Microsoft Excel for Windows. For Microsoft Excel for the Macintosh,
title specifies the prompt string. If omitted, "File Save As" will be
used for Microsoft Excel for Windows, and "Save As:" For Microsoft Excel
for the Macintosh.

Button\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text used for the save button
in the dialog. If omitted, "Save" will be used as the default. This
argument is ignored on the Microsoft Excel for Windows.

File\_filter&nbsp;&nbsp;&nbsp;&nbsp;is the file filtering criteria to
use, as text. For Microsoft Excel for Windows, file\_filter consists of
two parts, a descriptive phrase denoting the file type followed by a
comma and then the MS-DOS wildcard file filter specification, as in
"Text Files (\*.TXT), \*.TXT, Add-in Files (\*.XLA), \*.XLA". Groups of
filter specifications are also separated by commas. Each separate pair
is listed in the file type drop-down list box. File\_filter can include
an asterisk (\*) to represent any sequence of characters and a question
mark (?) to represent any single character. For Microsoft Excel for the
Macintosh, file\_filter consists of file type codes separated by commas,
as in "TEXT,XLA,XLS4". Spaces are significant and should not be inserted
before or after the comma separators unless they are part of the file
type code.

Filter\_index&nbsp;&nbsp;&nbsp;&nbsp;specifies the index number of the
default file filtering criteria from 1 to the number of filters
specified in file\_filter. If omitted or greater than the number of
filters present, 1 will be used as the starting index number. The
argument is ignored on Microsoft Excel for the Macintosh.

**Remarks**

  - > To use multiple MS-DOS wildcard expressions within file\_filter
    > for a single file filter type, separate the wildcard expressions
    > with semicolons, as in "VB Files (\*.bas; \*.txt), \*.bas;\*.txt".

  - > If file\_filter is omitted, "ALL Files (\*.\*), \*.\*" will be
    > used as the default in Microsoft Excel for Windows. The default
    > for Microsoft Excel for the Macintosh is all file types.

  - > If the user cancels the dialog box, FALSE is returned.

**Examples**

SAVE.DIALOG("TRAVEL.XLS","How do you want to save this file?",,  
"Text Files (\*.TXT), \*.TXT, Add-in Files (\*.XLA), \*.XLA, ALL FILES
(\*.\*), \*.\*") opens a File Save As dialog box titled "How do you want
to save this file?", with "TRAVEL.XLS" as the suggested file name, and
with three file filter criteria in the drop-down list box.

**Related Function**

[OPEN.DIALOG](OPEN.DIALOG.md)&nbsp;&nbsp;&nbsp;Displays the standard Microsoft Excel File
[O](O.md)pen dialog box with the specified file filters



Return to [README](README.md#S)

# SAVE.DIALOG

Displays the standard Microsoft Excel File Save As dialog box and gets a
file name from the user. This function returns the path and file name of
the file that has been saved. Use SAVE.AS to automatically save a file
with a particular format and other properties.

**Syntax**

**SAVE.DIALOG**(init\_filename, title, button\_text, file\_filter,
filter\_index)

Init\_filename&nbsp;&nbsp; Specifies the suggested filename for saving.
If omitted, the active workbook's name is used, as returned by the
GET.DOCUMENT(1) function.

Title&nbsp;&nbsp;&nbsp;&nbsp;Specifies the default window title on
Microsoft Excel for Windows. For Microsoft Excel for the Macintosh,
title specifies the prompt string. If omitted, "File Save As" will be
used for Microsoft Excel for Windows, and "Save As:" For Microsoft Excel
for the Macintosh.

Button\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text used for the save button
in the dialog. If omitted, "Save" will be used as the default. This
argument is ignored on the Microsoft Excel for Windows.

File\_filter&nbsp;&nbsp;&nbsp;&nbsp;is the file filtering criteria to
use, as text. For Microsoft Excel for Windows, file\_filter consists of
two parts, a descriptive phrase denoting the file type followed by a
comma and then the MS-DOS wildcard file filter specification, as in
"Text Files (\*.TXT), \*.TXT, Add-in Files (\*.XLA), \*.XLA". Groups of
filter specifications are also separated by commas. Each separate pair
is listed in the file type drop-down list box. File\_filter can include
an asterisk (\*) to represent any sequence of characters and a question
mark (?) to represent any single character. For Microsoft Excel for the
Macintosh, file\_filter consists of file type codes separated by commas,
as in "TEXT,XLA,XLS4". Spaces are significant and should not be inserted
before or after the comma separators unless they are part of the file
type code.

Filter\_index&nbsp;&nbsp;&nbsp;&nbsp;specifies the index number of the
default file filtering criteria from 1 to the number of filters
specified in file\_filter. If omitted or greater than the number of
filters present, 1 will be used as the starting index number. The
argument is ignored on Microsoft Excel for the Macintosh.

**Remarks**

  - > To use multiple MS-DOS wildcard expressions within file\_filter
    > for a single file filter type, separate the wildcard expressions
    > with semicolons, as in "VB Files (\*.bas; \*.txt), \*.bas;\*.txt".

  - > If file\_filter is omitted, "ALL Files (\*.\*), \*.\*" will be
    > used as the default in Microsoft Excel for Windows. The default
    > for Microsoft Excel for the Macintosh is all file types.

  - > If the user cancels the dialog box, FALSE is returned.

**Examples**

SAVE.DIALOG("TRAVEL.XLS","How do you want to save this file?",,  
"Text Files (\*.TXT), \*.TXT, Add-in Files (\*.XLA), \*.XLA, ALL FILES
(\*.\*), \*.\*") opens a File Save As dialog box titled "How do you want
to save this file?", with "TRAVEL.XLS" as the suggested file name, and
with three file filter criteria in the drop-down list box.

**Related Function**

[OPEN.DIALOG](OPEN.DIALOG.md)&nbsp;&nbsp;&nbsp;Displays the standard Microsoft Excel File
[O](O.md)pen dialog box with the specified file filters



Return to [README](README.md#S)

# SAVE.DIALOG

Displays the standard Microsoft Excel File Save As dialog box and gets a
file name from the user. This function returns the path and file name of
the file that has been saved. Use SAVE.AS to automatically save a file
with a particular format and other properties.

**Syntax**

**SAVE.DIALOG**(init\_filename, title, button\_text, file\_filter,
filter\_index)

Init\_filename&nbsp;&nbsp; Specifies the suggested filename for saving.
If omitted, the active workbook's name is used, as returned by the
GET.DOCUMENT(1) function.

Title&nbsp;&nbsp;&nbsp;&nbsp;Specifies the default window title on
Microsoft Excel for Windows. For Microsoft Excel for the Macintosh,
title specifies the prompt string. If omitted, "File Save As" will be
used for Microsoft Excel for Windows, and "Save As:" For Microsoft Excel
for the Macintosh.

Button\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text used for the save button
in the dialog. If omitted, "Save" will be used as the default. This
argument is ignored on the Microsoft Excel for Windows.

File\_filter&nbsp;&nbsp;&nbsp;&nbsp;is the file filtering criteria to
use, as text. For Microsoft Excel for Windows, file\_filter consists of
two parts, a descriptive phrase denoting the file type followed by a
comma and then the MS-DOS wildcard file filter specification, as in
"Text Files (\*.TXT), \*.TXT, Add-in Files (\*.XLA), \*.XLA". Groups of
filter specifications are also separated by commas. Each separate pair
is listed in the file type drop-down list box. File\_filter can include
an asterisk (\*) to represent any sequence of characters and a question
mark (?) to represent any single character. For Microsoft Excel for the
Macintosh, file\_filter consists of file type codes separated by commas,
as in "TEXT,XLA,XLS4". Spaces are significant and should not be inserted
before or after the comma separators unless they are part of the file
type code.

Filter\_index&nbsp;&nbsp;&nbsp;&nbsp;specifies the index number of the
default file filtering criteria from 1 to the number of filters
specified in file\_filter. If omitted or greater than the number of
filters present, 1 will be used as the starting index number. The
argument is ignored on Microsoft Excel for the Macintosh.

**Remarks**

  - > To use multiple MS-DOS wildcard expressions within file\_filter
    > for a single file filter type, separate the wildcard expressions
    > with semicolons, as in "VB Files (\*.bas; \*.txt), \*.bas;\*.txt".

  - > If file\_filter is omitted, "ALL Files (\*.\*), \*.\*" will be
    > used as the default in Microsoft Excel for Windows. The default
    > for Microsoft Excel for the Macintosh is all file types.

  - > If the user cancels the dialog box, FALSE is returned.

**Examples**

SAVE.DIALOG("TRAVEL.XLS","How do you want to save this file?",,  
"Text Files (\*.TXT), \*.TXT, Add-in Files (\*.XLA), \*.XLA, ALL FILES
(\*.\*), \*.\*") opens a File Save As dialog box titled "How do you want
to save this file?", with "TRAVEL.XLS" as the suggested file name, and
with three file filter criteria in the drop-down list box.

**Related Function**

[OPEN.DIALOG](OPEN.DIALOG.md)&nbsp;&nbsp;&nbsp;Displays the standard Microsoft Excel File
[O](O.md)pen dialog box with the specified file filters



Return to [README](README.md#S)

