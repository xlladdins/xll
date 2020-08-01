# OPEN

Equivalent to clicking the Open command on the File menu. Opens an
existing workbook.

**Syntax**

**OPEN**(file\_text, update\_links, read\_only, format, prot\_pwd,
write\_res\_pwd, ignore\_rorec, file\_origin, custom\_delimit,
add\_logical, editable, file\_access, notify\_logical, converter)

**OPEN**?(file\_text, update\_links, read\_only, format, prot\_pwd,
write\_res\_pwd, ignore\_rorec, file\_origin, custom\_delimit,
add\_logical, editable, file\_access, notify\_logical, converter)

File\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name, as text, of the workbook
you want to open. File\_text can include a drive and path, and can be a
network pathname. In the dialog-box form in Microsoft Excel for Windows,
file\_text can include an asterisk (\*) to represent any sequence of
characters and a question mark (?) to represent any single character.

Update\_links&nbsp;&nbsp;&nbsp;&nbsp;specifies whether and how to update
external and remote references. If update\_links is omitted, Microsoft
Excel displays a message asking if you want to update links.

|                         |                                                |
| ----------------------- | ---------------------------------------------- |
| **If update\_links is** | **Then Microsoft Excel**                       |
| 0                       | Updates neither external nor remote references |
| 1                       | Updates external references only               |
| 2                       | Updates remote references only                 |
| 3                       | Updates external and remote references         |

**Note**&nbsp;&nbsp;&nbsp;When you are opening a file in WKS, WK1, or
WK3 format, the update\_links argument specifies whether Microsoft Excel
generates charts from any graphs attached to the WKS, WK1, or WK3 file.

|                         |                |
| ----------------------- | -------------- |
| **If update\_links is** | **Charts are** |
| 0                       | Not created    |
| 2                       | Created        |

Read\_only&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Read Only check box
in the Open dialog box. If read\_only is TRUE, the workbook can be
modified but changes cannot be saved; if FALSE or omitted, changes to
the workbook can be saved.

Format&nbsp;&nbsp;&nbsp;&nbsp;specifies what character to use as a
delimiter when opening text files. If format is omitted, Microsoft Excel
uses the current delimiter setting.

|                  |                             |
| ---------------- | --------------------------- |
| **If format is** | **Values are separated by** |
| 1                | Tabs                        |
| 2                | Commas                      |
| 3                | Spaces                      |
| 4                | Semicolons                  |
| 5                | Nothing                     |
| 6                | Custom characters           |

Prot\_pwd&nbsp;&nbsp;&nbsp;&nbsp;is the password, as text, required to
unprotect a protected file. If prot\_pwd is omitted and file\_text
requires a password, the Password dialog box is displayed. Passwords are
case-sensitive. Passwords are not recorded when you open a workbook and
supply the password with the macro recorder on.

Write\_res\_pwd&nbsp;&nbsp;&nbsp;&nbsp;is the password, as text,
required to open a read-only file with full write privileges. If
write\_res\_pwd is omitted and file\_text requires a password, the
Password dialog box is displayed.

Ignore\_rorec&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that controls
whether the read-only recommended message is displayed. If ignore\_rorec
is TRUE, Microsoft Excel prevents display of the message; if FALSE or
omitted, and if read\_only is also FALSE or omitted, Microsoft Excel
displays the alert when opening a read-only recommended workbook.

File\_origin&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying whether a
text file originated on the Macintosh or in Windows.

|                  |                                    |
| ---------------- | ---------------------------------- |
| **File\_origin** | **Original operating environment** |
| 1                | Macintosh                          |
| 2                | Windows (ANSI)                     |
| 3                | MS-DOS (PC-8)                      |

OmittedCurrent operating environment

Custom\_delimit&nbsp;&nbsp;&nbsp;&nbsp;is the character you want to use
as a custom delimiter when opening text files.

  - > Custom\_delimit is text or a reference or formula that returns
    > text, such as CHAR(124).

  - > Custom\_delimit is required if format is 6; it is ignored if
    > format is not 6.

  - > Only the first character in custom\_delimit is used.

Add\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether or not to add file\_text to the open workbook. If add\_logical
is TRUE, the document is added; if FALSE or omitted, it is not added.
This argument is for compatibility with workbooks from Microsoft Excel
version 4.0.

Editable&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to
opening a file (such as a template) while holding down SHIFT key. If
TRUE, editable is the equivalent to holding down the SHIFT key while
clicking the OK button in the Open dialog box. If FALSE or omitted, this
argument is ignored.

File\_access&nbsp;&nbsp;&nbsp;&nbsp; is a number specifying how the file
is to be accessed. If the file is being opened for the first time, this
argument is ignored. If the file is already opened, this argument
determines how to change the user's access permissions for the file.

|                 |                             |
| --------------- | --------------------------- |
| **File Access** | **How Accessed**            |
| 1               | Revert to saved copy        |
| 2               | Change to read/write access |
| 3               | Change to read only access  |

Notify\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether the user should be notified when the shared workbook is
available to be opened across a network. If TRUE, the user will be
notified when the workbook is available to be opened. If FALSE or
omitted, the user will not be notified when the file available to be
opened.

Converter&nbsp;&nbsp;&nbsp;&nbsp; is a number corresponding to the file
converter to use to open the file. Normally, Microsoft Excel
automatically determines which file converter to use; therefore, this
argument can usually be excluded. If you want to be certain, however,
that a specific manually installed converter be used, then include this
argument. Use GET.WORKSPACE(62) to determine which numbers corresponds
to all of the installed converters.

**Related Functions**

[CLOSE](CLOSE.md)&nbsp;&nbsp;&nbsp;Closes the active window

[FCLOSE](FCLOSE.md)&nbsp;&nbsp;&nbsp;Closes a text file

[FOPEN](FOPEN.md)&nbsp;&nbsp;&nbsp;Opens a file with the type of permission
specified

[OPEN.LINKS](OPEN.LINKS.md)&nbsp;&nbsp;&nbsp;Opens specified supporting workbooks



Return to [README](README.md#O)

# OPEN

Equivalent to clicking the Open command on the File menu. Opens an
existing workbook.

**Syntax**

**OPEN**(file\_text, update\_links, read\_only, format, prot\_pwd,
write\_res\_pwd, ignore\_rorec, file\_origin, custom\_delimit,
add\_logical, editable, file\_access, notify\_logical, converter)

**OPEN**?(file\_text, update\_links, read\_only, format, prot\_pwd,
write\_res\_pwd, ignore\_rorec, file\_origin, custom\_delimit,
add\_logical, editable, file\_access, notify\_logical, converter)

File\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name, as text, of the workbook
you want to open. File\_text can include a drive and path, and can be a
network pathname. In the dialog-box form in Microsoft Excel for Windows,
file\_text can include an asterisk (\*) to represent any sequence of
characters and a question mark (?) to represent any single character.

Update\_links&nbsp;&nbsp;&nbsp;&nbsp;specifies whether and how to update
external and remote references. If update\_links is omitted, Microsoft
Excel displays a message asking if you want to update links.

|                         |                                                |
| ----------------------- | ---------------------------------------------- |
| **If update\_links is** | **Then Microsoft Excel**                       |
| 0                       | Updates neither external nor remote references |
| 1                       | Updates external references only               |
| 2                       | Updates remote references only                 |
| 3                       | Updates external and remote references         |

**Note**&nbsp;&nbsp;&nbsp;When you are opening a file in WKS, WK1, or
WK3 format, the update\_links argument specifies whether Microsoft Excel
generates charts from any graphs attached to the WKS, WK1, or WK3 file.

|                         |                |
| ----------------------- | -------------- |
| **If update\_links is** | **Charts are** |
| 0                       | Not created    |
| 2                       | Created        |

Read\_only&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Read Only check box
in the Open dialog box. If read\_only is TRUE, the workbook can be
modified but changes cannot be saved; if FALSE or omitted, changes to
the workbook can be saved.

Format&nbsp;&nbsp;&nbsp;&nbsp;specifies what character to use as a
delimiter when opening text files. If format is omitted, Microsoft Excel
uses the current delimiter setting.

|                  |                             |
| ---------------- | --------------------------- |
| **If format is** | **Values are separated by** |
| 1                | Tabs                        |
| 2                | Commas                      |
| 3                | Spaces                      |
| 4                | Semicolons                  |
| 5                | Nothing                     |
| 6                | Custom characters           |

Prot\_pwd&nbsp;&nbsp;&nbsp;&nbsp;is the password, as text, required to
unprotect a protected file. If prot\_pwd is omitted and file\_text
requires a password, the Password dialog box is displayed. Passwords are
case-sensitive. Passwords are not recorded when you open a workbook and
supply the password with the macro recorder on.

Write\_res\_pwd&nbsp;&nbsp;&nbsp;&nbsp;is the password, as text,
required to open a read-only file with full write privileges. If
write\_res\_pwd is omitted and file\_text requires a password, the
Password dialog box is displayed.

Ignore\_rorec&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that controls
whether the read-only recommended message is displayed. If ignore\_rorec
is TRUE, Microsoft Excel prevents display of the message; if FALSE or
omitted, and if read\_only is also FALSE or omitted, Microsoft Excel
displays the alert when opening a read-only recommended workbook.

File\_origin&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying whether a
text file originated on the Macintosh or in Windows.

|                  |                                    |
| ---------------- | ---------------------------------- |
| **File\_origin** | **Original operating environment** |
| 1                | Macintosh                          |
| 2                | Windows (ANSI)                     |
| 3                | MS-DOS (PC-8)                      |

OmittedCurrent operating environment

Custom\_delimit&nbsp;&nbsp;&nbsp;&nbsp;is the character you want to use
as a custom delimiter when opening text files.

  - > Custom\_delimit is text or a reference or formula that returns
    > text, such as CHAR(124).

  - > Custom\_delimit is required if format is 6; it is ignored if
    > format is not 6.

  - > Only the first character in custom\_delimit is used.

Add\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether or not to add file\_text to the open workbook. If add\_logical
is TRUE, the document is added; if FALSE or omitted, it is not added.
This argument is for compatibility with workbooks from Microsoft Excel
version 4.0.

Editable&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to
opening a file (such as a template) while holding down SHIFT key. If
TRUE, editable is the equivalent to holding down the SHIFT key while
clicking the OK button in the Open dialog box. If FALSE or omitted, this
argument is ignored.

File\_access&nbsp;&nbsp;&nbsp;&nbsp; is a number specifying how the file
is to be accessed. If the file is being opened for the first time, this
argument is ignored. If the file is already opened, this argument
determines how to change the user's access permissions for the file.

|                 |                             |
| --------------- | --------------------------- |
| **File Access** | **How Accessed**            |
| 1               | Revert to saved copy        |
| 2               | Change to read/write access |
| 3               | Change to read only access  |

Notify\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether the user should be notified when the shared workbook is
available to be opened across a network. If TRUE, the user will be
notified when the workbook is available to be opened. If FALSE or
omitted, the user will not be notified when the file available to be
opened.

Converter&nbsp;&nbsp;&nbsp;&nbsp; is a number corresponding to the file
converter to use to open the file. Normally, Microsoft Excel
automatically determines which file converter to use; therefore, this
argument can usually be excluded. If you want to be certain, however,
that a specific manually installed converter be used, then include this
argument. Use GET.WORKSPACE(62) to determine which numbers corresponds
to all of the installed converters.

**Related Functions**

[CLOSE](CLOSE.md)&nbsp;&nbsp;&nbsp;Closes the active window

[FCLOSE](FCLOSE.md)&nbsp;&nbsp;&nbsp;Closes a text file

[FOPEN](FOPEN.md)&nbsp;&nbsp;&nbsp;Opens a file with the type of permission
specified

[OPEN.LINKS](OPEN.LINKS.md)&nbsp;&nbsp;&nbsp;Opens specified supporting workbooks



Return to [README](README.md#O)

