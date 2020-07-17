SAVE.AS
=======

Equivalent to clicking the Save As command on the File menu. Use SAVE.AS
to specify a new filename, file type, protection password, or
write-reservation password, or to create a backup file.

**Syntax**

**SAVE.AS**(document\_text, type\_num, prot\_pwd, backup,
write\_res\_pwd, read\_only\_rec)

**SAVE.AS**?(document\_text, type\_num, prot\_pwd, backup,
write\_res\_pwd, read\_only\_rec)

Document\_text    specifies the name of a workbook to save, such as
SALES.XLS (in Microsoft Excel for Windows) or SALES (in Microsoft Excel
for the Macintosh). You can include a full path in document\_text, such
as C:\\EXCEL\\ANALYZE.XLS (in Microsoft Excel for Windows) or
HARDDISK:FINANCIALS:ANALYZE (in Microsoft Excel for the Macintosh).

Type\_num    is a number specifying the file format in which to save the
workbook.

+-----------------+---------------------------------------------------+
| > **Type\_num** | > **File format**                                 |
+-----------------+---------------------------------------------------+
| > 1 or omitted  | > Normal                                          |
+-----------------+---------------------------------------------------+
| > 2             | > SYLK                                            |
+-----------------+---------------------------------------------------+
| > 3             | > Text                                            |
+-----------------+---------------------------------------------------+
| > 4             | > WKS                                             |
+-----------------+---------------------------------------------------+
| > 5             | > WK1                                             |
+-----------------+---------------------------------------------------+
| > 6             | > CSV                                             |
+-----------------+---------------------------------------------------+
| > 7             | > DBF2                                            |
+-----------------+---------------------------------------------------+
| > 8             | > DBF3                                            |
+-----------------+---------------------------------------------------+
| > 9             | > DIF                                             |
+-----------------+---------------------------------------------------+
| > 10            | > Reserved                                        |
+-----------------+---------------------------------------------------+
| > 11            | > DBF4                                            |
+-----------------+---------------------------------------------------+
| > 12            | > Reserved                                        |
+-----------------+---------------------------------------------------+
| > 13            | > Reserved                                        |
+-----------------+---------------------------------------------------+
| > 14            | > Reserved                                        |
+-----------------+---------------------------------------------------+
| > 15            | > WK3                                             |
+-----------------+---------------------------------------------------+
| > 16            | > Microsoft Excel 2.x                             |
+-----------------+---------------------------------------------------+
| > 17            | > Template                                        |
+-----------------+---------------------------------------------------+
| > 18            | > Add-in macro (For compatibility only. In        |
|                 | > Microsoft Excel version 5.0, this saves as      |
|                 | > normal.)                                        |
+-----------------+---------------------------------------------------+
| > 19            | > Text (Macintosh)                                |
+-----------------+---------------------------------------------------+
| > 20            | > Text (Windows)                                  |
+-----------------+---------------------------------------------------+
| > 21            | > Text (MS-DOS)                                   |
+-----------------+---------------------------------------------------+
| > 22            | > CSV (Macintosh)                                 |
+-----------------+---------------------------------------------------+
| > 23            | > CSV (Windows)                                   |
+-----------------+---------------------------------------------------+
| > 24            | > CSV (MS-DOS)                                    |
+-----------------+---------------------------------------------------+
| > 25            | > International macro                             |
+-----------------+---------------------------------------------------+
| > 26            | > International add-in macro                      |
+-----------------+---------------------------------------------------+
| > 27            | > Reserved                                        |
+-----------------+---------------------------------------------------+
| > 28            | > Reserved                                        |
+-----------------+---------------------------------------------------+
| > 29            | > Microsoft Excel 3.0                             |
+-----------------+---------------------------------------------------+
| > 30            | > WK1 / FMT                                       |
+-----------------+---------------------------------------------------+
| > 31            | > WK1 / Allways                                   |
+-----------------+---------------------------------------------------+
| > 32            | > WK3 / FM3                                       |
+-----------------+---------------------------------------------------+
| > 33            | > Microsoft Excel 4.0                             |
+-----------------+---------------------------------------------------+
| > 34            | > WQ1                                             |
+-----------------+---------------------------------------------------+
| > 35            | > Microsoft Excel 4.0 workbook                    |
+-----------------+---------------------------------------------------+
| > 36            | > Formatted text (space delimited)                |
+-----------------+---------------------------------------------------+

The following table shows which values of type\_num apply to the six
Microsoft Excel document types.

+-----------------------+---------------------------------------+
| > **Document Type**   | > **Type\_num**                       |
+-----------------------+---------------------------------------+
| > Worksheet           | > All except 10, 12-14, 18, 25-28, 36 |
+-----------------------+---------------------------------------+
| > Chart sheet         | > All except 10, 12-14, 18, 25-28     |
+-----------------------+---------------------------------------+
| > Visual Basic module | > 1, 3, 17                            |
+-----------------------+---------------------------------------+
| > Dialog              | > 1, 17                               |
+-----------------------+---------------------------------------+
| > Macro sheet         | > 1-3, 6, 9, 16-29, 33                |
+-----------------------+---------------------------------------+
| > Workbook            | > 1, 15, 35                           |
+-----------------------+---------------------------------------+

Prot\_pwd    corresponds to the Protection Password box in the Save
Options dialog box in Microsoft Excel 95 or earlier versions, or the
Password To Open box in Microsoft Excel 97 or later.

-   Prot\_pwd is a password given as text or as a reference to a cell
    > containing text. Prot\_pwd should be no more than 15 characters.

-   If a file is saved with a password, the password must be supplied
    > for the file to be opened.

>  

Backup    is a logical value corresponding to the Always Create Backup
check box in the Save Options dialog box and specifies whether to make a
backup workbook. If backup is TRUE, Microsoft Excel creates a backup
file; if FALSE, no backup file is created; if omitted, the status is
unchanged.

Write\_res\_pwd    corresponds to the Write Reservation Password box in
the Save Options dialog box in Microsoft Excel 95 or earlier versions,
or the Password To Modify box in Microsoft Excel 97 or later. Allows the
user to write to a file. If a file is saved with a password and the
password is not supplied when the file is opened, the file is opened
read-only.

Read\_only\_rec    is a logical value corresponding to the Read-Only
Recommended check box in the Save Options dialog box.

-   If read\_only\_rec is TRUE, Microsoft Excel saves the workbook as a
    > read-only recommended workbook; if FALSE, Microsoft Excel saves
    > the workbook normally; if omitted, Microsoft Excel uses the
    > current settings.

-   When you open a workbook that was saved as read-only recommended,
    > Microsoft Excel displays a message recommending that you open the
    > workbook as read-only.

>  

**Related Functions**

CLOSE   Closes the active window

GET.DOCUMENT   Returns information about a workbook

SAVE   Saves the active workbook

SAVE.WORKBOOK   Saves a workbook

Return to [top](#Q)

SAVE.COPY.AS
