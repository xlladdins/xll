SAVE.AS

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

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>File format</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1 or omitted</p>
</blockquote></td>
<td><blockquote>
<p>Normal</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>SYLK</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>WKS</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>WK1</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>CSV</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>DBF2</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>DBF3</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>DIF</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>Reserved</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>DBF4</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>Reserved</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>Reserved</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>Reserved</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>15</p>
</blockquote></td>
<td><blockquote>
<p>WK3</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>Microsoft Excel 2.x</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>17</p>
</blockquote></td>
<td><blockquote>
<p>Template</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>18</p>
</blockquote></td>
<td><blockquote>
<p>Add-in macro (For compatibility only. In Microsoft Excel version 5.0, this saves as normal.)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>19</p>
</blockquote></td>
<td><blockquote>
<p>Text (Macintosh)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>20</p>
</blockquote></td>
<td><blockquote>
<p>Text (Windows)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>21</p>
</blockquote></td>
<td><blockquote>
<p>Text (MS-DOS)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>22</p>
</blockquote></td>
<td><blockquote>
<p>CSV (Macintosh)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>23</p>
</blockquote></td>
<td><blockquote>
<p>CSV (Windows)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>24</p>
</blockquote></td>
<td><blockquote>
<p>CSV (MS-DOS)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>25</p>
</blockquote></td>
<td><blockquote>
<p>International macro</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>26</p>
</blockquote></td>
<td><blockquote>
<p>International add-in macro</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>27</p>
</blockquote></td>
<td><blockquote>
<p>Reserved</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>28</p>
</blockquote></td>
<td><blockquote>
<p>Reserved</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>29</p>
</blockquote></td>
<td><blockquote>
<p>Microsoft Excel 3.0</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>30</p>
</blockquote></td>
<td><blockquote>
<p>WK1 / FMT</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>31</p>
</blockquote></td>
<td><blockquote>
<p>WK1 / Allways</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>32</p>
</blockquote></td>
<td><blockquote>
<p>WK3 / FM3</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>33</p>
</blockquote></td>
<td><blockquote>
<p>Microsoft Excel 4.0</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>34</p>
</blockquote></td>
<td><blockquote>
<p>WQ1</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>35</p>
</blockquote></td>
<td><blockquote>
<p>Microsoft Excel 4.0 workbook</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>36</p>
</blockquote></td>
<td><blockquote>
<p>Formatted text (space delimited)</p>
</blockquote></td>
</tr>
</tbody>
</table>

The following table shows which values of type\_num apply to the six
Microsoft Excel document types.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Document Type</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Worksheet</p>
</blockquote></td>
<td><blockquote>
<p>All except 10, 12-14, 18, 25-28, 36</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Chart sheet</p>
</blockquote></td>
<td><blockquote>
<p>All except 10, 12-14, 18, 25-28</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Visual Basic module</p>
</blockquote></td>
<td><blockquote>
<p>1, 3, 17</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Dialog</p>
</blockquote></td>
<td><blockquote>
<p>1, 17</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Macro sheet</p>
</blockquote></td>
<td><blockquote>
<p>1-3, 6, 9, 16-29, 33</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Workbook</p>
</blockquote></td>
<td><blockquote>
<p>1, 15, 35</p>
</blockquote></td>
</tr>
</tbody>
</table>

Prot\_pwd    corresponds to the Protection Password box in the Save
Options dialog box in Microsoft Excel 95 or earlier versions, or the
Password To Open box in Microsoft Excel 97 or later.

  - > Prot\_pwd is a password given as text or as a reference to a cell
    > containing text. Prot\_pwd should be no more than 15 characters.

  - > If a file is saved with a password, the password must be supplied
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

  - > If read\_only\_rec is TRUE, Microsoft Excel saves the workbook as
    > a read-only recommended workbook; if FALSE, Microsoft Excel saves
    > the workbook normally; if omitted, Microsoft Excel uses the
    > current settings.

  - > When you open a workbook that was saved as read-only recommended,
    > Microsoft Excel displays a message recommending that you open the
    > workbook as read-only.

>  

**Related Functions**

[CLOSE](CLOSE.md)   Closes the active window

[GET.DOCUMENT](GET.DOCUMENT.md)   Returns information about a workbook

[SAVE](SAVE.md)   Saves the active workbook

[SAVE.WORKBOOK](SAVE.WORKBOOK.md)   Saves a workbook


