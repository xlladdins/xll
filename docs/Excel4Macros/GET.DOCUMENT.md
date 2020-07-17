GET.DOCUMENT
============

Returns information about a sheet in a workbook.

**Syntax**

**GET.DOCUMENT**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of information you
want. The following lists show the possible values of type\_num and the
corresponding results.

+---------------+-----------------------------------------------------+
| **Type\_num** | **Returns**                                         |
+---------------+-----------------------------------------------------+
| 1             | Returns the name of the workbook and worksheet as   |
|               | text. If there is only one sheet in the workbook    |
|               | and the sheet name is the same as the workbook name |
|               | less any extension, returns the name of the book.   |
|               | The book name does not include the drive, directory |
|               | or folder, or window number. Otherwise, returns the |
|               | book and sheet name in the form                     |
|               | \"\[BOOK1.XLS\]Sheet1\". It is usually best to use  |
|               | GET.DOCUMENT(76) and GET.DOCUMENT(88) to return the |
|               | name of the active worksheet and the active         |
|               | workbook.                                           |
+---------------+-----------------------------------------------------+
| 2             | Path of the directory or folder containing          |
|               | name\_text, as text. If the workbook name\_text     |
|               | hasn\'t been saved yet, returns the \#N/A error     |
|               | value.                                              |
+---------------+-----------------------------------------------------+
| 3             | Number indicating the type of sheet. If name\_text  |
|               | is a sheet, then the return value is one of the     |
|               | following numbers. If name\_text is a book, then    |
|               | the return value is always 5. If name\_text is      |
|               | omitted, then the sheet type is returned. If the    |
|               | book has one sheet that is named the same as the    |
|               | book, then the sheet type is returned.              |
|               |                                                     |
|               | 1 = Worksheet                                       |
|               |                                                     |
|               | 2 = Chart                                           |
|               |                                                     |
|               | 3 = Macro sheet                                     |
|               |                                                     |
|               | 4 = Info window if active                           |
|               |                                                     |
|               | 5 = Reserved                                        |
|               |                                                     |
|               | 6 = Module                                          |
|               |                                                     |
|               | 7 = Dialog                                          |
+---------------+-----------------------------------------------------+
| 4             | If changes have been made to the sheet since it was |
|               | last saved, returns TRUE; otherwise, returns FALSE. |
+---------------+-----------------------------------------------------+
| 5             | If the sheet is read-only, returns TRUE; otherwise, |
|               | returns FALSE.                                      |
+---------------+-----------------------------------------------------+
| 6             | If the sheet is password protected, returns TRUE;   |
|               | otherwise, returns FALSE.                           |
+---------------+-----------------------------------------------------+
| 7             | If cells in a sheet, the contents of a sheet, or    |
|               | the series in a chart are protected, returns TRUE;  |
|               | otherwise, returns FALSE.                           |
+---------------+-----------------------------------------------------+
| 8             | If the workbook windows are protected, returns      |
|               | TRUE; otherwise, returns FALSE.                     |
+---------------+-----------------------------------------------------+

The next four values of type\_num apply only to charts.

+---------------+-----------------------------------------------------+
| **Type\_num** | **Returns**                                         |
+---------------+-----------------------------------------------------+
| 9             | Number indicating the type of the main chart:       |
|               |                                                     |
|               | 1 = Area                                            |
|               |                                                     |
|               | 2 = Bar                                             |
|               |                                                     |
|               | 3 = Column                                          |
|               |                                                     |
|               | 4 = Line                                            |
|               |                                                     |
|               | 5 = Pie                                             |
|               |                                                     |
|               | 6 = XY (scatter)                                    |
|               |                                                     |
|               | 7 = 3-D area                                        |
|               |                                                     |
|               | 8 = 3-D column                                      |
|               |                                                     |
|               | 9 = 3-D line                                        |
|               |                                                     |
|               | 10 = 3-D pie                                        |
|               |                                                     |
|               | 11 = Radar                                          |
|               |                                                     |
|               | 12 = 3-D bar                                        |
|               |                                                     |
|               | 13 = 3-D surface                                    |
|               |                                                     |
|               | 14 = Doughnut                                       |
+---------------+-----------------------------------------------------+
| 10            | Number indicating the type of the overlay chart.    |
|               | Same as 1, 2, 3, 4, 5, 6, 11, and 14 for main chart |
|               | above. If there is no overlay chart, returns the    |
|               | \#N/A error value.                                  |
+---------------+-----------------------------------------------------+
| 11            | Number of series in the main chart.                 |
+---------------+-----------------------------------------------------+
| 12            | Number of series in the overlay chart.              |
+---------------+-----------------------------------------------------+

The next values of type\_num apply to worksheets and macro sheets and to
charts when appropriate.

+---------------+-----------------------------------------------------+
| **Type\_num** | **Returns**                                         |
+---------------+-----------------------------------------------------+
| 9             | Number of the first used row. If the sheet is       |
|               | empty, returns 0.                                   |
+---------------+-----------------------------------------------------+
| 10            | Number of the last used row. If the sheet is empty, |
|               | returns 0.                                          |
+---------------+-----------------------------------------------------+
| 11            | Number of the first used column. If the sheet is    |
|               | empty, returns 0.                                   |
+---------------+-----------------------------------------------------+
| 12            | Number of the last used column. If the sheet is     |
|               | empty, returns 0.                                   |
+---------------+-----------------------------------------------------+
| 13            | Number of windows.                                  |
+---------------+-----------------------------------------------------+
| 14            | Number indicating calculation mode:                 |
|               |                                                     |
|               | 1 = Automatic                                       |
|               |                                                     |
|               | 2 = Automatic except tables                         |
|               |                                                     |
|               | 3 = Manual                                          |
+---------------+-----------------------------------------------------+
| 15            | If the Iteration check box is selected in the       |
|               | Calculation tab of the Options dialog box, returns  |
|               | TRUE; otherwise, returns FALSE.                     |
+---------------+-----------------------------------------------------+
| 16            | Maximum number of iterations.                       |
+---------------+-----------------------------------------------------+
| 17            | Maximum change between iterations.                  |
+---------------+-----------------------------------------------------+
| 18            | If the Update Remote References check box is        |
|               | selected in the Calculation tab of the Options      |
|               | dialog box, returns TRUE; otherwise, returns FALSE. |
+---------------+-----------------------------------------------------+
| 19            | If the Precision As Displayed check box is selected |
|               | in the Calculation tab of the Options dialog box,   |
|               | returns TRUE; otherwise, returns FALSE.             |
+---------------+-----------------------------------------------------+
| 20            | If the 1904 Date System check box is selected in    |
|               | the Calculation tab of the Options dialog box,      |
|               | returns TRUE; otherwise, returns FALSE.             |
+---------------+-----------------------------------------------------+

Type\_num values of 21 through 29 correspond to the four default fonts
in previous versions of Microsoft Excel. These values are provided only
for macro compatibility.

The next values of type\_num apply to worksheets and macro sheets, and
to charts if indicated.

+---------------+-----------------------------------------------------+
| **Type\_num** | **Returns**                                         |
+---------------+-----------------------------------------------------+
| 30            | Horizontal array of consolidation references for    |
|               | the current sheet, in the form of text. If the list |
|               | is empty, returns the \#N/A error value.            |
+---------------+-----------------------------------------------------+
| 31            | Number from 1 to 11, indicating the function used   |
|               | in the current consolidation. The function that     |
|               | corresponds to each number is listed under the      |
|               | CONSOLIDATE function. The default function is SUM.  |
+---------------+-----------------------------------------------------+
| 32            | Three-item horizontal array indicating the status   |
|               | of the check boxes in the Data Consolidate dialog   |
|               | box. An item is TRUE if the check box is selected   |
|               | or FALSE if the check box is cleared. The first     |
|               | item indicates the Top Row check box, the second    |
|               | the Left Column check box, and the third the Create |
|               | Links To Source Data check box.                     |
+---------------+-----------------------------------------------------+
| 33            | If the Recalculate Before Save check box is         |
|               | selected in the Calculation tab of the Options      |
|               | dialog box, returns TRUE; otherwise, returns FALSE. |
+---------------+-----------------------------------------------------+
| 34            | If the workbook is read-only recommended, returns   |
|               | TRUE; otherwise, returns FALSE.                     |
+---------------+-----------------------------------------------------+
| 35            | If the workbook is write-reserved, returns TRUE;    |
|               | otherwise, returns FALSE.                           |
+---------------+-----------------------------------------------------+
| 36            | If the workbook has a write-reservation password    |
|               | and it is opened with read/write permission,        |
|               | returns the name of the user who originally saved   |
|               | the file with the write-reservation password. If    |
|               | the file is opened as read-only, or if a password   |
|               | has not been added to the workbook, returns the     |
|               | name of the current user.                           |
+---------------+-----------------------------------------------------+
| 37            | Number corresponding to the file type of the        |
|               | workbook as displayed in the Save As dialog box.    |
|               | See the SAVE.AS function for a list of all the file |
|               | types that Microsoft Excel recognizes.              |
+---------------+-----------------------------------------------------+
| 38            | If the Summary Rows Below Detail check box is       |
|               | selected in the Outline dialog box, returns TRUE;   |
|               | otherwise, returns FALSE.                           |
+---------------+-----------------------------------------------------+
| 39            | If the Summary Columns To Right Of Detail check box |
|               | is selected in the Outline dialog box, returns      |
|               | TRUE; otherwise, returns FALSE.                     |
+---------------+-----------------------------------------------------+
| 40            | If the Always Create Backup check box is selected   |
|               | in the Save Options dialog box, returns TRUE;       |
|               | otherwise, returns FALSE.                           |
+---------------+-----------------------------------------------------+
| 41            | Number from 1 to 3 indicating whether objects are   |
|               | displayed:                                          |
|               |                                                     |
|               | 1 = All objects are displayed                       |
|               |                                                     |
|               | 2 = Placeholders for pictures and charts            |
|               |                                                     |
|               | 3 = All objects are hidden                          |
+---------------+-----------------------------------------------------+
| 42            | Horizontal array of all objects in the sheet. If    |
|               | there are no objects, returns the \#N/A error       |
|               | value.                                              |
+---------------+-----------------------------------------------------+
| 43            | If the Save External Link Values check box is       |
|               | selected in the Calculation tab of the Options      |
|               | dialog box, returns TRUE; otherwise, returns FALSE. |
+---------------+-----------------------------------------------------+
| 44            | If objects in a workbook are protected, returns     |
|               | TRUE; otherwise, returns FALSE.                     |
+---------------+-----------------------------------------------------+
| 45            | A number from 0 to 3 indicating how windows are     |
|               | synchronized:                                       |
|               |                                                     |
|               | 0 = Not synchronized                                |
|               |                                                     |
|               | 1 = Synchronized horizontally                       |
|               |                                                     |
|               | 2 = Synchronized vertically                         |
|               |                                                     |
|               | 3 = Synchronized horizontally and vertically        |
+---------------+-----------------------------------------------------+
| 46            | A seven-item horizontal array of print settings     |
|               | that can be set by the LINE.PRINT macro function:   |
|               |                                                     |
|               | Setup text                                          |
|               |                                                     |
|               | Left margin                                         |
|               |                                                     |
|               | Right margin                                        |
|               |                                                     |
|               | Top margin                                          |
|               |                                                     |
|               | Bottom margin                                       |
|               |                                                     |
|               | Page length                                         |
|               |                                                     |
|               | A logical value indicating whether output will be   |
|               | formatted (TRUE) or unformatted (FALSE) when        |
|               | printed                                             |
+---------------+-----------------------------------------------------+
| 47            | If the Transition Formula Evaluation check box is   |
|               | selected in the Transition tab of the Options       |
|               | dialog box, returns TRUE; otherwise, returns FALSE. |
+---------------+-----------------------------------------------------+
| 48            | The standard column width setting.                  |
+---------------+-----------------------------------------------------+

The next values of type\_num correspond to printing and page settings.

+---------------+-----------------------------------------------------+
| **Type\_num** | **Returns**                                         |
+---------------+-----------------------------------------------------+
| 49            | The starting page number, or the \#N/A error value  |
|               | if none is specified or if \"Auto\" is entered in   |
|               | the First page Number text box on the Page tab of   |
|               | the Page Setup dialog box.                          |
+---------------+-----------------------------------------------------+
| 50            | The total number of pages that would be printed     |
|               | based on current settings, excluding notes, or 1 if |
|               | the document is a chart.                            |
+---------------+-----------------------------------------------------+
| 51            | The total number of pages that would be printed if  |
|               | you print only notes, or the \#N/A error value if   |
|               | the document is a chart.                            |
+---------------+-----------------------------------------------------+
| 52            | Four-item horizontal array indicating the margin    |
|               | settings (left, right, top, bottom) in the          |
|               | currently specified units.                          |
+---------------+-----------------------------------------------------+
| 53            | A number indicating the orientation:                |
|               |                                                     |
|               | 1 = Portrait                                        |
|               |                                                     |
|               | 2 = Landscape                                       |
+---------------+-----------------------------------------------------+
| 54            | The header as a text string, including formatting   |
|               | codes.                                              |
+---------------+-----------------------------------------------------+
| 55            | The footer as a text string, including formatting   |
|               | codes.                                              |
+---------------+-----------------------------------------------------+
| 56            | Horizontal array of two logical values              |
|               | corresponding to horizontal and vertical centering. |
+---------------+-----------------------------------------------------+
| 57            | If row or column headings are to be printed,        |
|               | returns TRUE; otherwise, returns FALSE.             |
+---------------+-----------------------------------------------------+
| 58            | If gridlines are to be printed, returns TRUE;       |
|               | otherwise, returns FALSE.                           |
+---------------+-----------------------------------------------------+
| 59            | If the sheet is printed in black and white only,    |
|               | returns TRUE; otherwise, returns FALSE.             |
+---------------+-----------------------------------------------------+
| 60            | A number from 1 to 3 indicating how the chart will  |
|               | be sized when it\'s printed:                        |
|               |                                                     |
|               | 1 = Size on screen                                  |
|               |                                                     |
|               | 2 = Scale to fit page                               |
|               |                                                     |
|               | 3 = Use full page                                   |
+---------------+-----------------------------------------------------+
| 61            | A number indicating the pagination order:           |
|               |                                                     |
|               | 1 = Down, then over                                 |
|               |                                                     |
|               | 2 = Over, then down                                 |
|               |                                                     |
|               | Returns the \#N/A error value if the document is a  |
|               | chart.                                              |
+---------------+-----------------------------------------------------+
| 62            | Percentage of reduction or enlargement, or 100% if  |
|               | none is specified. Returns the \#N/A error value if |
|               | not supported by the current printer or if the      |
|               | document is a chart.                                |
+---------------+-----------------------------------------------------+
| 63            | A two-item horizontal array indicating the number   |
|               | of pages to which the printout should be scaled to  |
|               | fit, with the first item equal to the width (or     |
|               | \#N/A if no width scaling is specified) and the     |
|               | second item equal to the height (or \#N/A if no     |
|               | height scaling is specified). \#N/A is also         |
|               | returned if the document is a chart.                |
+---------------+-----------------------------------------------------+
| 64            | An array of row numbers corresponding to rows that  |
|               | are immediately below a manual or automatic page    |
|               | break.                                              |
+---------------+-----------------------------------------------------+
| 65            | An array of column numbers corresponding to columns |
|               | that are immediately to the right of a manual or    |
|               | automatic page break.                               |
+---------------+-----------------------------------------------------+

**Note   **GET.DOCUMENT(62) and GET.DOCUMENT(63) are mutually exclusive.
If one returns a value, then the other returns the \#N/A error value.

The next values of type\_num correspond to various workbook settings.

+---------------+-----------------------------------------------------+
| **Type\_num** | **Returns**                                         |
+---------------+-----------------------------------------------------+
| 66            | In Microsoft Excel for Windows, if the Transition   |
|               | Formula Entry check box is selected in the          |
|               | Transition tab of the Options dialog box, returns   |
|               | TRUE; otherwise, returns FALSE.                     |
+---------------+-----------------------------------------------------+
| 67            | Microsoft Excel version 5.0 or later always returns |
|               | TRUE here.                                          |
+---------------+-----------------------------------------------------+
| 68            | Microsoft Excel version 5.0 or later always returns |
|               | the book name.                                      |
+---------------+-----------------------------------------------------+
| 69            | Returns TRUE if Page Breaks is chosen in the View   |
|               | tab of the Options dialog box; otherwise, returns   |
|               | FALSE.                                              |
+---------------+-----------------------------------------------------+
| 70            | Returns the names of all PivotTable reports in the  |
|               | current sheet as a horizontal array.                |
+---------------+-----------------------------------------------------+
| 71            | Returns an horizontal array of all the styles in a  |
|               | workbook.                                           |
+---------------+-----------------------------------------------------+
| 72            | Returns an horizontal array of all chart types      |
|               | displayed on the current sheet.                     |
+---------------+-----------------------------------------------------+
| 73            | Returns an array of the number of series in each    |
|               | chart of the current sheet.                         |
+---------------+-----------------------------------------------------+
| 74            | Returns the object ID of the control that currently |
|               | has the focus on a running user-defined dialog      |
|               | (based on the dialog sheet).                        |
+---------------+-----------------------------------------------------+
| 75            | Returns the object ID of the object that is the     |
|               | current default button on a running user-defined    |
|               | dialog (based on the dialog sheet).                 |
+---------------+-----------------------------------------------------+
| 76            | Returns the name of the active sheet or macro sheet |
|               | in the form \[Book1\]Sheet1.                        |
+---------------+-----------------------------------------------------+
| 77            | In Microsoft Excel for Windows, returns the paper   |
|               | size, as integer:                                   |
|               |                                                     |
|               | 1 = Letter 8.5 x 11 in                              |
|               |                                                     |
|               | 2 = Letter Small 8.5 x 11 in                        |
|               |                                                     |
|               | 5 = Legal 8.5 x 14 in                               |
|               |                                                     |
|               | 9 = A4 210 x 297 mm                                 |
|               |                                                     |
|               | 10 = A4 Small 210 x 297 mm                          |
|               |                                                     |
|               | 13 = B5 182 x 257 mm                                |
|               |                                                     |
|               | 18 = Note 8.5 x 11 in                               |
+---------------+-----------------------------------------------------+
| 78            | Returns the print resolution, as a horizontal array |
|               | of two numbers.                                     |
+---------------+-----------------------------------------------------+
| 79            | Returns TRUE if the Draft Quality check box has     |
|               | been selected from the sheet tab in the Page Setup  |
|               | dialog box; otherwise, returns FALSE.               |
+---------------+-----------------------------------------------------+
| 80            | Returns TRUE if the Comments checkbox has been      |
|               | selected on the Sheet tab in the Page Setup dialog  |
|               | box; otherwise, returns FALSE.                      |
+---------------+-----------------------------------------------------+
| 81            | Returns the Print Area from the Sheet tab of the    |
|               | Page Setup dialog box as a cell reference.          |
+---------------+-----------------------------------------------------+
| 82            | Returns the Print Titles from the Sheet tab of the  |
|               | Page Setup dialog box as an array of cell           |
|               | references.                                         |
+---------------+-----------------------------------------------------+
| 83            | Returns TRUE if the worksheet is protected for      |
|               | scenarios; otherwise, returns FALSE.                |
+---------------+-----------------------------------------------------+
| 84            | Returns the value of the first circular reference   |
|               | on the sheet, or \#N/A if there are no circular     |
|               | references.                                         |
+---------------+-----------------------------------------------------+
| 85            | Returns the advanced filter mode state of the       |
|               | sheet. This is the mode without drop-down arrows on |
|               | top. Returns TRUE if the list has been filtered by  |
|               | clicking Filter, then Advanced Filter on the Data   |
|               | menu. Otherwise, returns FALSE.                     |
+---------------+-----------------------------------------------------+
| 86            | Returns the automatic filter mode state of the      |
|               | sheet. This is the mode with drop-down arrows on    |
|               | top. Returns TRUE if you have chosen Filter, then   |
|               | AutoFilter from the Data menu and the filter        |
|               | drop-down arrows are displayed. Otherwise, returns  |
|               | FALSE.                                              |
+---------------+-----------------------------------------------------+
| 87            | Returns the position number of the sheet. The first |
|               | sheet is position 1. Hidden sheet are included in   |
|               | the count.                                          |
+---------------+-----------------------------------------------------+
| 88            | Returns the name of the active workbook in the form |
|               | \"Book1\".                                          |
+---------------+-----------------------------------------------------+

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

**Examples**

The following macro formula returns TRUE if the contents of the active
workbook are protected:

GET.DOCUMENT(7)

In Microsoft Excel for Windows, the following macro formula returns the
number of windows in SALES.XLS:

GET.DOCUMENT(13, \"SALES.XLS\")

In Microsoft Excel for the Macintosh, the following macro formula
returns 3 if the overlay chart on SALES CHART is a column chart:

GET.DOCUMENT(10, \"SALES CHART\")

To find out if SHEET1 is password-protected and if its contents and
windows are protected, enter the following formula in a three-cell
horizontal array:

GET.DOCUMENT({6, 7, 8}, \"SHEET1\")

**Related Functions**

GET.CELL   Returns information about the specified cell

GET.WINDOW   Returns information about a window

GET.WORKSPACE   Returns information about the workspace

Return to [top](#E)

GET.FORMULA
