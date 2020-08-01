# GET.DOCUMENT

Returns information about a sheet in a workbook.

**Syntax**

**GET.DOCUMENT**(**type\_num**, name\_text)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies what type of
information you want. The following lists show the possible values of
type\_num and the corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the workbook and worksheet as text. If there is only one sheet in the workbook and the sheet name is the same as the workbook name less any extension, returns the name of the book. The book name does not include the drive, directory or folder, or window number. Otherwise, returns the book and sheet name in the form "[BOOK1.XLS]Sheet1". It is usually best to use GET.DOCUMENT(76) and GET.DOCUMENT(88) to return the name of the active worksheet and the active workbook.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Path of the directory or folder containing name_text, as text. If the workbook name_text hasn't been saved yet, returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td><p>Number indicating the type of sheet. If name_text is a sheet, then the return value is one of the following numbers. If name_text is a book, then the return value is always 5. If name_text is omitted, then the sheet type is returned. If the book has one sheet that is named the same as the book, then the sheet type is returned.</p>
<p>1 = Worksheet</p>
<p>2 = Chart</p>
<p>3 = Macro sheet</p>
<p>4 = Info window if active</p>
<p>5 = Reserved</p>
<p>6 = Module</p>
<p>7 = Dialog</p></td>
</tr>
<tr class="odd">
<td>4</td>
<td>If changes have been made to the sheet since it was last saved, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If the sheet is read-only, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the sheet is password protected, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If cells in a sheet, the contents of a sheet, or the series in a chart are protected, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If the workbook windows are protected, returns TRUE; otherwise, returns FALSE.</td>
</tr>
</tbody>
</table>

The next four values of type\_num apply only to charts.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>9</td>
<td><p>Number indicating the type of the main chart:</p>
<p>1 = Area</p>
<p>2 = Bar</p>
<p>3 = Column</p>
<p>4 = Line</p>
<p>5 = Pie</p>
<p>6 = XY (scatter)</p>
<p>7 = 3-D area</p>
<p>8 = 3-D column</p>
<p>9 = 3-D line</p>
<p>10 = 3-D pie</p>
<p>11 = Radar</p>
<p>12 = 3-D bar</p>
<p>13 = 3-D surface</p>
<p>14 = Doughnut</p></td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating the type of the overlay chart. Same as 1, 2, 3, 4, 5, 6, 11, and 14 for main chart above. If there is no overlay chart, returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Number of series in the main chart.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Number of series in the overlay chart.</td>
</tr>
</tbody>
</table>

The next values of type\_num apply to worksheets and macro sheets and to
charts when appropriate.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>9</td>
<td>Number of the first used row. If the sheet is empty, returns 0.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number of the last used row. If the sheet is empty, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Number of the first used column. If the sheet is empty, returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Number of the last used column. If the sheet is empty, returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Number of windows.</td>
</tr>
<tr class="odd">
<td>14</td>
<td><p>Number indicating calculation mode:</p>
<p>1 = Automatic</p>
<p>2 = Automatic except tables</p>
<p>3 = Manual</p></td>
</tr>
<tr class="even">
<td>15</td>
<td>If the Iteration check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Maximum number of iterations.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Maximum change between iterations.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If the Update Remote References check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If the Precision As Displayed check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If the 1904 Date System check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
</tbody>
</table>

Type\_num values of 21 through 29 correspond to the four default fonts
in previous versions of Microsoft Excel. These values are provided only
for macro compatibility.

The next values of type\_num apply to worksheets and macro sheets, and
to charts if indicated.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>30</td>
<td>Horizontal array of consolidation references for the current sheet, in the form of text. If the list is empty, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>31</td>
<td>Number from 1 to 11, indicating the function used in the current consolidation. The function that corresponds to each number is listed under the CONSOLIDATE function. The default function is SUM.</td>
</tr>
<tr class="even">
<td>32</td>
<td>Three-item horizontal array indicating the status of the check boxes in the Data Consolidate dialog box. An item is TRUE if the check box is selected or FALSE if the check box is cleared. The first item indicates the Top Row check box, the second the Left Column check box, and the third the Create Links To Source Data check box.</td>
</tr>
<tr class="odd">
<td>33</td>
<td>If the Recalculate Before Save check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>34</td>
<td>If the workbook is read-only recommended, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>35</td>
<td>If the workbook is write-reserved, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>36</td>
<td>If the workbook has a write-reservation password and it is opened with read/write permission, returns the name of the user who originally saved the file with the write-reservation password. If the file is opened as read-only, or if a password has not been added to the workbook, returns the name of the current user.</td>
</tr>
<tr class="odd">
<td>37</td>
<td>Number corresponding to the file type of the workbook as displayed in the Save As dialog box. See the SAVE.AS function for a list of all the file types that Microsoft Excel recognizes.</td>
</tr>
<tr class="even">
<td>38</td>
<td>If the Summary Rows Below Detail check box is selected in the Outline dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>If the Summary Columns To Right Of Detail check box is selected in the Outline dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If the Always Create Backup check box is selected in the Save Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td><p>Number from 1 to 3 indicating whether objects are displayed:</p>
<p>1 = All objects are displayed</p>
<p>2 = Placeholders for pictures and charts</p>
<p>3 = All objects are hidden</p></td>
</tr>
<tr class="even">
<td>42</td>
<td>Horizontal array of all objects in the sheet. If there are no objects, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If the Save External Link Values check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>If objects in a workbook are protected, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>45</td>
<td><p>A number from 0 to 3 indicating how windows are synchronized:</p>
<p>0 = Not synchronized</p>
<p>1 = Synchronized horizontally</p>
<p>2 = Synchronized vertically</p>
<p>3 = Synchronized horizontally and vertically</p></td>
</tr>
<tr class="even">
<td>46</td>
<td><p>A seven-item horizontal array of print settings that can be set by the LINE.PRINT macro function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>A logical value indicating whether output will be formatted (TRUE) or unformatted (FALSE) when printed</p></td>
</tr>
<tr class="odd">
<td>47</td>
<td>If the Transition Formula Evaluation check box is selected in the Transition tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>48</td>
<td>The standard column width setting.</td>
</tr>
</tbody>
</table>

The next values of type\_num correspond to printing and page settings.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>49</td>
<td>The starting page number, or the #N/A error value if none is specified or if "Auto" is entered in the First page Number text box on the Page tab of the Page Setup dialog box.</td>
</tr>
<tr class="odd">
<td>50</td>
<td>The total number of pages that would be printed based on current settings, excluding notes, or 1 if the document is a chart.</td>
</tr>
<tr class="even">
<td>51</td>
<td>The total number of pages that would be printed if you print only notes, or the #N/A error value if the document is a chart.</td>
</tr>
<tr class="odd">
<td>52</td>
<td>Four-item horizontal array indicating the margin settings (left, right, top, bottom) in the currently specified units.</td>
</tr>
<tr class="even">
<td>53</td>
<td><p>A number indicating the orientation:</p>
<p>1 = Portrait</p>
<p>2 = Landscape</p></td>
</tr>
<tr class="odd">
<td>54</td>
<td>The header as a text string, including formatting codes.</td>
</tr>
<tr class="even">
<td>55</td>
<td>The footer as a text string, including formatting codes.</td>
</tr>
<tr class="odd">
<td>56</td>
<td>Horizontal array of two logical values corresponding to horizontal and vertical centering.</td>
</tr>
<tr class="even">
<td>57</td>
<td>If row or column headings are to be printed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>58</td>
<td>If gridlines are to be printed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>59</td>
<td>If the sheet is printed in black and white only, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>60</td>
<td><p>A number from 1 to 3 indicating how the chart will be sized when it's printed:</p>
<p>1 = Size on screen</p>
<p>2 = Scale to fit page</p>
<p>3 = Use full page</p></td>
</tr>
<tr class="even">
<td>61</td>
<td><p>A number indicating the pagination order:</p>
<p>1 = Down, then over</p>
<p>2 = Over, then down</p>
<p>Returns the #N/A error value if the document is a chart.</p></td>
</tr>
<tr class="odd">
<td>62</td>
<td>Percentage of reduction or enlargement, or 100% if none is specified. Returns the #N/A error value if not supported by the current printer or if the document is a chart.</td>
</tr>
<tr class="even">
<td>63</td>
<td>A two-item horizontal array indicating the number of pages to which the printout should be scaled to fit, with the first item equal to the width (or #N/A if no width scaling is specified) and the second item equal to the height (or #N/A if no height scaling is specified). #N/A is also returned if the document is a chart.</td>
</tr>
<tr class="odd">
<td>64</td>
<td>An array of row numbers corresponding to rows that are immediately below a manual or automatic page break.</td>
</tr>
<tr class="even">
<td>65</td>
<td>An array of column numbers corresponding to columns that are immediately to the right of a manual or automatic page break.</td>
</tr>
</tbody>
</table>

**Note**&nbsp;&nbsp;&nbsp;GET.DOCUMENT(62) and GET.DOCUMENT(63) are
mutually exclusive. If one returns a value, then the other returns the
\#N/A error value.

The next values of type\_num correspond to various workbook settings.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>66</td>
<td>In Microsoft Excel for Windows, if the Transition Formula Entry check box is selected in the Transition tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>67</td>
<td>Microsoft Excel version 5.0 or later always returns TRUE here.</td>
</tr>
<tr class="even">
<td>68</td>
<td>Microsoft Excel version 5.0 or later always returns the book name.</td>
</tr>
<tr class="odd">
<td>69</td>
<td>Returns TRUE if Page Breaks is chosen in the View tab of the Options dialog box; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>70</td>
<td>Returns the names of all PivotTable reports in the current sheet as a horizontal array.</td>
</tr>
<tr class="odd">
<td>71</td>
<td>Returns an horizontal array of all the styles in a workbook.</td>
</tr>
<tr class="even">
<td>72</td>
<td>Returns an horizontal array of all chart types displayed on the current sheet.</td>
</tr>
<tr class="odd">
<td>73</td>
<td>Returns an array of the number of series in each chart of the current sheet.</td>
</tr>
<tr class="even">
<td>74</td>
<td>Returns the object ID of the control that currently has the focus on a running user-defined dialog (based on the dialog sheet).</td>
</tr>
<tr class="odd">
<td>75</td>
<td>Returns the object ID of the object that is the current default button on a running user-defined dialog (based on the dialog sheet).</td>
</tr>
<tr class="even">
<td>76</td>
<td>Returns the name of the active sheet or macro sheet in the form [Book1]Sheet1.</td>
</tr>
<tr class="odd">
<td>77</td>
<td><p>In Microsoft Excel for Windows, returns the paper size, as integer:</p>
<p>1 = Letter 8.5 x 11 in</p>
<p>2 = Letter Small 8.5 x 11 in</p>
<p>5 = Legal 8.5 x 14 in</p>
<p>9 = A4 210 x 297 mm</p>
<p>10 = A4 Small 210 x 297 mm</p>
<p>13 = B5 182 x 257 mm</p>
<p>18 = Note 8.5 x 11 in</p></td>
</tr>
<tr class="even">
<td>78</td>
<td>Returns the print resolution, as a horizontal array of two numbers.</td>
</tr>
<tr class="odd">
<td>79</td>
<td>Returns TRUE if the Draft Quality check box has been selected from the sheet tab in the Page Setup dialog box; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>80</td>
<td>Returns TRUE if the Comments checkbox has been selected on the Sheet tab in the Page Setup dialog box; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>81</td>
<td>Returns the Print Area from the Sheet tab of the Page Setup dialog box as a cell reference.</td>
</tr>
<tr class="even">
<td>82</td>
<td>Returns the Print Titles from the Sheet tab of the Page Setup dialog box as an array of cell references.</td>
</tr>
<tr class="odd">
<td>83</td>
<td>Returns TRUE if the worksheet is protected for scenarios; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>84</td>
<td>Returns the value of the first circular reference on the sheet, or #N/A if there are no circular references.</td>
</tr>
<tr class="odd">
<td>85</td>
<td>Returns the advanced filter mode state of the sheet. This is the mode without drop-down arrows on top. Returns TRUE if the list has been filtered by clicking Filter, then Advanced Filter on the Data menu. Otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>86</td>
<td>Returns the automatic filter mode state of the sheet. This is the mode with drop-down arrows on top. Returns TRUE if you have chosen Filter, then AutoFilter from the Data menu and the filter drop-down arrows are displayed. Otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>87</td>
<td>Returns the position number of the sheet. The first sheet is position 1. Hidden sheet are included in the count.</td>
</tr>
<tr class="even">
<td>88</td>
<td>Returns the name of the active workbook in the form "Book1".</td>
</tr>
</tbody>
</table>

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of an open workbook. If
name\_text is omitted, it is assumed to be the active workbook.

**Examples**

The following macro formula returns TRUE if the contents of the active
workbook are protected:

GET.DOCUMENT(7)

In Microsoft Excel for Windows, the following macro formula returns the
number of windows in SALES.XLS:

GET.DOCUMENT(13, "SALES.XLS")

In Microsoft Excel for the Macintosh, the following macro formula
returns 3 if the overlay chart on SALES CHART is a column chart:

GET.DOCUMENT(10, "SALES CHART")

To find out if SHEET1 is password-protected and if its contents and
windows are protected, enter the following formula in a three-cell
horizontal array:

GET.DOCUMENT({6, 7, 8}, "SHEET1")

**Related Functions**

[GET.CELL](GET.CELL.md)&nbsp;&nbsp;&nbsp;Returns information about the specified cell

[GET.WINDOW](GET.WINDOW.md)&nbsp;&nbsp;&nbsp;Returns information about a window

[GET.WORKSPACE](GET.WORKSPACE.md)&nbsp;&nbsp;&nbsp;Returns information about the workspace



Return to [README](README.md#G)

# GET.DOCUMENT

Returns information about a sheet in a workbook.

**Syntax**

**GET.DOCUMENT**(**type\_num**, name\_text)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies what type of
information you want. The following lists show the possible values of
type\_num and the corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the workbook and worksheet as text. If there is only one sheet in the workbook and the sheet name is the same as the workbook name less any extension, returns the name of the book. The book name does not include the drive, directory or folder, or window number. Otherwise, returns the book and sheet name in the form "[BOOK1.XLS]Sheet1". It is usually best to use GET.DOCUMENT(76) and GET.DOCUMENT(88) to return the name of the active worksheet and the active workbook.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Path of the directory or folder containing name_text, as text. If the workbook name_text hasn't been saved yet, returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td><p>Number indicating the type of sheet. If name_text is a sheet, then the return value is one of the following numbers. If name_text is a book, then the return value is always 5. If name_text is omitted, then the sheet type is returned. If the book has one sheet that is named the same as the book, then the sheet type is returned.</p>
<p>1 = Worksheet</p>
<p>2 = Chart</p>
<p>3 = Macro sheet</p>
<p>4 = Info window if active</p>
<p>5 = Reserved</p>
<p>6 = Module</p>
<p>7 = Dialog</p></td>
</tr>
<tr class="odd">
<td>4</td>
<td>If changes have been made to the sheet since it was last saved, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If the sheet is read-only, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the sheet is password protected, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If cells in a sheet, the contents of a sheet, or the series in a chart are protected, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If the workbook windows are protected, returns TRUE; otherwise, returns FALSE.</td>
</tr>
</tbody>
</table>

The next four values of type\_num apply only to charts.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>9</td>
<td><p>Number indicating the type of the main chart:</p>
<p>1 = Area</p>
<p>2 = Bar</p>
<p>3 = Column</p>
<p>4 = Line</p>
<p>5 = Pie</p>
<p>6 = XY (scatter)</p>
<p>7 = 3-D area</p>
<p>8 = 3-D column</p>
<p>9 = 3-D line</p>
<p>10 = 3-D pie</p>
<p>11 = Radar</p>
<p>12 = 3-D bar</p>
<p>13 = 3-D surface</p>
<p>14 = Doughnut</p></td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating the type of the overlay chart. Same as 1, 2, 3, 4, 5, 6, 11, and 14 for main chart above. If there is no overlay chart, returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Number of series in the main chart.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Number of series in the overlay chart.</td>
</tr>
</tbody>
</table>

The next values of type\_num apply to worksheets and macro sheets and to
charts when appropriate.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>9</td>
<td>Number of the first used row. If the sheet is empty, returns 0.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number of the last used row. If the sheet is empty, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Number of the first used column. If the sheet is empty, returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Number of the last used column. If the sheet is empty, returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Number of windows.</td>
</tr>
<tr class="odd">
<td>14</td>
<td><p>Number indicating calculation mode:</p>
<p>1 = Automatic</p>
<p>2 = Automatic except tables</p>
<p>3 = Manual</p></td>
</tr>
<tr class="even">
<td>15</td>
<td>If the Iteration check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Maximum number of iterations.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Maximum change between iterations.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If the Update Remote References check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If the Precision As Displayed check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If the 1904 Date System check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
</tbody>
</table>

Type\_num values of 21 through 29 correspond to the four default fonts
in previous versions of Microsoft Excel. These values are provided only
for macro compatibility.

The next values of type\_num apply to worksheets and macro sheets, and
to charts if indicated.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>30</td>
<td>Horizontal array of consolidation references for the current sheet, in the form of text. If the list is empty, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>31</td>
<td>Number from 1 to 11, indicating the function used in the current consolidation. The function that corresponds to each number is listed under the CONSOLIDATE function. The default function is SUM.</td>
</tr>
<tr class="even">
<td>32</td>
<td>Three-item horizontal array indicating the status of the check boxes in the Data Consolidate dialog box. An item is TRUE if the check box is selected or FALSE if the check box is cleared. The first item indicates the Top Row check box, the second the Left Column check box, and the third the Create Links To Source Data check box.</td>
</tr>
<tr class="odd">
<td>33</td>
<td>If the Recalculate Before Save check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>34</td>
<td>If the workbook is read-only recommended, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>35</td>
<td>If the workbook is write-reserved, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>36</td>
<td>If the workbook has a write-reservation password and it is opened with read/write permission, returns the name of the user who originally saved the file with the write-reservation password. If the file is opened as read-only, or if a password has not been added to the workbook, returns the name of the current user.</td>
</tr>
<tr class="odd">
<td>37</td>
<td>Number corresponding to the file type of the workbook as displayed in the Save As dialog box. See the SAVE.AS function for a list of all the file types that Microsoft Excel recognizes.</td>
</tr>
<tr class="even">
<td>38</td>
<td>If the Summary Rows Below Detail check box is selected in the Outline dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>If the Summary Columns To Right Of Detail check box is selected in the Outline dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If the Always Create Backup check box is selected in the Save Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td><p>Number from 1 to 3 indicating whether objects are displayed:</p>
<p>1 = All objects are displayed</p>
<p>2 = Placeholders for pictures and charts</p>
<p>3 = All objects are hidden</p></td>
</tr>
<tr class="even">
<td>42</td>
<td>Horizontal array of all objects in the sheet. If there are no objects, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If the Save External Link Values check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>If objects in a workbook are protected, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>45</td>
<td><p>A number from 0 to 3 indicating how windows are synchronized:</p>
<p>0 = Not synchronized</p>
<p>1 = Synchronized horizontally</p>
<p>2 = Synchronized vertically</p>
<p>3 = Synchronized horizontally and vertically</p></td>
</tr>
<tr class="even">
<td>46</td>
<td><p>A seven-item horizontal array of print settings that can be set by the LINE.PRINT macro function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>A logical value indicating whether output will be formatted (TRUE) or unformatted (FALSE) when printed</p></td>
</tr>
<tr class="odd">
<td>47</td>
<td>If the Transition Formula Evaluation check box is selected in the Transition tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>48</td>
<td>The standard column width setting.</td>
</tr>
</tbody>
</table>

The next values of type\_num correspond to printing and page settings.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>49</td>
<td>The starting page number, or the #N/A error value if none is specified or if "Auto" is entered in the First page Number text box on the Page tab of the Page Setup dialog box.</td>
</tr>
<tr class="odd">
<td>50</td>
<td>The total number of pages that would be printed based on current settings, excluding notes, or 1 if the document is a chart.</td>
</tr>
<tr class="even">
<td>51</td>
<td>The total number of pages that would be printed if you print only notes, or the #N/A error value if the document is a chart.</td>
</tr>
<tr class="odd">
<td>52</td>
<td>Four-item horizontal array indicating the margin settings (left, right, top, bottom) in the currently specified units.</td>
</tr>
<tr class="even">
<td>53</td>
<td><p>A number indicating the orientation:</p>
<p>1 = Portrait</p>
<p>2 = Landscape</p></td>
</tr>
<tr class="odd">
<td>54</td>
<td>The header as a text string, including formatting codes.</td>
</tr>
<tr class="even">
<td>55</td>
<td>The footer as a text string, including formatting codes.</td>
</tr>
<tr class="odd">
<td>56</td>
<td>Horizontal array of two logical values corresponding to horizontal and vertical centering.</td>
</tr>
<tr class="even">
<td>57</td>
<td>If row or column headings are to be printed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>58</td>
<td>If gridlines are to be printed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>59</td>
<td>If the sheet is printed in black and white only, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>60</td>
<td><p>A number from 1 to 3 indicating how the chart will be sized when it's printed:</p>
<p>1 = Size on screen</p>
<p>2 = Scale to fit page</p>
<p>3 = Use full page</p></td>
</tr>
<tr class="even">
<td>61</td>
<td><p>A number indicating the pagination order:</p>
<p>1 = Down, then over</p>
<p>2 = Over, then down</p>
<p>Returns the #N/A error value if the document is a chart.</p></td>
</tr>
<tr class="odd">
<td>62</td>
<td>Percentage of reduction or enlargement, or 100% if none is specified. Returns the #N/A error value if not supported by the current printer or if the document is a chart.</td>
</tr>
<tr class="even">
<td>63</td>
<td>A two-item horizontal array indicating the number of pages to which the printout should be scaled to fit, with the first item equal to the width (or #N/A if no width scaling is specified) and the second item equal to the height (or #N/A if no height scaling is specified). #N/A is also returned if the document is a chart.</td>
</tr>
<tr class="odd">
<td>64</td>
<td>An array of row numbers corresponding to rows that are immediately below a manual or automatic page break.</td>
</tr>
<tr class="even">
<td>65</td>
<td>An array of column numbers corresponding to columns that are immediately to the right of a manual or automatic page break.</td>
</tr>
</tbody>
</table>

**Note**&nbsp;&nbsp;&nbsp;GET.DOCUMENT(62) and GET.DOCUMENT(63) are
mutually exclusive. If one returns a value, then the other returns the
\#N/A error value.

The next values of type\_num correspond to various workbook settings.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>66</td>
<td>In Microsoft Excel for Windows, if the Transition Formula Entry check box is selected in the Transition tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>67</td>
<td>Microsoft Excel version 5.0 or later always returns TRUE here.</td>
</tr>
<tr class="even">
<td>68</td>
<td>Microsoft Excel version 5.0 or later always returns the book name.</td>
</tr>
<tr class="odd">
<td>69</td>
<td>Returns TRUE if Page Breaks is chosen in the View tab of the Options dialog box; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>70</td>
<td>Returns the names of all PivotTable reports in the current sheet as a horizontal array.</td>
</tr>
<tr class="odd">
<td>71</td>
<td>Returns an horizontal array of all the styles in a workbook.</td>
</tr>
<tr class="even">
<td>72</td>
<td>Returns an horizontal array of all chart types displayed on the current sheet.</td>
</tr>
<tr class="odd">
<td>73</td>
<td>Returns an array of the number of series in each chart of the current sheet.</td>
</tr>
<tr class="even">
<td>74</td>
<td>Returns the object ID of the control that currently has the focus on a running user-defined dialog (based on the dialog sheet).</td>
</tr>
<tr class="odd">
<td>75</td>
<td>Returns the object ID of the object that is the current default button on a running user-defined dialog (based on the dialog sheet).</td>
</tr>
<tr class="even">
<td>76</td>
<td>Returns the name of the active sheet or macro sheet in the form [Book1]Sheet1.</td>
</tr>
<tr class="odd">
<td>77</td>
<td><p>In Microsoft Excel for Windows, returns the paper size, as integer:</p>
<p>1 = Letter 8.5 x 11 in</p>
<p>2 = Letter Small 8.5 x 11 in</p>
<p>5 = Legal 8.5 x 14 in</p>
<p>9 = A4 210 x 297 mm</p>
<p>10 = A4 Small 210 x 297 mm</p>
<p>13 = B5 182 x 257 mm</p>
<p>18 = Note 8.5 x 11 in</p></td>
</tr>
<tr class="even">
<td>78</td>
<td>Returns the print resolution, as a horizontal array of two numbers.</td>
</tr>
<tr class="odd">
<td>79</td>
<td>Returns TRUE if the Draft Quality check box has been selected from the sheet tab in the Page Setup dialog box; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>80</td>
<td>Returns TRUE if the Comments checkbox has been selected on the Sheet tab in the Page Setup dialog box; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>81</td>
<td>Returns the Print Area from the Sheet tab of the Page Setup dialog box as a cell reference.</td>
</tr>
<tr class="even">
<td>82</td>
<td>Returns the Print Titles from the Sheet tab of the Page Setup dialog box as an array of cell references.</td>
</tr>
<tr class="odd">
<td>83</td>
<td>Returns TRUE if the worksheet is protected for scenarios; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>84</td>
<td>Returns the value of the first circular reference on the sheet, or #N/A if there are no circular references.</td>
</tr>
<tr class="odd">
<td>85</td>
<td>Returns the advanced filter mode state of the sheet. This is the mode without drop-down arrows on top. Returns TRUE if the list has been filtered by clicking Filter, then Advanced Filter on the Data menu. Otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>86</td>
<td>Returns the automatic filter mode state of the sheet. This is the mode with drop-down arrows on top. Returns TRUE if you have chosen Filter, then AutoFilter from the Data menu and the filter drop-down arrows are displayed. Otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>87</td>
<td>Returns the position number of the sheet. The first sheet is position 1. Hidden sheet are included in the count.</td>
</tr>
<tr class="even">
<td>88</td>
<td>Returns the name of the active workbook in the form "Book1".</td>
</tr>
</tbody>
</table>

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of an open workbook. If
name\_text is omitted, it is assumed to be the active workbook.

**Examples**

The following macro formula returns TRUE if the contents of the active
workbook are protected:

GET.DOCUMENT(7)

In Microsoft Excel for Windows, the following macro formula returns the
number of windows in SALES.XLS:

GET.DOCUMENT(13, "SALES.XLS")

In Microsoft Excel for the Macintosh, the following macro formula
returns 3 if the overlay chart on SALES CHART is a column chart:

GET.DOCUMENT(10, "SALES CHART")

To find out if SHEET1 is password-protected and if its contents and
windows are protected, enter the following formula in a three-cell
horizontal array:

GET.DOCUMENT({6, 7, 8}, "SHEET1")

**Related Functions**

[GET.CELL](GET.CELL.md)&nbsp;&nbsp;&nbsp;Returns information about the specified cell

[GET.WINDOW](GET.WINDOW.md)&nbsp;&nbsp;&nbsp;Returns information about a window

[GET.WORKSPACE](GET.WORKSPACE.md)&nbsp;&nbsp;&nbsp;Returns information about the workspace



Return to [README](README.md#G)

# GET.DOCUMENT

Returns information about a sheet in a workbook.

**Syntax**

**GET.DOCUMENT**(**type\_num**, name\_text)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies what type of
information you want. The following lists show the possible values of
type\_num and the corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the workbook and worksheet as text. If there is only one sheet in the workbook and the sheet name is the same as the workbook name less any extension, returns the name of the book. The book name does not include the drive, directory or folder, or window number. Otherwise, returns the book and sheet name in the form "[BOOK1.XLS]Sheet1". It is usually best to use GET.DOCUMENT(76) and GET.DOCUMENT(88) to return the name of the active worksheet and the active workbook.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Path of the directory or folder containing name_text, as text. If the workbook name_text hasn't been saved yet, returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td><p>Number indicating the type of sheet. If name_text is a sheet, then the return value is one of the following numbers. If name_text is a book, then the return value is always 5. If name_text is omitted, then the sheet type is returned. If the book has one sheet that is named the same as the book, then the sheet type is returned.</p>
<p>1 = Worksheet</p>
<p>2 = Chart</p>
<p>3 = Macro sheet</p>
<p>4 = Info window if active</p>
<p>5 = Reserved</p>
<p>6 = Module</p>
<p>7 = Dialog</p></td>
</tr>
<tr class="odd">
<td>4</td>
<td>If changes have been made to the sheet since it was last saved, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If the sheet is read-only, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the sheet is password protected, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If cells in a sheet, the contents of a sheet, or the series in a chart are protected, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If the workbook windows are protected, returns TRUE; otherwise, returns FALSE.</td>
</tr>
</tbody>
</table>

The next four values of type\_num apply only to charts.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>9</td>
<td><p>Number indicating the type of the main chart:</p>
<p>1 = Area</p>
<p>2 = Bar</p>
<p>3 = Column</p>
<p>4 = Line</p>
<p>5 = Pie</p>
<p>6 = XY (scatter)</p>
<p>7 = 3-D area</p>
<p>8 = 3-D column</p>
<p>9 = 3-D line</p>
<p>10 = 3-D pie</p>
<p>11 = Radar</p>
<p>12 = 3-D bar</p>
<p>13 = 3-D surface</p>
<p>14 = Doughnut</p></td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating the type of the overlay chart. Same as 1, 2, 3, 4, 5, 6, 11, and 14 for main chart above. If there is no overlay chart, returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Number of series in the main chart.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Number of series in the overlay chart.</td>
</tr>
</tbody>
</table>

The next values of type\_num apply to worksheets and macro sheets and to
charts when appropriate.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>9</td>
<td>Number of the first used row. If the sheet is empty, returns 0.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number of the last used row. If the sheet is empty, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Number of the first used column. If the sheet is empty, returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Number of the last used column. If the sheet is empty, returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Number of windows.</td>
</tr>
<tr class="odd">
<td>14</td>
<td><p>Number indicating calculation mode:</p>
<p>1 = Automatic</p>
<p>2 = Automatic except tables</p>
<p>3 = Manual</p></td>
</tr>
<tr class="even">
<td>15</td>
<td>If the Iteration check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Maximum number of iterations.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Maximum change between iterations.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If the Update Remote References check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If the Precision As Displayed check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If the 1904 Date System check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
</tbody>
</table>

Type\_num values of 21 through 29 correspond to the four default fonts
in previous versions of Microsoft Excel. These values are provided only
for macro compatibility.

The next values of type\_num apply to worksheets and macro sheets, and
to charts if indicated.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>30</td>
<td>Horizontal array of consolidation references for the current sheet, in the form of text. If the list is empty, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>31</td>
<td>Number from 1 to 11, indicating the function used in the current consolidation. The function that corresponds to each number is listed under the CONSOLIDATE function. The default function is SUM.</td>
</tr>
<tr class="even">
<td>32</td>
<td>Three-item horizontal array indicating the status of the check boxes in the Data Consolidate dialog box. An item is TRUE if the check box is selected or FALSE if the check box is cleared. The first item indicates the Top Row check box, the second the Left Column check box, and the third the Create Links To Source Data check box.</td>
</tr>
<tr class="odd">
<td>33</td>
<td>If the Recalculate Before Save check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>34</td>
<td>If the workbook is read-only recommended, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>35</td>
<td>If the workbook is write-reserved, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>36</td>
<td>If the workbook has a write-reservation password and it is opened with read/write permission, returns the name of the user who originally saved the file with the write-reservation password. If the file is opened as read-only, or if a password has not been added to the workbook, returns the name of the current user.</td>
</tr>
<tr class="odd">
<td>37</td>
<td>Number corresponding to the file type of the workbook as displayed in the Save As dialog box. See the SAVE.AS function for a list of all the file types that Microsoft Excel recognizes.</td>
</tr>
<tr class="even">
<td>38</td>
<td>If the Summary Rows Below Detail check box is selected in the Outline dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>If the Summary Columns To Right Of Detail check box is selected in the Outline dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If the Always Create Backup check box is selected in the Save Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td><p>Number from 1 to 3 indicating whether objects are displayed:</p>
<p>1 = All objects are displayed</p>
<p>2 = Placeholders for pictures and charts</p>
<p>3 = All objects are hidden</p></td>
</tr>
<tr class="even">
<td>42</td>
<td>Horizontal array of all objects in the sheet. If there are no objects, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If the Save External Link Values check box is selected in the Calculation tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>If objects in a workbook are protected, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>45</td>
<td><p>A number from 0 to 3 indicating how windows are synchronized:</p>
<p>0 = Not synchronized</p>
<p>1 = Synchronized horizontally</p>
<p>2 = Synchronized vertically</p>
<p>3 = Synchronized horizontally and vertically</p></td>
</tr>
<tr class="even">
<td>46</td>
<td><p>A seven-item horizontal array of print settings that can be set by the LINE.PRINT macro function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>A logical value indicating whether output will be formatted (TRUE) or unformatted (FALSE) when printed</p></td>
</tr>
<tr class="odd">
<td>47</td>
<td>If the Transition Formula Evaluation check box is selected in the Transition tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>48</td>
<td>The standard column width setting.</td>
</tr>
</tbody>
</table>

The next values of type\_num correspond to printing and page settings.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>49</td>
<td>The starting page number, or the #N/A error value if none is specified or if "Auto" is entered in the First page Number text box on the Page tab of the Page Setup dialog box.</td>
</tr>
<tr class="odd">
<td>50</td>
<td>The total number of pages that would be printed based on current settings, excluding notes, or 1 if the document is a chart.</td>
</tr>
<tr class="even">
<td>51</td>
<td>The total number of pages that would be printed if you print only notes, or the #N/A error value if the document is a chart.</td>
</tr>
<tr class="odd">
<td>52</td>
<td>Four-item horizontal array indicating the margin settings (left, right, top, bottom) in the currently specified units.</td>
</tr>
<tr class="even">
<td>53</td>
<td><p>A number indicating the orientation:</p>
<p>1 = Portrait</p>
<p>2 = Landscape</p></td>
</tr>
<tr class="odd">
<td>54</td>
<td>The header as a text string, including formatting codes.</td>
</tr>
<tr class="even">
<td>55</td>
<td>The footer as a text string, including formatting codes.</td>
</tr>
<tr class="odd">
<td>56</td>
<td>Horizontal array of two logical values corresponding to horizontal and vertical centering.</td>
</tr>
<tr class="even">
<td>57</td>
<td>If row or column headings are to be printed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>58</td>
<td>If gridlines are to be printed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>59</td>
<td>If the sheet is printed in black and white only, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>60</td>
<td><p>A number from 1 to 3 indicating how the chart will be sized when it's printed:</p>
<p>1 = Size on screen</p>
<p>2 = Scale to fit page</p>
<p>3 = Use full page</p></td>
</tr>
<tr class="even">
<td>61</td>
<td><p>A number indicating the pagination order:</p>
<p>1 = Down, then over</p>
<p>2 = Over, then down</p>
<p>Returns the #N/A error value if the document is a chart.</p></td>
</tr>
<tr class="odd">
<td>62</td>
<td>Percentage of reduction or enlargement, or 100% if none is specified. Returns the #N/A error value if not supported by the current printer or if the document is a chart.</td>
</tr>
<tr class="even">
<td>63</td>
<td>A two-item horizontal array indicating the number of pages to which the printout should be scaled to fit, with the first item equal to the width (or #N/A if no width scaling is specified) and the second item equal to the height (or #N/A if no height scaling is specified). #N/A is also returned if the document is a chart.</td>
</tr>
<tr class="odd">
<td>64</td>
<td>An array of row numbers corresponding to rows that are immediately below a manual or automatic page break.</td>
</tr>
<tr class="even">
<td>65</td>
<td>An array of column numbers corresponding to columns that are immediately to the right of a manual or automatic page break.</td>
</tr>
</tbody>
</table>

**Note**&nbsp;&nbsp;&nbsp;GET.DOCUMENT(62) and GET.DOCUMENT(63) are
mutually exclusive. If one returns a value, then the other returns the
\#N/A error value.

The next values of type\_num correspond to various workbook settings.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>66</td>
<td>In Microsoft Excel for Windows, if the Transition Formula Entry check box is selected in the Transition tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>67</td>
<td>Microsoft Excel version 5.0 or later always returns TRUE here.</td>
</tr>
<tr class="even">
<td>68</td>
<td>Microsoft Excel version 5.0 or later always returns the book name.</td>
</tr>
<tr class="odd">
<td>69</td>
<td>Returns TRUE if Page Breaks is chosen in the View tab of the Options dialog box; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>70</td>
<td>Returns the names of all PivotTable reports in the current sheet as a horizontal array.</td>
</tr>
<tr class="odd">
<td>71</td>
<td>Returns an horizontal array of all the styles in a workbook.</td>
</tr>
<tr class="even">
<td>72</td>
<td>Returns an horizontal array of all chart types displayed on the current sheet.</td>
</tr>
<tr class="odd">
<td>73</td>
<td>Returns an array of the number of series in each chart of the current sheet.</td>
</tr>
<tr class="even">
<td>74</td>
<td>Returns the object ID of the control that currently has the focus on a running user-defined dialog (based on the dialog sheet).</td>
</tr>
<tr class="odd">
<td>75</td>
<td>Returns the object ID of the object that is the current default button on a running user-defined dialog (based on the dialog sheet).</td>
</tr>
<tr class="even">
<td>76</td>
<td>Returns the name of the active sheet or macro sheet in the form [Book1]Sheet1.</td>
</tr>
<tr class="odd">
<td>77</td>
<td><p>In Microsoft Excel for Windows, returns the paper size, as integer:</p>
<p>1 = Letter 8.5 x 11 in</p>
<p>2 = Letter Small 8.5 x 11 in</p>
<p>5 = Legal 8.5 x 14 in</p>
<p>9 = A4 210 x 297 mm</p>
<p>10 = A4 Small 210 x 297 mm</p>
<p>13 = B5 182 x 257 mm</p>
<p>18 = Note 8.5 x 11 in</p></td>
</tr>
<tr class="even">
<td>78</td>
<td>Returns the print resolution, as a horizontal array of two numbers.</td>
</tr>
<tr class="odd">
<td>79</td>
<td>Returns TRUE if the Draft Quality check box has been selected from the sheet tab in the Page Setup dialog box; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>80</td>
<td>Returns TRUE if the Comments checkbox has been selected on the Sheet tab in the Page Setup dialog box; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>81</td>
<td>Returns the Print Area from the Sheet tab of the Page Setup dialog box as a cell reference.</td>
</tr>
<tr class="even">
<td>82</td>
<td>Returns the Print Titles from the Sheet tab of the Page Setup dialog box as an array of cell references.</td>
</tr>
<tr class="odd">
<td>83</td>
<td>Returns TRUE if the worksheet is protected for scenarios; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>84</td>
<td>Returns the value of the first circular reference on the sheet, or #N/A if there are no circular references.</td>
</tr>
<tr class="odd">
<td>85</td>
<td>Returns the advanced filter mode state of the sheet. This is the mode without drop-down arrows on top. Returns TRUE if the list has been filtered by clicking Filter, then Advanced Filter on the Data menu. Otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>86</td>
<td>Returns the automatic filter mode state of the sheet. This is the mode with drop-down arrows on top. Returns TRUE if you have chosen Filter, then AutoFilter from the Data menu and the filter drop-down arrows are displayed. Otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>87</td>
<td>Returns the position number of the sheet. The first sheet is position 1. Hidden sheet are included in the count.</td>
</tr>
<tr class="even">
<td>88</td>
<td>Returns the name of the active workbook in the form "Book1".</td>
</tr>
</tbody>
</table>

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of an open workbook. If
name\_text is omitted, it is assumed to be the active workbook.

**Examples**

The following macro formula returns TRUE if the contents of the active
workbook are protected:

GET.DOCUMENT(7)

In Microsoft Excel for Windows, the following macro formula returns the
number of windows in SALES.XLS:

GET.DOCUMENT(13, "SALES.XLS")

In Microsoft Excel for the Macintosh, the following macro formula
returns 3 if the overlay chart on SALES CHART is a column chart:

GET.DOCUMENT(10, "SALES CHART")

To find out if SHEET1 is password-protected and if its contents and
windows are protected, enter the following formula in a three-cell
horizontal array:

GET.DOCUMENT({6, 7, 8}, "SHEET1")

**Related Functions**

[GET.CELL](GET.CELL.md)&nbsp;&nbsp;&nbsp;Returns information about the specified cell

[GET.WINDOW](GET.WINDOW.md)&nbsp;&nbsp;&nbsp;Returns information about a window

[GET.WORKSPACE](GET.WORKSPACE.md)&nbsp;&nbsp;&nbsp;Returns information about the workspace



Return to [README](README.md#G)

