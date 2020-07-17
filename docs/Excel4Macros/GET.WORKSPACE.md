GET.WORKSPACE
=============

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

+---------------+-----------------------------------------------------+
| **Type\_num** | **Returns**                                         |
+---------------+-----------------------------------------------------+
| 1             | Name of the environment in which Microsoft Excel is |
|               | running, as text, followed by the environment\'s    |
|               | version number.                                     |
+---------------+-----------------------------------------------------+
| 2             | The version number of Microsoft Excel, as text (for |
|               | example, \"5.0\").                                  |
+---------------+-----------------------------------------------------+
| 3             | If fixed decimals are set, returns the number of    |
|               | decimals; otherwise, returns 0.                     |
+---------------+-----------------------------------------------------+
| 4             | If in R1C1 mode, returns TRUE; if in A1 mode,       |
|               | returns FALSE.                                      |
+---------------+-----------------------------------------------------+
| 5             | If scroll bars are displayed, returns TRUE;         |
|               | otherwise, returns FALSE. See also GET.WINDOW(26)   |
|               | and GET.WINDOW(27).                                 |
+---------------+-----------------------------------------------------+
| 6             | If the status bar is displayed, returns TRUE;       |
|               | otherwise, returns FALSE.                           |
+---------------+-----------------------------------------------------+
| 7             | If the formula bar is displayed, returns TRUE;      |
|               | otherwise, returns FALSE.                           |
+---------------+-----------------------------------------------------+
| 8             | If remote DDE requests are enabled, returns TRUE;   |
|               | otherwise, returns FALSE.                           |
+---------------+-----------------------------------------------------+
| 9             | Returns the alternate menu key as text; if no       |
|               | alternate menu key is set, returns the \#N/A error  |
|               | value.                                              |
+---------------+-----------------------------------------------------+
| 10            | Number indicating special modes:\                   |
|               | 1 = Data Find\                                      |
|               | 2 = Copy\                                           |
|               | 3 = Cut\                                            |
|               | 4 = Data Entry\                                     |
|               | 5 = Unused\                                         |
|               | 6 = Copy and Data Entry\                            |
|               | 7 = Cut and Data Entry\                             |
|               | If no special mode is set, returns 0.               |
+---------------+-----------------------------------------------------+
| 11            | X position of the Microsoft Excel workspace window, |
|               | measured in points from the left edge of the screen |
|               | to the left edge of the window. In Microsoft Excel  |
|               | for the Macintosh, always returns 0.                |
+---------------+-----------------------------------------------------+
| 12            | Y position of the Microsoft Excel workspace window, |
|               | measured in points from the top edge of the screen  |
|               | to the top edge of the window. In Microsoft Excel   |
|               | for the Macintosh, always returns 0.                |
+---------------+-----------------------------------------------------+
| 13            | Usable workspace width, in points.                  |
+---------------+-----------------------------------------------------+
| 14            | Usable workspace height, in points.                 |
+---------------+-----------------------------------------------------+
| 15            | Number indicating maximized or minimized status of  |
|               | Microsoft Excel:\                                   |
|               | 1 = Neither\                                        |
|               | 2 = Minimized\                                      |
|               | 3 = Maximized\                                      |
|               | Microsoft Excel for the Macintosh always returns 3. |
+---------------+-----------------------------------------------------+
| 16            | Amount of memory free (in kilobytes).               |
+---------------+-----------------------------------------------------+
| 17            | Total memory available to Microsoft Excel (in       |
|               | kilobytes).                                         |
+---------------+-----------------------------------------------------+
| 18            | If a math coprocessor is present, returns TRUE;     |
|               | otherwise, returns FALSE.                           |
+---------------+-----------------------------------------------------+
| 19            | If a mouse is present, returns TRUE; otherwise,     |
|               | returns FALSE. In Microsoft Excel for the           |
|               | Macintosh, always returns TRUE.                     |
+---------------+-----------------------------------------------------+
| 20            | If a group is present in the workspace, returns a   |
|               | horizontal array of sheets in the group; otherwise  |
|               | returns the \#N/A error value.                      |
+---------------+-----------------------------------------------------+
| 21            | If the Standard toolbar is displayed, returns TRUE; |
|               | otherwise, returns FALSE.                           |
+---------------+-----------------------------------------------------+
| 22            | DDE-application-specific error code.                |
+---------------+-----------------------------------------------------+
| 23            | Full path of the default startup directory or       |
|               | folder.                                             |
+---------------+-----------------------------------------------------+
| 24            | Full path of the alternate startup directory or     |
|               | folder; returns the \#N/A error value if no         |
|               | alternate path has been specified.                  |
+---------------+-----------------------------------------------------+
| 25            | If Microsoft Excel is set for relative recording,   |
|               | returns TRUE; if set for absolute recording,        |
|               | returns FALSE.                                      |
+---------------+-----------------------------------------------------+
| 26            | Name of user.                                       |
+---------------+-----------------------------------------------------+
| 27            | Name of organization.                               |
+---------------+-----------------------------------------------------+
| 28            | If Microsoft Excel menus are switched to by the     |
|               | transition menu or help key, returns 1; if Lotus    |
|               | 1-2-3 Help is switched to, returns 2.               |
+---------------+-----------------------------------------------------+
| 29            | If transition navigation keys are enabled, returns  |
|               | TRUE.                                               |
+---------------+-----------------------------------------------------+
| 30            | A nine-item horizontal array of global (default)    |
|               | print settings that can be set by the LINE.PRINT    |
|               | function:                                           |
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
|               | Logical value indicating whether to wait after      |
|               | printing each page (TRUE) or use continuous form    |
|               | feeding (FALSE)                                     |
|               |                                                     |
|               | Logical value indicating whether the printer has    |
|               | automatic line feeding (TRUE) or requires line feed |
|               | characters (FALSE)                                  |
|               |                                                     |
|               | The number of the printer port                      |
+---------------+-----------------------------------------------------+
| 31            | If a currently running macro is in single step      |
|               | mode, returns TRUE; otherwise, returns FALSE.       |
+---------------+-----------------------------------------------------+
| 32            | The current location of Microsoft Excel as a        |
|               | complete path.                                      |
+---------------+-----------------------------------------------------+
| 33            | A horizontal array of the names in the New list, in |
|               | the order they appear.                              |
+---------------+-----------------------------------------------------+
| 34            | A horizontal array of template files (with complete |
|               | paths) in the New list, in the order they appear    |
|               | (returns the names of custom template files and the |
|               | \#N/A error value for built-in document types).     |
+---------------+-----------------------------------------------------+
| 35            | If a macro is paused, returns TRUE; FALSE           |
|               | otherwise.                                          |
+---------------+-----------------------------------------------------+
| 36            | If the Allow Cell Drag And Drop check box is        |
|               | selected in the Edit tab of the Options dialog box  |
|               | that appears when you click the Options command on  |
|               | the Tools menu, returns TRUE; otherwise, returns    |
|               | FALSE.                                              |
+---------------+-----------------------------------------------------+
| 37            | A 45-item horizontal array of the items related to  |
|               | country versions and settings. Use the following    |
|               | macro formula to return a specific item, where      |
|               | number is a number in the list below:               |
|               |                                                     |
|               | INDEX(GET.WORKSPACE(37), number)                    |
+---------------+-----------------------------------------------------+
|               | These values apply to country codes:\               |
|               | 1 = Number corresponding to the country version of  |
|               | Microsoft Excel.\                                   |
|               | 2 = Number corresponding to the current country     |
|               | setting in the Microsoft Windows Control Panel or   |
|               | the country number as determined by your Apple      |
|               | system software                                     |
+---------------+-----------------------------------------------------+
|               | These values apply to number separators:\           |
|               | 3 = Decimal separator\                              |
|               | 4 = Zero (or 1000) separator\                       |
|               | 5 = List separator                                  |
+---------------+-----------------------------------------------------+
|               | These values apply to R1C1-style references:\       |
|               | 6 = Row character\                                  |
|               | 7 = Column character\                               |
|               | 8 = Lowercase row character\                        |
|               | 9 = Lowercase column character\                     |
|               | 10 = Character used instead of the left bracket     |
|               | (\[)\                                               |
|               | 11 = Character used instead of the right bracket    |
|               | (\])                                                |
+---------------+-----------------------------------------------------+
|               | These values apply to array characters:\            |
|               | 12 = Character used instead of the left bracket     |
|               | ({)\                                                |
|               | 13 = Character used instead of the right bracket    |
|               | (})\                                                |
|               | 14 = Column separator\                              |
|               | 15 = Row separator\                                 |
|               | 16 = Alternate array item separator to use if the   |
|               | current array separator is the same as the decimal  |
|               | separator                                           |
+---------------+-----------------------------------------------------+
|               | These values apply to format code symbols:\         |
|               | 17 = Date separator\                                |
|               | 18 = Time separator\                                |
|               | 19 = Year symbol\                                   |
|               | 20 = Month symbol\                                  |
|               | 21 = Day symbol\                                    |
|               | 22 = Hour symbol\                                   |
|               | 23 = Minute symbol\                                 |
|               | 24 = Second symbol\                                 |
|               | 25 = Currency symbol\                               |
|               | 26 = \"General\" symbol                             |
+---------------+-----------------------------------------------------+
|               | These values apply to format codes:\                |
|               | 27 = Number of decimal digits to use in currency    |
|               | formats\                                            |
|               | 28 = Number indicating the current format for       |
|               | negative currencies:\                               |
|               |    0 = (\$currency) or (currency\$)\                |
|               |    1 = -\$currency or -currency\$\                  |
|               |    2 = \$-currency or currency-\$\                  |
|               |    3 = \$currency- or currency\$-\                  |
|               | where currency is any number and the \$ represents  |
|               | the current currency symbol.\                       |
|               | 29 = Number of decimal digits to use in noncurrency |
|               | number formats\                                     |
|               | 30 = Number of characters to use in month names\    |
|               | 31 = Number of characters to use in weekday names\  |
|               | 32 = Number indicating the date order:\             |
|               |    0 = Month-Day-Year\                              |
|               |    1 = Day-Month-Year\                              |
|               |    2 = Year-Month-Day                               |
+---------------+-----------------------------------------------------+
|               | These values apply to logical format values:\       |
|               | 33 = TRUE if using 24-hour time; FALSE if using     |
|               | 12-hour time.\                                      |
|               | 34 = TRUE if not displaying functions in English;   |
|               | otherwise, returns FALSE.\                          |
|               | 35 = TRUE if using the metric system; FALSE if      |
|               | using the English measurement system.\              |
|               | 36 = TRUE if a space is added before the currency   |
|               | symbol; otherwise, returns FALSE.\                  |
|               | 37 = TRUE if currency symbol precedes currency      |
|               | values; FALSE if it follows currency values.\       |
|               | 38 = TRUE if using minus sign for negative numbers; |
|               | FALSE if using parentheses.\                        |
|               | 39 = TRUE if trailing zeros are displayed for zero  |
|               | currency values; otherwise, returns FALSE.\         |
|               | 40 = TRUE if leading zeros are displayed for zero   |
|               | currency values; otherwise, returns FALSE.\         |
|               | 41 = TRUE if leading zero is displayed in months    |
|               | (when months are displayed as numbers); otherwise,  |
|               | returns FALSE.\                                     |
|               | 42 = TRUE if leading zero is shown in days (when    |
|               | days are displayed as numbers); otherwise, returns  |
|               | FALSE.\                                             |
|               | 43 = TRUE if using four-digit years; FALSE if using |
|               | two-digit years.\                                   |
|               | 44 = TRUE if date order is month-day-year when      |
|               | displaying dates in long form; FALSE if date order  |
|               | is day-month-year.\                                 |
|               | 45 = TRUE if leading zero is shown in the time;     |
|               | otherwise, returns FALSE.                           |
+---------------+-----------------------------------------------------+
| 38            | The number 0, 1, or 2 indicating the type of        |
|               | error-checking as set by the ERROR function. For    |
|               | more information, see ERROR.                        |
+---------------+-----------------------------------------------------+
| 39            | A reference in R1C1-text form to the currently      |
|               | defined error-handling macro (set by the ERROR      |
|               | function), or the \#N/A error value if none is      |
|               | specified.                                          |
+---------------+-----------------------------------------------------+
| 40            | If screen updating is turned on (set by the ECHO    |
|               | function), returns TRUE; otherwise, returns FALSE.  |
+---------------+-----------------------------------------------------+
| 41            | A horizontal array of cell ranges, as R1C1-style    |
|               | text, that were previously selected with the Go To  |
|               | command from the Edit menu or the FORMULA.GOTO      |
|               | macro function. If the book has multiple sheets, or |
|               | if the single sheet in the workbook is named        |
|               | differently than the workbook itself, returns names |
|               | as \[Book\]Sheet.                                   |
+---------------+-----------------------------------------------------+
| 42            | If your computer is capable of playing sounds,      |
|               | returns TRUE; otherwise, returns FALSE.             |
+---------------+-----------------------------------------------------+
| 43            | If your computer is capable of recording sounds,    |
|               | returns TRUE; otherwise, returns FALSE.             |
+---------------+-----------------------------------------------------+
| 44            | A three-column array of all currently registered    |
|               | procedures in dynamic link libraries (DLLs). The    |
|               | first column contains the names of the DLLs that    |
|               | contain the procedures (in Microsoft Excel for      |
|               | Windows) or the names of the files that contain the |
|               | code resources (in Microsoft Excel for the          |
|               | Macintosh). The second column contains the names of |
|               | the procedures in the DLLs (in Microsoft Excel for  |
|               | Windows) or code resources (in Microsoft Excel for  |
|               | the Macintosh). The third column contains text      |
|               | strings specifying the data types of the return     |
|               | values, and the number and data types of the        |
|               | arguments. For more information about DLLs and code |
|               | resources and data types, see Using the CALL and    |
|               | REGISTER functions in Microsoft Excel Help.         |
+---------------+-----------------------------------------------------+
| 45            | If Microsoft Windows for Pen Computing is running,  |
|               | returns TRUE; otherwise, returns FALSE.             |
+---------------+-----------------------------------------------------+
| 46            | If the Move Selection After Enter check box is      |
|               | selected in the Edit tab of the Options dialog box, |
|               | returns TRUE; otherwise, returns FALSE.             |
+---------------+-----------------------------------------------------+
| 47            | Reserved.                                           |
+---------------+-----------------------------------------------------+
| 48            | Path to the library subdirectory for Microsoft      |
|               | Excel, as text.                                     |
+---------------+-----------------------------------------------------+
| 49            | MAPI session currently in use, returned as a string |
|               | of hex digits encoding the mail session value.      |
+---------------+-----------------------------------------------------+
| 50            | If the Full Screen mode is on, returns TRUE;        |
|               | otherwise, FALSE.                                   |
+---------------+-----------------------------------------------------+
| 51            | If the formula bar is displayed in Full Screen      |
|               | mode, returns TRUE; otherwise, FALSE.               |
+---------------+-----------------------------------------------------+
| 52            | If the status bar is displayed in Full Screen mode, |
|               | returns TRUE; otherwise, FALSE.                     |
+---------------+-----------------------------------------------------+
| 53            | The name of the topmost custom dialog sheet         |
|               | currently running in a modal window, or \#N/A if no |
|               | dialog sheet is currently running.                  |
+---------------+-----------------------------------------------------+
| 54            | If the Edit Directly In Cell check box is selected  |
|               | on the Edit tab in the Options dialog box, returns  |
|               | TRUE; otherwise, returns FALSE.                     |
+---------------+-----------------------------------------------------+
| 55            | TRUE if the Alert Before Overwriting Cells check    |
|               | box in the Edit tab on Options dialog box is        |
|               | selected; otherwise, FALSE.                         |
+---------------+-----------------------------------------------------+
| 56            | Standard font name in the General tab in the        |
|               | Options dialog box, as text.                        |
+---------------+-----------------------------------------------------+
| 57            | Standard font size in the General tab in the        |
|               | Options dialog box, as a number                     |
+---------------+-----------------------------------------------------+
| 58            | If the Recently Used File list check box in the     |
|               | General tab on the Options dialog box is selected,  |
|               | returns TRUE; otherwise, FALSE.                     |
+---------------+-----------------------------------------------------+
| 59            | If the Display Old Menus check box in the General   |
|               | tab on the Options dialog box is selected, returns  |
|               | TRUE; otherwise, FALSE.                             |
+---------------+-----------------------------------------------------+
| 60            | If the Tip Wizard is enabled, returns TRUE;         |
|               | otherwise, FALSE.                                   |
+---------------+-----------------------------------------------------+
| 61            | Number of custom list entries listed in the Custom  |
|               | Lists tab of the Options dialog box.                |
+---------------+-----------------------------------------------------+
| 62            | Returns information about available file            |
|               | converters.                                         |
+---------------+-----------------------------------------------------+
| 63            | Returns the type of mail system in use by Excel:\   |
|               | 0 = No mail transport detected\                     |
|               | 1 = MAPI based transport\                           |
|               | 2 = PowerTalk based transport (Macintosh only)      |
+---------------+-----------------------------------------------------+
| 64            | If the Ask To Update Automatic Links check box in   |
|               | the Edit tab of the Options dialog box is selected, |
|               | returns TRUE; otherwise, FALSE.                     |
+---------------+-----------------------------------------------------+
| 65            | If the Cut, Copy, And Sort Objects With Cells check |
|               | box in the Edit tab on the Options dialog box is    |
|               | selected, returns TRUE; otherwise, FALSE.           |
+---------------+-----------------------------------------------------+
| 66            | Default number of sheets in a new workbook, as a    |
|               | number, from the General tab on Options dialog box. |
+---------------+-----------------------------------------------------+
| 67            | Default file directory location, as text, from the  |
|               | General tab in the Options dialog box.              |
+---------------+-----------------------------------------------------+
| 68            | If the Show ScreenTips On Toolbars check box in the |
|               | Options tab in the Customize dialog box is          |
|               | selected, returns TRUE; otherwise, FALSE.           |
+---------------+-----------------------------------------------------+
| 69            | If the Large Icons check box in the Options tab in  |
|               | the Customize dialog box is selected, returns TRUE; |
|               | otherwise, FALSE.                                   |
+---------------+-----------------------------------------------------+
| 70            | If the Prompt For Workbook Properties check box in  |
|               | the General tab on the Options dialog box is        |
|               | selected, returns TRUE; otherwise, FALSE.           |
+---------------+-----------------------------------------------------+
| 71            | TRUE if Microsoft Excel is open for in-place object |
|               | editing (OLE). If FALSE, it is opened normally.     |
+---------------+-----------------------------------------------------+
| 72            | TRUE if the Color Toolbars check box is selected in |
|               | the Toolbars dialog box. FALSE if the Color         |
|               | Toolbars check box is not selected. This argument   |
|               | is for compatibility with Microsoft Excel version   |
|               | 5.0.                                                |
+---------------+-----------------------------------------------------+

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

GET.WINDOW   Returns information about a window

Return to [top](#E)

GOAL.SEEK
