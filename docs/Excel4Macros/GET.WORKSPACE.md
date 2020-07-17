GET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr class="odd">
<td>49</td>
<td>MAPI session currently in use, returned as a string of hex digits encoding the mail session value.</td>
</tr>
<tr class="even">
<td>50</td>
<td>If the Full Screen mode is on, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>51</td>
<td>If the formula bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>52</td>
<td>If the status bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>53</td>
<td>The name of the topmost custom dialog sheet currently running in a modal window, or #N/A if no dialog sheet is currently running.</td>
</tr>
<tr class="even">
<td>54</td>
<td>If the Edit Directly In Cell check box is selected on the Edit tab in the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>55</td>
<td>TRUE if the Alert Before Overwriting Cells check box in the Edit tab on Options dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>56</td>
<td>Standard font name in the General tab in the Options dialog box, as text.</td>
</tr>
<tr class="odd">
<td>57</td>
<td>Standard font size in the General tab in the Options dialog box, as a number</td>
</tr>
<tr class="even">
<td>58</td>
<td>If the Recently Used File list check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>59</td>
<td>If the Display Old Menus check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>60</td>
<td>If the Tip Wizard is enabled, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>61</td>
<td>Number of custom list entries listed in the Custom Lists tab of the Options dialog box.</td>
</tr>
<tr class="even">
<td>62</td>
<td>Returns information about available file converters.</td>
</tr>
<tr class="odd">
<td>63</td>
<td>Returns the type of mail system in use by Excel:<br />
0 = No mail transport detected<br />
1 = MAPI based transport<br />
2 = PowerTalk based transport (Macintosh only)</td>
</tr>
<tr class="even">
<td>64</td>
<td>If the Ask To Update Automatic Links check box in the Edit tab of the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>65</td>
<td>If the Cut, Copy, And Sort Objects With Cells check box in the Edit tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>66</td>
<td>Default number of sheets in a new workbook, as a number, from the General tab on Options dialog box.</td>
</tr>
<tr class="odd">
<td>67</td>
<td>Default file directory location, as text, from the General tab in the Options dialog box.</td>
</tr>
<tr class="even">
<td>68</td>
<td>If the Show ScreenTips On Toolbars check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>69</td>
<td>If the Large Icons check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>70</td>
<td>If the Prompt For Workbook Properties check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>71</td>
<td>TRUE if Microsoft Excel is open for in-place object editing (OLE). If FALSE, it is opened normally.</td>
</tr>
<tr class="even">
<td>72</td>
<td>TRUE if the Color Toolbars check box is selected in the Toolbars dialog box. FALSE if the Color Toolbars check box is not selected. This argument is for compatibility with Microsoft Excel version 5.0.</td>
</tr>
</tbody>
</table>

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

GET.WINDOW   Returns information about a window


GET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr class="odd">
<td>49</td>
<td>MAPI session currently in use, returned as a string of hex digits encoding the mail session value.</td>
</tr>
<tr class="even">
<td>50</td>
<td>If the Full Screen mode is on, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>51</td>
<td>If the formula bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>52</td>
<td>If the status bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>53</td>
<td>The name of the topmost custom dialog sheet currently running in a modal window, or #N/A if no dialog sheet is currently running.</td>
</tr>
<tr class="even">
<td>54</td>
<td>If the Edit Directly In Cell check box is selected on the Edit tab in the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>55</td>
<td>TRUE if the Alert Before Overwriting Cells check box in the Edit tab on Options dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>56</td>
<td>Standard font name in the General tab in the Options dialog box, as text.</td>
</tr>
<tr class="odd">
<td>57</td>
<td>Standard font size in the General tab in the Options dialog box, as a number</td>
</tr>
<tr class="even">
<td>58</td>
<td>If the Recently Used File list check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>59</td>
<td>If the Display Old Menus check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>60</td>
<td>If the Tip Wizard is enabled, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>61</td>
<td>Number of custom list entries listed in the Custom Lists tab of the Options dialog box.</td>
</tr>
<tr class="even">
<td>62</td>
<td>Returns information about available file converters.</td>
</tr>
<tr class="odd">
<td>63</td>
<td>Returns the type of mail system in use by Excel:<br />
0 = No mail transport detected<br />
1 = MAPI based transport<br />
2 = PowerTalk based transport (Macintosh only)</td>
</tr>
<tr class="even">
<td>64</td>
<td>If the Ask To Update Automatic Links check box in the Edit tab of the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>65</td>
<td>If the Cut, Copy, And Sort Objects With Cells check box in the Edit tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>66</td>
<td>Default number of sheets in a new workbook, as a number, from the General tab on Options dialog box.</td>
</tr>
<tr class="odd">
<td>67</td>
<td>Default file directory location, as text, from the General tab in the Options dialog box.</td>
</tr>
<tr class="even">
<td>68</td>
<td>If the Show ScreenTips On Toolbars check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>69</td>
<td>If the Large Icons check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>70</td>
<td>If the Prompt For Workbook Properties check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>71</td>
<td>TRUE if Microsoft Excel is open for in-place object editing (OLE). If FALSE, it is opened normally.</td>
</tr>
<tr class="even">
<td>72</td>
<td>TRUE if the Color Toolbars check box is selected in the Toolbars dialog box. FALSE if the Color Toolbars check box is not selected. This argument is for compatibility with Microsoft Excel version 5.0.</td>
</tr>
</tbody>
</table>

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

GET.WINDOW   Returns information about a window


GET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr class="odd">
<td>49</td>
<td>MAPI session currently in use, returned as a string of hex digits encoding the mail session value.</td>
</tr>
<tr class="even">
<td>50</td>
<td>If the Full Screen mode is on, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>51</td>
<td>If the formula bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>52</td>
<td>If the status bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>53</td>
<td>The name of the topmost custom dialog sheet currently running in a modal window, or #N/A if no dialog sheet is currently running.</td>
</tr>
<tr class="even">
<td>54</td>
<td>If the Edit Directly In Cell check box is selected on the Edit tab in the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>55</td>
<td>TRUE if the Alert Before Overwriting Cells check box in the Edit tab on Options dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>56</td>
<td>Standard font name in the General tab in the Options dialog box, as text.</td>
</tr>
<tr class="odd">
<td>57</td>
<td>Standard font size in the General tab in the Options dialog box, as a number</td>
</tr>
<tr class="even">
<td>58</td>
<td>If the Recently Used File list check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>59</td>
<td>If the Display Old Menus check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>60</td>
<td>If the Tip Wizard is enabled, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>61</td>
<td>Number of custom list entries listed in the Custom Lists tab of the Options dialog box.</td>
</tr>
<tr class="even">
<td>62</td>
<td>Returns information about available file converters.</td>
</tr>
<tr class="odd">
<td>63</td>
<td>Returns the type of mail system in use by Excel:<br />
0 = No mail transport detected<br />
1 = MAPI based transport<br />
2 = PowerTalk based transport (Macintosh only)</td>
</tr>
<tr class="even">
<td>64</td>
<td>If the Ask To Update Automatic Links check box in the Edit tab of the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>65</td>
<td>If the Cut, Copy, And Sort Objects With Cells check box in the Edit tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>66</td>
<td>Default number of sheets in a new workbook, as a number, from the General tab on Options dialog box.</td>
</tr>
<tr class="odd">
<td>67</td>
<td>Default file directory location, as text, from the General tab in the Options dialog box.</td>
</tr>
<tr class="even">
<td>68</td>
<td>If the Show ScreenTips On Toolbars check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>69</td>
<td>If the Large Icons check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>70</td>
<td>If the Prompt For Workbook Properties check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>71</td>
<td>TRUE if Microsoft Excel is open for in-place object editing (OLE). If FALSE, it is opened normally.</td>
</tr>
<tr class="even">
<td>72</td>
<td>TRUE if the Color Toolbars check box is selected in the Toolbars dialog box. FALSE if the Color Toolbars check box is not selected. This argument is for compatibility with Microsoft Excel version 5.0.</td>
</tr>
</tbody>
</table>

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

GET.WINDOW   Returns information about a window


GET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr class="odd">
<td>49</td>
<td>MAPI session currently in use, returned as a string of hex digits encoding the mail session value.</td>
</tr>
<tr class="even">
<td>50</td>
<td>If the Full Screen mode is on, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>51</td>
<td>If the formula bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>52</td>
<td>If the status bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>53</td>
<td>The name of the topmost custom dialog sheet currently running in a modal window, or #N/A if no dialog sheet is currently running.</td>
</tr>
<tr class="even">
<td>54</td>
<td>If the Edit Directly In Cell check box is selected on the Edit tab in the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>55</td>
<td>TRUE if the Alert Before Overwriting Cells check box in the Edit tab on Options dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>56</td>
<td>Standard font name in the General tab in the Options dialog box, as text.</td>
</tr>
<tr class="odd">
<td>57</td>
<td>Standard font size in the General tab in the Options dialog box, as a number</td>
</tr>
<tr class="even">
<td>58</td>
<td>If the Recently Used File list check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>59</td>
<td>If the Display Old Menus check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>60</td>
<td>If the Tip Wizard is enabled, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>61</td>
<td>Number of custom list entries listed in the Custom Lists tab of the Options dialog box.</td>
</tr>
<tr class="even">
<td>62</td>
<td>Returns information about available file converters.</td>
</tr>
<tr class="odd">
<td>63</td>
<td>Returns the type of mail system in use by Excel:<br />
0 = No mail transport detected<br />
1 = MAPI based transport<br />
2 = PowerTalk based transport (Macintosh only)</td>
</tr>
<tr class="even">
<td>64</td>
<td>If the Ask To Update Automatic Links check box in the Edit tab of the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>65</td>
<td>If the Cut, Copy, And Sort Objects With Cells check box in the Edit tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>66</td>
<td>Default number of sheets in a new workbook, as a number, from the General tab on Options dialog box.</td>
</tr>
<tr class="odd">
<td>67</td>
<td>Default file directory location, as text, from the General tab in the Options dialog box.</td>
</tr>
<tr class="even">
<td>68</td>
<td>If the Show ScreenTips On Toolbars check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>69</td>
<td>If the Large Icons check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>70</td>
<td>If the Prompt For Workbook Properties check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>71</td>
<td>TRUE if Microsoft Excel is open for in-place object editing (OLE). If FALSE, it is opened normally.</td>
</tr>
<tr class="even">
<td>72</td>
<td>TRUE if the Color Toolbars check box is selected in the Toolbars dialog box. FALSE if the Color Toolbars check box is not selected. This argument is for compatibility with Microsoft Excel version 5.0.</td>
</tr>
</tbody>
</table>

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

GET.WINDOW   Returns information about a window


GET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr class="odd">
<td>49</td>
<td>MAPI session currently in use, returned as a string of hex digits encoding the mail session value.</td>
</tr>
<tr class="even">
<td>50</td>
<td>If the Full Screen mode is on, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>51</td>
<td>If the formula bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>52</td>
<td>If the status bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>53</td>
<td>The name of the topmost custom dialog sheet currently running in a modal window, or #N/A if no dialog sheet is currently running.</td>
</tr>
<tr class="even">
<td>54</td>
<td>If the Edit Directly In Cell check box is selected on the Edit tab in the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>55</td>
<td>TRUE if the Alert Before Overwriting Cells check box in the Edit tab on Options dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>56</td>
<td>Standard font name in the General tab in the Options dialog box, as text.</td>
</tr>
<tr class="odd">
<td>57</td>
<td>Standard font size in the General tab in the Options dialog box, as a number</td>
</tr>
<tr class="even">
<td>58</td>
<td>If the Recently Used File list check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>59</td>
<td>If the Display Old Menus check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>60</td>
<td>If the Tip Wizard is enabled, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>61</td>
<td>Number of custom list entries listed in the Custom Lists tab of the Options dialog box.</td>
</tr>
<tr class="even">
<td>62</td>
<td>Returns information about available file converters.</td>
</tr>
<tr class="odd">
<td>63</td>
<td>Returns the type of mail system in use by Excel:<br />
0 = No mail transport detected<br />
1 = MAPI based transport<br />
2 = PowerTalk based transport (Macintosh only)</td>
</tr>
<tr class="even">
<td>64</td>
<td>If the Ask To Update Automatic Links check box in the Edit tab of the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>65</td>
<td>If the Cut, Copy, And Sort Objects With Cells check box in the Edit tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>66</td>
<td>Default number of sheets in a new workbook, as a number, from the General tab on Options dialog box.</td>
</tr>
<tr class="odd">
<td>67</td>
<td>Default file directory location, as text, from the General tab in the Options dialog box.</td>
</tr>
<tr class="even">
<td>68</td>
<td>If the Show ScreenTips On Toolbars check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>69</td>
<td>If the Large Icons check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>70</td>
<td>If the Prompt For Workbook Properties check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>71</td>
<td>TRUE if Microsoft Excel is open for in-place object editing (OLE). If FALSE, it is opened normally.</td>
</tr>
<tr class="even">
<td>72</td>
<td>TRUE if the Color Toolbars check box is selected in the Toolbars dialog box. FALSE if the Color Toolbars check box is not selected. This argument is for compatibility with Microsoft Excel version 5.0.</td>
</tr>
</tbody>
</table>

**Related Functions**
**Related Functions**

GET.DOCUMENT   Returns information about a workbook

GET.WINDOW   Returns information about a window


GET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr class="odd">
<td>49</td>
<td>MAPI session currently in use, returned as a string of hex digits encoding the mail session value.</td>
</tr>
<tr class="even">
<td>50</td>
<td>If the Full Screen mode is on, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>51</td>
<td>If the formula bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>52</td>
<td>If the status bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>53</td>
<td>The name of the topmost custom dialog sheet currently running in a modal window, or #N/A if no dialog sheet is currently running.</td>
</tr>
<tr class="even">
<td>54</td>
<td>If the Edit Directly In Cell check box is selected on the Edit tab in the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>55</td>
<td>TRUE if the Alert Before Overwriting Cells check box in the Edit tab on Options dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>56</td>
<td>Standard font name in the General tab in the Options dialog box, as text.</td>
</tr>
<tr class="odd">
<td>57</td>
<td>Standard font size in the General tab in the Options dialog box, as a number</td>
</tr>
<tr class="even">
<td>58</td>
<td>If the Recently Used File list check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>59</td>
<td>If the Display Old Menus check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>60</td>
<td>If the Tip Wizard is enabled, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>61</td>
<td>Number of custom list entries listed in the Custom Lists tab of the Options dialog box.</td>
</tr>
<tr class="even">
<td>62</td>
<td>Returns information about available file converters.</td>
</tr>
<tr class="odd">
<td>63</td>
<td>Returns the type of mail system in use by Excel:<br />
0 = No mail transport detected<br />
1 = MAPI based transport<br />
2 = PowerTalk based transport (Macintosh only)</td>
</tr>
<tr class="even">
<td>64</td>
<td>If the Ask To Update Automatic Links check box in the Edit tab of the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>65</td>
<td>If the Cut, Copy, And Sort Objects With Cells check box in the Edit tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>66</td>
<td>Default number of sheets in a new workbook, as a number, from the General tab on Options dialog box.</td>
</tr>
<tr class="odd">
<td>67</td>
<td>Default file directory location, as text, from the General tab in the Options dialog box.</td>
</tr>
<tr class="even">
<td>68</td>
<td>If the Show ScreenTips On Toolbars check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>69</td>
<td>If the Large Icons check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>70</td>
<td>If the Prompt For Workbook Properties check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>71</td>
<td>TRUE if Microsoft Excel is open for in-place object editing (OLE). If FALSE, it is opened normally.</td>
</tr>
<tr class="even">
<td>72</td>
<td>TRUE if the Color Toolbars check box is selected in the Toolbars dialog box. FALSE if the Color Toolbars check box is not selected. This argument is for compatibility with Microsoft Excel version 5.0.</td>
</tr>
</tbody>
</table>

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

GET.WINDOW   Returns information about a window


GET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr class="odd">
<td>49</td>
<td>MAPI session currently in use, returned as a string of hex digits encoding the mail session value.</td>
</tr>
<tr class="even">
<td>50</td>
<td>If the Full Screen mode is on, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>51</td>
<td>If the formula bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>52</td>
<td>If the status bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>53</td>
<td>The name of the topmost custom dialog sheet currently running in a modal window, or #N/A if no dialog sheet is currently running.</td>
</tr>
<tr class="even">
<td>54</td>
<td>If the Edit Directly In Cell check box is selected on the Edit tab in the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>55</td>
<td>TRUE if the Alert Before Overwriting Cells check box in the Edit tab on Options dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>56</td>
<td>Standard font name in the General tab in the Options dialog box, as text.</td>
</tr>
<tr class="odd">
<td>57</td>
<td>Standard font size in the General tab in the Options dialog box, as a number</td>
</tr>
<tr class="even">
<td>58</td>
<td>If the Recently Used File list check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>59</td>
<td>If the Display Old Menus check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>60</td>
<td>If the Tip Wizard is enabled, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>61</td>
<td>Number of custom list entries listed in the Custom Lists tab of the Options dialog box.</td>
</tr>
<tr class="even">
<td>62</td>
<td>Returns information about available file converters.</td>
</tr>
<tr class="odd">
<td>63</td>
<td>Returns the type of mail system in use by Excel:<br />
0 = No mail transport detected<br />
1 = MAPI based transport<br />
2 = PowerTalk based transport (Macintosh only)</td>
</tr>
<tr class="even">
<td>64</td>
<td>If the Ask To Update Automatic Links check box in the Edit tab of the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>65</td>
<td>If the Cut, Copy, And Sort Objects With Cells check box in the Edit tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>66</td>
<td>Default number of sheets in a new workbook, as a number, from the General tab on Options dialog box.</td>
</tr>
<tr class="odd">
<td>67</td>
<td>Default file directory location, as text, from the General tab in the Options dialog box.</td>
</tr>
<tr class="even">
<td>68</td>
<td>If the Show ScreenTips On Toolbars check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>69</td>
<td>If the Large Icons check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>70</td>
<td>If the Prompt For Workbook Properties check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>71</td>
<td>TRUE if Microsoft Excel is open for in-place object editing (OLE). If FALSE, it is opened normally.</td>
</tr>
<tr class="even">
<td>72</td>
<td>TRUE if the Color Toolbars check box is selected in the Toolbars dialog box. FALSE if the Color Toolbars check box is not selected. This argument is for compatibility with Microsoft Excel version 5.0.</td>
</tr>
</tbody>
</table>

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

GET.WINDOW   Returns information about a window


GET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr clasGET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr class="odd">
<td>49</td>
<td>MAPI session currently in use, returned as a string of hex digits encoding the mail session value.</td>
</tr>
<tr class="even">
<td>50</td>
<td>If the Full Screen mode is on, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>51</td>
<td>If the formula bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>52</td>
<td>If the status bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>53</td>
<td>The name of the topmost custom dialog sheet currently running in a modal window, or #N/A if no dialog sheet is currently running.</td>
</tr>
<tr class="even">
<td>54</td>
<td>If the Edit Directly In Cell check box is selected on the Edit tab in the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>55</td>
<td>TRUE if the Alert Before Overwriting Cells check box in the Edit tab on Options dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>56</td>
<td>Standard font name in the General tab in the Options dialog box, as text.</td>
</tr>
<tr class="odd">
<td>57</td>
<td>Standard font size in the General tab in the Options dialog box, as a number</td>
</tr>
<tr class="even">
<td>58</td>
<td>If the Recently Used File list check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>59</td>
<td>If the Display Old Menus check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>60</td>
<td>If the Tip Wizard is enabled, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>61</td>
<td>Number of custom list entries listed in the Custom Lists tab of the Options dialog box.</td>
</tr>
<tr class="even">
<td>62</td>
<td>Returns information about available file converters.</td>
</tr>
<tr class="odd">
<td>63</td>
<td>Returns the type of mail system in use by Excel:<br />
0 = No mail transport detected<br />
1 = MAPI based transport<br />
2 = PowerTalk based transport (Macintosh only)</td>
</tr>
<tr class="even">
<td>64</td>
<td>If the Ask To Update Automatic Links check box in the Edit tab of the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>65</td>
<td>If the Cut, Copy, And Sort Objects With Cells check box in the Edit tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>66</td>
<td>Default number of sheets in a new workbook, as a number, from the General tab on Options dialog box.</td>
</tr>
<tr class="odd">
<td>67</td>
<td>Default file directory location, as text, from the General tab in the Options dialog box.</td>
</tr>
<tr class="even">
<td>68</td>
<td>If the Show ScreenTips On Toolbars check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>69</td>
<td>If the Large Icons check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>70</td>
<td>If the Prompt For Workbook Properties check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>71</td>
<td>TRUE if Microsoft Excel is open for in-place object editing (OLE). If FALSE, it is opened normally.</td>
</tr>
<tr class="even">
<td>72</td>
<td>TRUE if the Color Toolbars check box is selected in the Toolbars dialog box. FALSE if the Color Toolbars check box is not selected. This argument is for compatibility with Microsoft Excel version 5.0.</td>
</tr>
</tbody>
</table>

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

GET.WINDOW   Returns information about a window


GET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr clasGET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr clasGET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr clasGET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr clasGET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr clasGET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr clasGET.WORKSPACE

Returns information about the workspace. Use GET.WORKSPACE in a macro
that depends on the status of the workspace, such as the environment,
version number, and available memory.

**Syntax**

**GET.WORKSPACE**(**type\_num**)

Type\_num    is a number specifying the type of workspace information
you want. The following list shows the type\_num values and their
corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Name of the environment in which Microsoft Excel is running, as text, followed by the environment's version number.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>The version number of Microsoft Excel, as text (for example, "5.0").</td>
</tr>
<tr class="even">
<td>3</td>
<td>If fixed decimals are set, returns the number of decimals; otherwise, returns 0.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>If in R1C1 mode, returns TRUE; if in A1 mode, returns FALSE.</td>
</tr>
<tr class="even">
<td>5</td>
<td>If scroll bars are displayed, returns TRUE; otherwise, returns FALSE. See also GET.WINDOW(26) and GET.WINDOW(27).</td>
</tr>
<tr class="odd">
<td>6</td>
<td>If the status bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>7</td>
<td>If the formula bar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>If remote DDE requests are enabled, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns the alternate menu key as text; if no alternate menu key is set, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating special modes:<br />
1 = Data Find<br />
2 = Copy<br />
3 = Cut<br />
4 = Data Entry<br />
5 = Unused<br />
6 = Copy and Data Entry<br />
7 = Cut and Data Entry<br />
If no special mode is set, returns 0.</td>
</tr>
<tr class="even">
<td>11</td>
<td>X position of the Microsoft Excel workspace window, measured in points from the left edge of the screen to the left edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Y position of the Microsoft Excel workspace window, measured in points from the top edge of the screen to the top edge of the window. In Microsoft Excel for the Macintosh, always returns 0.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Usable workspace width, in points.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Usable workspace height, in points.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number indicating maximized or minimized status of Microsoft Excel:<br />
1 = Neither<br />
2 = Minimized<br />
3 = Maximized<br />
Microsoft Excel for the Macintosh always returns 3.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Amount of memory free (in kilobytes).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Total memory available to Microsoft Excel (in kilobytes).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>If a math coprocessor is present, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If a mouse is present, returns TRUE; otherwise, returns FALSE. In Microsoft Excel for the Macintosh, always returns TRUE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If a group is present in the workspace, returns a horizontal array of sheets in the group; otherwise returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If the Standard toolbar is displayed, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>DDE-application-specific error code.</td>
</tr>
<tr class="even">
<td>23</td>
<td>Full path of the default startup directory or folder.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Full path of the alternate startup directory or folder; returns the #N/A error value if no alternate path has been specified.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If Microsoft Excel is set for relative recording, returns TRUE; if set for absolute recording, returns FALSE.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Name of user.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Name of organization.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>If Microsoft Excel menus are switched to by the transition menu or help key, returns 1; if Lotus 1-2-3 Help is switched to, returns 2.</td>
</tr>
<tr class="even">
<td>29</td>
<td>If transition navigation keys are enabled, returns TRUE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td><p>A nine-item horizontal array of global (default) print settings that can be set by the LINE.PRINT function:</p>
<p>Setup text</p>
<p>Left margin</p>
<p>Right margin</p>
<p>Top margin</p>
<p>Bottom margin</p>
<p>Page length</p>
<p>Logical value indicating whether to wait after printing each page (TRUE) or use continuous form feeding (FALSE)</p>
<p>Logical value indicating whether the printer has automatic line feeding (TRUE) or requires line feed characters (FALSE)</p>
<p>The number of the printer port</p></td>
</tr>
<tr class="even">
<td>31</td>
<td>If a currently running macro is in single step mode, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The current location of Microsoft Excel as a complete path.</td>
</tr>
<tr class="even">
<td>33</td>
<td>A horizontal array of the names in the New list, in the order they appear.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>A horizontal array of template files (with complete paths) in the New list, in the order they appear (returns the names of custom template files and the #N/A error value for built-in document types).</td>
</tr>
<tr class="even">
<td>35</td>
<td>If a macro is paused, returns TRUE; FALSE otherwise.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>If the Allow Cell Drag And Drop check box is selected in the Edit tab of the Options dialog box that appears when you click the Options command on the Tools menu, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>37</td>
<td><p>A 45-item horizontal array of the items related to country versions and settings. Use the following macro formula to return a specific item, where number is a number in the list below:</p>
<p>INDEX(GET.WORKSPACE(37), number)</p></td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to country codes:<br />
1 = Number corresponding to the country version of Microsoft Excel.<br />
2 = Number corresponding to the current country setting in the Microsoft Windows Control Panel or the country number as determined by your Apple system software</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to number separators:<br />
3 = Decimal separator<br />
4 = Zero (or 1000) separator<br />
5 = List separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to R1C1-style references:<br />
6 = Row character<br />
7 = Column character<br />
8 = Lowercase row character<br />
9 = Lowercase column character<br />
10 = Character used instead of the left bracket ([)<br />
11 = Character used instead of the right bracket (])</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to array characters:<br />
12 = Character used instead of the left bracket ({)<br />
13 = Character used instead of the right bracket (})<br />
14 = Column separator<br />
15 = Row separator<br />
16 = Alternate array item separator to use if the current array separator is the same as the decimal separator</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to format code symbols:<br />
17 = Date separator<br />
18 = Time separator<br />
19 = Year symbol<br />
20 = Month symbol<br />
21 = Day symbol<br />
22 = Hour symbol<br />
23 = Minute symbol<br />
24 = Second symbol<br />
25 = Currency symbol<br />
26 = "General" symbol</td>
</tr>
<tr class="even">
<td> </td>
<td>These values apply to format codes:<br />
27 = Number of decimal digits to use in currency formats<br />
28 = Number indicating the current format for negative currencies:<br />
   0 = ($currency) or (currency$)<br />
   1 = -$currency or -currency$<br />
   2 = $-currency or currency-$<br />
   3 = $currency- or currency$-<br />
where currency is any number and the $ represents the current currency symbol.<br />
29 = Number of decimal digits to use in noncurrency number formats<br />
30 = Number of characters to use in month names<br />
31 = Number of characters to use in weekday names<br />
32 = Number indicating the date order:<br />
   0 = Month-Day-Year<br />
   1 = Day-Month-Year<br />
   2 = Year-Month-Day</td>
</tr>
<tr class="odd">
<td> </td>
<td>These values apply to logical format values:<br />
33 = TRUE if using 24-hour time; FALSE if using 12-hour time.<br />
34 = TRUE if not displaying functions in English; otherwise, returns FALSE.<br />
35 = TRUE if using the metric system; FALSE if using the English measurement system.<br />
36 = TRUE if a space is added before the currency symbol; otherwise, returns FALSE.<br />
37 = TRUE if currency symbol precedes currency values; FALSE if it follows currency values.<br />
38 = TRUE if using minus sign for negative numbers; FALSE if using parentheses.<br />
39 = TRUE if trailing zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
40 = TRUE if leading zeros are displayed for zero currency values; otherwise, returns FALSE.<br />
41 = TRUE if leading zero is displayed in months (when months are displayed as numbers); otherwise, returns FALSE.<br />
42 = TRUE if leading zero is shown in days (when days are displayed as numbers); otherwise, returns FALSE.<br />
43 = TRUE if using four-digit years; FALSE if using two-digit years.<br />
44 = TRUE if date order is month-day-year when displaying dates in long form; FALSE if date order is day-month-year.<br />
45 = TRUE if leading zero is shown in the time; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>38</td>
<td>The number 0, 1, or 2 indicating the type of error-checking as set by the ERROR function. For more information, see ERROR.</td>
</tr>
<tr class="odd">
<td>39</td>
<td>A reference in R1C1-text form to the currently defined error-handling macro (set by the ERROR function), or the #N/A error value if none is specified.</td>
</tr>
<tr class="even">
<td>40</td>
<td>If screen updating is turned on (set by the ECHO function), returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>41</td>
<td>A horizontal array of cell ranges, as R1C1-style text, that were previously selected with the Go To command from the Edit menu or the FORMULA.GOTO macro function. If the book has multiple sheets, or if the single sheet in the workbook is named differently than the workbook itself, returns names as [Book]Sheet.</td>
</tr>
<tr class="even">
<td>42</td>
<td>If your computer is capable of playing sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>43</td>
<td>If your computer is capable of recording sounds, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>44</td>
<td>A three-column array of all currently registered procedures in dynamic link libraries (DLLs). The first column contains the names of the DLLs that contain the procedures (in Microsoft Excel for Windows) or the names of the files that contain the code resources (in Microsoft Excel for the Macintosh). The second column contains the names of the procedures in the DLLs (in Microsoft Excel for Windows) or code resources (in Microsoft Excel for the Macintosh). The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. For more information about DLLs and code resources and data types, see Using the CALL and REGISTER functions in Microsoft Excel Help.</td>
</tr>
<tr class="odd">
<td>45</td>
<td>If Microsoft Windows for Pen Computing is running, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>46</td>
<td>If the Move Selection After Enter check box is selected in the Edit tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>47</td>
<td>Reserved.</td>
</tr>
<tr class="even">
<td>48</td>
<td>Path to the library subdirectory for Microsoft Excel, as text.</td>
</tr>
<tr class="odd">
<td>49</td>
<td>MAPI session currently in use, returned as a string of hex digits encoding the mail session value.</td>
</tr>
<tr class="even">
<td>50</td>
<td>If the Full Screen mode is on, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>51</td>
<td>If the formula bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>52</td>
<td>If the status bar is displayed in Full Screen mode, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>53</td>
<td>The name of the topmost custom dialog sheet currently running in a modal window, or #N/A if no dialog sheet is currently running.</td>
</tr>
<tr class="even">
<td>54</td>
<td>If the Edit Directly In Cell check box is selected on the Edit tab in the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>55</td>
<td>TRUE if the Alert Before Overwriting Cells check box in the Edit tab on Options dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>56</td>
<td>Standard font name in the General tab in the Options dialog box, as text.</td>
</tr>
<tr class="odd">
<td>57</td>
<td>Standard font size in the General tab in the Options dialog box, as a number</td>
</tr>
<tr class="even">
<td>58</td>
<td>If the Recently Used File list check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>59</td>
<td>If the Display Old Menus check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>60</td>
<td>If the Tip Wizard is enabled, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>61</td>
<td>Number of custom list entries listed in the Custom Lists tab of the Options dialog box.</td>
</tr>
<tr class="even">
<td>62</td>
<td>Returns information about available file converters.</td>
</tr>
<tr class="odd">
<td>63</td>
<td>Returns the type of mail system in use by Excel:<br />
0 = No mail transport detected<br />
1 = MAPI based transport<br />
2 = PowerTalk based transport (Macintosh only)</td>
</tr>
<tr class="even">
<td>64</td>
<td>If the Ask To Update Automatic Links check box in the Edit tab of the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>65</td>
<td>If the Cut, Copy, And Sort Objects With Cells check box in the Edit tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>66</td>
<td>Default number of sheets in a new workbook, as a number, from the General tab on Options dialog box.</td>
</tr>
<tr class="odd">
<td>67</td>
<td>Default file directory location, as text, from the General tab in the Options dialog box.</td>
</tr>
<tr class="even">
<td>68</td>
<td>If the Show ScreenTips On Toolbars check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>69</td>
<td>If the Large Icons check box in the Options tab in the Customize dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>70</td>
<td>If the Prompt For Workbook Properties check box in the General tab on the Options dialog box is selected, returns TRUE; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>71</td>
<td>TRUE if Microsoft Excel is open for in-place object editing (OLE). If FALSE, it is opened normally.</td>
</tr>
<tr class="even">
<td>72</td>
<td>TRUE if the Color Toolbars check box is selected in the Toolbars dialog box. FALSE if the Color Toolbars check box is not selected. This argument is for compatibility with Microsoft Excel version 5.0.</td>
</tr>
</tbody>
</table>

**Related Functions**

[GET.DOCUMENT](GET.DOCUMENT.md)   Returns information about a workbook

[GET.WINDOW](GET.WINDOW.md)   Returns information about a window


