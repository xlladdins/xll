# GET.CELL

Returns information about the formatting, location, or contents of a
cell. Use GET.CELL in a macro whose behavior is determined by the status
of a particular cell.

**Syntax**

**GET.CELL**(**type\_num**, reference)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies what type of
cell information you want. The following list shows the possible values
of type\_num and the corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Absolute reference of the upper-left cell in reference, as text in the current workspace reference style.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Row number of the top cell in reference.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Column number of the leftmost cell in reference.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Same as TYPE(reference).</td>
</tr>
<tr class="even">
<td>5</td>
<td>Contents of reference.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Formula in reference, as text, in either A1 or R1C1 style depending on the workspace setting.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Number format of the cell, as text (for example, "m/d/yy" or "General").</td>
</tr>
<tr class="odd">
<td>8</td>
<td><p>Number indicating the cell's horizontal alignment:</p>
<p>1 = General</p>
<p>2 = Left</p>
<p>3 = Center</p>
<p>4 = Right</p>
<p>5 = Fill</p>
<p>6 = Justify</p>
<p>7 = Center across cells</p></td>
</tr>
<tr class="even">
<td>9</td>
<td><p>Number indicating the left-border style assigned to the cell:</p>
<p>0 = No border</p>
<p>1 = Thin line</p>
<p>2 = Medium line</p>
<p>3 = Dashed line</p>
<p>4 = Dotted line</p>
<p>5 = Thick line</p>
<p>6 = Double line</p>
<p>7 = Hairline</p></td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating the right-border style assigned to the cell. See type_num 9 for descriptions of the numbers returned.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Number indicating the top-border style assigned to the cell. See type_num 9 for descriptions of the numbers returned.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Number indicating the bottom-border style assigned to the cell. See type_num 9 for descriptions of the numbers returned.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Number from 0 to 18, indicating the pattern of the selected cell as displayed in the Patterns tab of the Format Cells dialog box, which appears when you click the Cells command on the Format menu. If no pattern is selected, returns 0.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>If the cell is locked, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>If the cell's formula is hidden, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>A two-item horizontal array containing the width of the active cell and a logical value indicating whether the cell's width is set to change as the standard width changes (TRUE) or is a custom width (FALSE).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Row height of cell, in points.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Name of font, as text.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Size of font, in points.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If all the characters in the cell, or only the first character, are bold, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If all the characters in the cell, or only the first character, are italic, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>If all the characters in the cell, or only the first character, are underlined, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>23</td>
<td>If all the characters in the cell, or only the first character, are struck through, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Font color of the first character in the cell, as a number in the range 1 to 56. If font color is automatic, returns 0.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If all the characters in the cell, or only the first character, are outlined, returns TRUE; otherwise, returns FALSE. Outline font format is not supported by Microsoft Excel for Windows.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>If all the characters in the cell, or only the first character, are shadowed, returns TRUE; otherwise, returns FALSE. Shadow font format is not supported by Microsoft Excel for Windows.</td>
</tr>
<tr class="even">
<td>27</td>
<td><p>Number indicating whether a manual page break occurs at the cell:</p>
<p>0 = No break</p>
<p>1 = Row</p>
<p>2 = Column</p>
<p>3 = Both row and column</p></td>
</tr>
<tr class="odd">
<td>28</td>
<td>Row level (outline)</td>
</tr>
<tr class="even">
<td>29</td>
<td>Column level (outline).</td>
</tr>
<tr class="odd">
<td>30</td>
<td>If the row containing the active cell is a summary row, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>31</td>
<td>If the column containing the active cell is a summary column, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>Name of the workbook and sheet containing the cell If the window contains only a single sheet that has the same name as the workbook without its extension, returns only the name of the book, in the form BOOK1.XLS. Otherwise, returns the name of the sheet in the form "[Book1]Sheet1".</td>
</tr>
<tr class="even">
<td>33</td>
<td>If the cell is formatted to wrap, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>Left-border color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="even">
<td>35</td>
<td>Right-border color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>Top-border color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="even">
<td>37</td>
<td>Bottom-border color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>Shade foreground color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="even">
<td>39</td>
<td>Shade background color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="odd">
<td>40</td>
<td>Style of the cell, as text.</td>
</tr>
<tr class="even">
<td>41</td>
<td>Returns the formula in the active cell without translating it (useful for international macro sheets).</td>
</tr>
<tr class="odd">
<td>42</td>
<td>The horizontal distance, measured in points, from the left edge of the active window to the left edge of the cell. May be a negative number if the window is scrolled beyond the cell.</td>
</tr>
<tr class="even">
<td>43</td>
<td>The vertical distance, measured in points, from the top edge of the active window to the top edge of the cell. May be a negative number if the window is scrolled beyond the cell.</td>
</tr>
<tr class="odd">
<td>44</td>
<td>The horizontal distance, measured in points, from the left edge of the active window to the right edge of the cell. May be a negative number if the window is scrolled beyond the cell.</td>
</tr>
<tr class="even">
<td>45</td>
<td>The vertical distance, measured in points, from the top edge of the active window to the bottom edge of the cell. May be a negative number if the window is scrolled beyond the cell.</td>
</tr>
<tr class="odd">
<td>46</td>
<td>If the cell contains a text note, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>47</td>
<td>If the cell contains a sound note, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>48</td>
<td>If the cells contains a formula, returns TRUE; if a constant, returns FALSE.</td>
</tr>
<tr class="even">
<td>49</td>
<td>If the cell is part of an array, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>50</td>
<td><p>Number indicating the cell's vertical alignment:</p>
<p>1 = Top</p>
<p>2 = Center</p>
<p>3 = Bottom</p>
<p>4 = Justified</p></td>
</tr>
<tr class="even">
<td>51</td>
<td><p>Number indicating the cell's vertical orientation:</p>
<p>0 = Horizontal</p>
<p>1 = Vertical</p>
<p>2 = Upward</p>
<p>3 = Downward</p></td>
</tr>
<tr class="odd">
<td>52</td>
<td>The cell prefix (or text alignment) character, or empty text ("") if the cell does not contain one.</td>
</tr>
<tr class="even">
<td>53</td>
<td>Contents of the cell as it is currently displayed, as text, including any additional numbers or symbols resulting from the cell's formatting.</td>
</tr>
<tr class="odd">
<td>54</td>
<td>Returns the name of the PivotTable report containing the active cell.</td>
</tr>
<tr class="even">
<td>55</td>
<td><p>Returns the position of a cell within the PivotTable report.</p>
<p>0 = Row header</p>
<p>1 = Column header</p>
<p>2 = Page header</p>
<p>3 = Data header</p>
<p>4 = Row item</p>
<p>5 = Column item</p>
<p>6 = Page item</p>
<p>7 = Data item</p>
<p>8 = Table body</p></td>
</tr>
<tr class="odd">
<td>56</td>
<td>Returns the name of the field containing the active cell reference if inside a PivotTable report.</td>
</tr>
<tr class="even">
<td>57</td>
<td>Returns TRUE if all the characters in the cell, or only the first character, are formatted with a superscript font; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>58</td>
<td>Returns the font style as text of all the characters in the cell, or only the first character as displayed in the Font tab of the Format Cells dialog box: for example, "Bold Italic".</td>
</tr>
<tr class="even">
<td>7</td>
<td><p>Returns the number for the underline style:</p>
<p>1 = None</p>
<p>2 = Single</p>
<p>3 = Double</p>
<p>4 = Single accounting</p>
<p>5 = Double accounting</p></td>
</tr>
<tr class="odd">
<td>60</td>
<td>Returns TRUE if all the characters in the cell, or only the first character, are formatted with a subscript font; otherwise, it returns FALSE.</td>
</tr>
<tr class="even">
<td>61</td>
<td>Returns the name of the PivotTable item for the active cell, as text.</td>
</tr>
<tr class="odd">
<td>62</td>
<td>Returns the name of the workbook and the current sheet in the form "[Book1]Sheet1".</td>
</tr>
<tr class="even">
<td>63</td>
<td>Returns the fill (background) color of the cell.</td>
</tr>
<tr class="odd">
<td>64</td>
<td>Returns the pattern (foreground) color of the cell.</td>
</tr>
<tr class="even">
<td>65</td>
<td>Returns TRUE if the Add Indent alignment option is on (Far East versions of Microsoft Excel only); otherwise, it returns FALSE.</td>
</tr>
<tr class="odd">
<td>66</td>
<td>Returns the book name of the workbook containing the cell in the form BOOK1.XLS.</td>
</tr>
</tbody>
</table>

Reference&nbsp;&nbsp;&nbsp;&nbsp;is a cell or a range of cells from
which you want information.

  - > If reference is a range of cells, the cell in the upper-left
    > corner of the first range in reference is used.

  - > If reference is omitted, the active cell is assumed.


**Tip**&nbsp;&nbsp;&nbsp;Use GET.CELL(17) to determine the height of a
cell and GET.CELL(44) - GET.CELL(42) to determine the width.

**Examples**

The following macro formula returns TRUE if cell B4 on sheet Sheet1 is
bold:

GET.CELL(20, Sheet1\!$B$4)

You can use the information returned by GET.CELL to initiate an action.
The following macro formula runs a custom function named BoldCell if the
GET.CELL formula returns FALSE:

IF(GET.CELL(20, Sheet1\!$B$4), , BoldCell())

**Related Functions**

[ABSREF](ABSREF.md)&nbsp;&nbsp;&nbsp;Returns the absolute reference of a range of
cells to another range

[ACTIVE.CELL](ACTIVE.CELL.md)&nbsp;&nbsp;&nbsp;Returns the reference of the active cell

[GET.FORMULA](GET.FORMULA.md)&nbsp;&nbsp;&nbsp;Returns the contents of a cell

[GET.NAME](GET.NAME.md)&nbsp;&nbsp;&nbsp;Returns the definition of a name

[GET.NOTE](GET.NOTE.md)&nbsp;&nbsp;&nbsp;Returns characters from a note

[RELREF](RELREF.md)&nbsp;&nbsp;&nbsp;Returns a relative reference



Return to [README](README.md#G)

# GET.CELL

Returns information about the formatting, location, or contents of a
cell. Use GET.CELL in a macro whose behavior is determined by the status
of a particular cell.

**Syntax**

**GET.CELL**(**type\_num**, reference)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies what type of
cell information you want. The following list shows the possible values
of type\_num and the corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Absolute reference of the upper-left cell in reference, as text in the current workspace reference style.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Row number of the top cell in reference.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Column number of the leftmost cell in reference.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Same as TYPE(reference).</td>
</tr>
<tr class="even">
<td>5</td>
<td>Contents of reference.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Formula in reference, as text, in either A1 or R1C1 style depending on the workspace setting.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Number format of the cell, as text (for example, "m/d/yy" or "General").</td>
</tr>
<tr class="odd">
<td>8</td>
<td><p>Number indicating the cell's horizontal alignment:</p>
<p>1 = General</p>
<p>2 = Left</p>
<p>3 = Center</p>
<p>4 = Right</p>
<p>5 = Fill</p>
<p>6 = Justify</p>
<p>7 = Center across cells</p></td>
</tr>
<tr class="even">
<td>9</td>
<td><p>Number indicating the left-border style assigned to the cell:</p>
<p>0 = No border</p>
<p>1 = Thin line</p>
<p>2 = Medium line</p>
<p>3 = Dashed line</p>
<p>4 = Dotted line</p>
<p>5 = Thick line</p>
<p>6 = Double line</p>
<p>7 = Hairline</p></td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating the right-border style assigned to the cell. See type_num 9 for descriptions of the numbers returned.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Number indicating the top-border style assigned to the cell. See type_num 9 for descriptions of the numbers returned.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Number indicating the bottom-border style assigned to the cell. See type_num 9 for descriptions of the numbers returned.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Number from 0 to 18, indicating the pattern of the selected cell as displayed in the Patterns tab of the Format Cells dialog box, which appears when you click the Cells command on the Format menu. If no pattern is selected, returns 0.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>If the cell is locked, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>If the cell's formula is hidden, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>A two-item horizontal array containing the width of the active cell and a logical value indicating whether the cell's width is set to change as the standard width changes (TRUE) or is a custom width (FALSE).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Row height of cell, in points.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Name of font, as text.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Size of font, in points.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If all the characters in the cell, or only the first character, are bold, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If all the characters in the cell, or only the first character, are italic, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>If all the characters in the cell, or only the first character, are underlined, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>23</td>
<td>If all the characters in the cell, or only the first character, are struck through, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Font color of the first character in the cell, as a number in the range 1 to 56. If font color is automatic, returns 0.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If all the characters in the cell, or only the first character, are outlined, returns TRUE; otherwise, returns FALSE. Outline font format is not supported by Microsoft Excel for Windows.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>If all the characters in the cell, or only the first character, are shadowed, returns TRUE; otherwise, returns FALSE. Shadow font format is not supported by Microsoft Excel for Windows.</td>
</tr>
<tr class="even">
<td>27</td>
<td><p>Number indicating whether a manual page break occurs at the cell:</p>
<p>0 = No break</p>
<p>1 = Row</p>
<p>2 = Column</p>
<p>3 = Both row and column</p></td>
</tr>
<tr class="odd">
<td>28</td>
<td>Row level (outline)</td>
</tr>
<tr class="even">
<td>29</td>
<td>Column level (outline).</td>
</tr>
<tr class="odd">
<td>30</td>
<td>If the row containing the active cell is a summary row, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>31</td>
<td>If the column containing the active cell is a summary column, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>Name of the workbook and sheet containing the cell If the window contains only a single sheet that has the same name as the workbook without its extension, returns only the name of the book, in the form BOOK1.XLS. Otherwise, returns the name of the sheet in the form "[Book1]Sheet1".</td>
</tr>
<tr class="even">
<td>33</td>
<td>If the cell is formatted to wrap, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>Left-border color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="even">
<td>35</td>
<td>Right-border color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>Top-border color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="even">
<td>37</td>
<td>Bottom-border color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>Shade foreground color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="even">
<td>39</td>
<td>Shade background color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="odd">
<td>40</td>
<td>Style of the cell, as text.</td>
</tr>
<tr class="even">
<td>41</td>
<td>Returns the formula in the active cell without translating it (useful for international macro sheets).</td>
</tr>
<tr class="odd">
<td>42</td>
<td>The horizontal distance, measured in points, from the left edge of the active window to the left edge of the cell. May be a negative number if the window is scrolled beyond the cell.</td>
</tr>
<tr class="even">
<td>43</td>
<td>The vertical distance, measured in points, from the top edge of the active window to the top edge of the cell. May be a negative number if the window is scrolled beyond the cell.</td>
</tr>
<tr class="odd">
<td>44</td>
<td>The horizontal distance, measured in points, from the left edge of the active window to the right edge of the cell. May be a negative number if the window is scrolled beyond the cell.</td>
</tr>
<tr class="even">
<td>45</td>
<td>The vertical distance, measured in points, from the top edge of the active window to the bottom edge of the cell. May be a negative number if the window is scrolled beyond the cell.</td>
</tr>
<tr class="odd">
<td>46</td>
<td>If the cell contains a text note, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>47</td>
<td>If the cell contains a sound note, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>48</td>
<td>If the cells contains a formula, returns TRUE; if a constant, returns FALSE.</td>
</tr>
<tr class="even">
<td>49</td>
<td>If the cell is part of an array, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>50</td>
<td><p>Number indicating the cell's vertical alignment:</p>
<p>1 = Top</p>
<p>2 = Center</p>
<p>3 = Bottom</p>
<p>4 = Justified</p></td>
</tr>
<tr class="even">
<td>51</td>
<td><p>Number indicating the cell's vertical orientation:</p>
<p>0 = Horizontal</p>
<p>1 = Vertical</p>
<p>2 = Upward</p>
<p>3 = Downward</p></td>
</tr>
<tr class="odd">
<td>52</td>
<td>The cell prefix (or text alignment) character, or empty text ("") if the cell does not contain one.</td>
</tr>
<tr class="even">
<td>53</td>
<td>Contents of the cell as it is currently displayed, as text, including any additional numbers or symbols resulting from the cell's formatting.</td>
</tr>
<tr class="odd">
<td>54</td>
<td>Returns the name of the PivotTable report containing the active cell.</td>
</tr>
<tr class="even">
<td>55</td>
<td><p>Returns the position of a cell within the PivotTable report.</p>
<p>0 = Row header</p>
<p>1 = Column header</p>
<p>2 = Page header</p>
<p>3 = Data header</p>
<p>4 = Row item</p>
<p>5 = Column item</p>
<p>6 = Page item</p>
<p>7 = Data item</p>
<p>8 = Table body</p></td>
</tr>
<tr class="odd">
<td>56</td>
<td>Returns the name of the field containing the active cell reference if inside a PivotTable report.</td>
</tr>
<tr class="even">
<td>57</td>
<td>Returns TRUE if all the characters in the cell, or only the first character, are formatted with a superscript font; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>58</td>
<td>Returns the font style as text of all the characters in the cell, or only the first character as displayed in the Font tab of the Format Cells dialog box: for example, "Bold Italic".</td>
</tr>
<tr class="even">
<td>7</td>
<td><p>Returns the number for the underline style:</p>
<p>1 = None</p>
<p>2 = Single</p>
<p>3 = Double</p>
<p>4 = Single accounting</p>
<p>5 = Double accounting</p></td>
</tr>
<tr class="odd">
<td>60</td>
<td>Returns TRUE if all the characters in the cell, or only the first character, are formatted with a subscript font; otherwise, it returns FALSE.</td>
</tr>
<tr class="even">
<td>61</td>
<td>Returns the name of the PivotTable item for the active cell, as text.</td>
</tr>
<tr class="odd">
<td>62</td>
<td>Returns the name of the workbook and the current sheet in the form "[Book1]Sheet1".</td>
</tr>
<tr class="even">
<td>63</td>
<td>Returns the fill (background) color of the cell.</td>
</tr>
<tr class="odd">
<td>64</td>
<td>Returns the pattern (foreground) color of the cell.</td>
</tr>
<tr class="even">
<td>65</td>
<td>Returns TRUE if the Add Indent alignment option is on (Far East versions of Microsoft Excel only); otherwise, it returns FALSE.</td>
</tr>
<tr class="odd">
<td>66</td>
<td>Returns the book name of the workbook containing the cell in the form BOOK1.XLS.</td>
</tr>
</tbody>
</table>

Reference&nbsp;&nbsp;&nbsp;&nbsp;is a cell or a range of cells from
which you want information.

  - > If reference is a range of cells, the cell in the upper-left
    > corner of the first range in reference is used.

  - > If reference is omitted, the active cell is assumed.


**Tip**&nbsp;&nbsp;&nbsp;Use GET.CELL(17) to determine the height of a
cell and GET.CELL(44) - GET.CELL(42) to determine the width.

**Examples**

The following macro formula returns TRUE if cell B4 on sheet Sheet1 is
bold:

GET.CELL(20, Sheet1\!$B$4)

You can use the information returned by GET.CELL to initiate an action.
The following macro formula runs a custom function named BoldCell if the
GET.CELL formula returns FALSE:

IF(GET.CELL(20, Sheet1\!$B$4), , BoldCell())

**Related Functions**

[ABSREF](ABSREF.md)&nbsp;&nbsp;&nbsp;Returns the absolute reference of a range of
cells to another range

[ACTIVE.CELL](ACTIVE.CELL.md)&nbsp;&nbsp;&nbsp;Returns the reference of the active cell

[GET.FORMULA](GET.FORMULA.md)&nbsp;&nbsp;&nbsp;Returns the contents of a cell

[GET.NAME](GET.NAME.md)&nbsp;&nbsp;&nbsp;Returns the definition of a name

[GET.NOTE](GET.NOTE.md)&nbsp;&nbsp;&nbsp;Returns characters from a note

[RELREF](RELREF.md)&nbsp;&nbsp;&nbsp;Returns a relative reference



Return to [README](README.md#G)

# GET.CELL

Returns information about the formatting, location, or contents of a
cell. Use GET.CELL in a macro whose behavior is determined by the status
of a particular cell.

**Syntax**

**GET.CELL**(**type\_num**, reference)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies what type of
cell information you want. The following list shows the possible values
of type\_num and the corresponding results.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Absolute reference of the upper-left cell in reference, as text in the current workspace reference style.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Row number of the top cell in reference.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Column number of the leftmost cell in reference.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Same as TYPE(reference).</td>
</tr>
<tr class="even">
<td>5</td>
<td>Contents of reference.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Formula in reference, as text, in either A1 or R1C1 style depending on the workspace setting.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Number format of the cell, as text (for example, "m/d/yy" or "General").</td>
</tr>
<tr class="odd">
<td>8</td>
<td><p>Number indicating the cell's horizontal alignment:</p>
<p>1 = General</p>
<p>2 = Left</p>
<p>3 = Center</p>
<p>4 = Right</p>
<p>5 = Fill</p>
<p>6 = Justify</p>
<p>7 = Center across cells</p></td>
</tr>
<tr class="even">
<td>9</td>
<td><p>Number indicating the left-border style assigned to the cell:</p>
<p>0 = No border</p>
<p>1 = Thin line</p>
<p>2 = Medium line</p>
<p>3 = Dashed line</p>
<p>4 = Dotted line</p>
<p>5 = Thick line</p>
<p>6 = Double line</p>
<p>7 = Hairline</p></td>
</tr>
<tr class="odd">
<td>10</td>
<td>Number indicating the right-border style assigned to the cell. See type_num 9 for descriptions of the numbers returned.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Number indicating the top-border style assigned to the cell. See type_num 9 for descriptions of the numbers returned.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Number indicating the bottom-border style assigned to the cell. See type_num 9 for descriptions of the numbers returned.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Number from 0 to 18, indicating the pattern of the selected cell as displayed in the Patterns tab of the Format Cells dialog box, which appears when you click the Cells command on the Format menu. If no pattern is selected, returns 0.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>If the cell is locked, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>If the cell's formula is hidden, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>A two-item horizontal array containing the width of the active cell and a logical value indicating whether the cell's width is set to change as the standard width changes (TRUE) or is a custom width (FALSE).</td>
</tr>
<tr class="even">
<td>17</td>
<td>Row height of cell, in points.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Name of font, as text.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Size of font, in points.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If all the characters in the cell, or only the first character, are bold, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>21</td>
<td>If all the characters in the cell, or only the first character, are italic, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>22</td>
<td>If all the characters in the cell, or only the first character, are underlined, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>23</td>
<td>If all the characters in the cell, or only the first character, are struck through, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Font color of the first character in the cell, as a number in the range 1 to 56. If font color is automatic, returns 0.</td>
</tr>
<tr class="even">
<td>25</td>
<td>If all the characters in the cell, or only the first character, are outlined, returns TRUE; otherwise, returns FALSE. Outline font format is not supported by Microsoft Excel for Windows.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>If all the characters in the cell, or only the first character, are shadowed, returns TRUE; otherwise, returns FALSE. Shadow font format is not supported by Microsoft Excel for Windows.</td>
</tr>
<tr class="even">
<td>27</td>
<td><p>Number indicating whether a manual page break occurs at the cell:</p>
<p>0 = No break</p>
<p>1 = Row</p>
<p>2 = Column</p>
<p>3 = Both row and column</p></td>
</tr>
<tr class="odd">
<td>28</td>
<td>Row level (outline)</td>
</tr>
<tr class="even">
<td>29</td>
<td>Column level (outline).</td>
</tr>
<tr class="odd">
<td>30</td>
<td>If the row containing the active cell is a summary row, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>31</td>
<td>If the column containing the active cell is a summary column, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>Name of the workbook and sheet containing the cell If the window contains only a single sheet that has the same name as the workbook without its extension, returns only the name of the book, in the form BOOK1.XLS. Otherwise, returns the name of the sheet in the form "[Book1]Sheet1".</td>
</tr>
<tr class="even">
<td>33</td>
<td>If the cell is formatted to wrap, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>Left-border color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="even">
<td>35</td>
<td>Right-border color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>Top-border color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="even">
<td>37</td>
<td>Bottom-border color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>Shade foreground color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="even">
<td>39</td>
<td>Shade background color as a number in the range 1 to 56. If color is automatic, returns 0.</td>
</tr>
<tr class="odd">
<td>40</td>
<td>Style of the cell, as text.</td>
</tr>
<tr class="even">
<td>41</td>
<td>Returns the formula in the active cell without translating it (useful for international macro sheets).</td>
</tr>
<tr class="odd">
<td>42</td>
<td>The horizontal distance, measured in points, from the left edge of the active window to the left edge of the cell. May be a negative number if the window is scrolled beyond the cell.</td>
</tr>
<tr class="even">
<td>43</td>
<td>The vertical distance, measured in points, from the top edge of the active window to the top edge of the cell. May be a negative number if the window is scrolled beyond the cell.</td>
</tr>
<tr class="odd">
<td>44</td>
<td>The horizontal distance, measured in points, from the left edge of the active window to the right edge of the cell. May be a negative number if the window is scrolled beyond the cell.</td>
</tr>
<tr class="even">
<td>45</td>
<td>The vertical distance, measured in points, from the top edge of the active window to the bottom edge of the cell. May be a negative number if the window is scrolled beyond the cell.</td>
</tr>
<tr class="odd">
<td>46</td>
<td>If the cell contains a text note, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>47</td>
<td>If the cell contains a sound note, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>48</td>
<td>If the cells contains a formula, returns TRUE; if a constant, returns FALSE.</td>
</tr>
<tr class="even">
<td>49</td>
<td>If the cell is part of an array, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>50</td>
<td><p>Number indicating the cell's vertical alignment:</p>
<p>1 = Top</p>
<p>2 = Center</p>
<p>3 = Bottom</p>
<p>4 = Justified</p></td>
</tr>
<tr class="even">
<td>51</td>
<td><p>Number indicating the cell's vertical orientation:</p>
<p>0 = Horizontal</p>
<p>1 = Vertical</p>
<p>2 = Upward</p>
<p>3 = Downward</p></td>
</tr>
<tr class="odd">
<td>52</td>
<td>The cell prefix (or text alignment) character, or empty text ("") if the cell does not contain one.</td>
</tr>
<tr class="even">
<td>53</td>
<td>Contents of the cell as it is currently displayed, as text, including any additional numbers or symbols resulting from the cell's formatting.</td>
</tr>
<tr class="odd">
<td>54</td>
<td>Returns the name of the PivotTable report containing the active cell.</td>
</tr>
<tr class="even">
<td>55</td>
<td><p>Returns the position of a cell within the PivotTable report.</p>
<p>0 = Row header</p>
<p>1 = Column header</p>
<p>2 = Page header</p>
<p>3 = Data header</p>
<p>4 = Row item</p>
<p>5 = Column item</p>
<p>6 = Page item</p>
<p>7 = Data item</p>
<p>8 = Table body</p></td>
</tr>
<tr class="odd">
<td>56</td>
<td>Returns the name of the field containing the active cell reference if inside a PivotTable report.</td>
</tr>
<tr class="even">
<td>57</td>
<td>Returns TRUE if all the characters in the cell, or only the first character, are formatted with a superscript font; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>58</td>
<td>Returns the font style as text of all the characters in the cell, or only the first character as displayed in the Font tab of the Format Cells dialog box: for example, "Bold Italic".</td>
</tr>
<tr class="even">
<td>7</td>
<td><p>Returns the number for the underline style:</p>
<p>1 = None</p>
<p>2 = Single</p>
<p>3 = Double</p>
<p>4 = Single accounting</p>
<p>5 = Double accounting</p></td>
</tr>
<tr class="odd">
<td>60</td>
<td>Returns TRUE if all the characters in the cell, or only the first character, are formatted with a subscript font; otherwise, it returns FALSE.</td>
</tr>
<tr class="even">
<td>61</td>
<td>Returns the name of the PivotTable item for the active cell, as text.</td>
</tr>
<tr class="odd">
<td>62</td>
<td>Returns the name of the workbook and the current sheet in the form "[Book1]Sheet1".</td>
</tr>
<tr class="even">
<td>63</td>
<td>Returns the fill (background) color of the cell.</td>
</tr>
<tr class="odd">
<td>64</td>
<td>Returns the pattern (foreground) color of the cell.</td>
</tr>
<tr class="even">
<td>65</td>
<td>Returns TRUE if the Add Indent alignment option is on (Far East versions of Microsoft Excel only); otherwise, it returns FALSE.</td>
</tr>
<tr class="odd">
<td>66</td>
<td>Returns the book name of the workbook containing the cell in the form BOOK1.XLS.</td>
</tr>
</tbody>
</table>

Reference&nbsp;&nbsp;&nbsp;&nbsp;is a cell or a range of cells from
which you want information.

  - > If reference is a range of cells, the cell in the upper-left
    > corner of the first range in reference is used.

  - > If reference is omitted, the active cell is assumed.


**Tip**&nbsp;&nbsp;&nbsp;Use GET.CELL(17) to determine the height of a
cell and GET.CELL(44) - GET.CELL(42) to determine the width.

**Examples**

The following macro formula returns TRUE if cell B4 on sheet Sheet1 is
bold:

GET.CELL(20, Sheet1\!$B$4)

You can use the information returned by GET.CELL to initiate an action.
The following macro formula runs a custom function named BoldCell if the
GET.CELL formula returns FALSE:

IF(GET.CELL(20, Sheet1\!$B$4), , BoldCell())

**Related Functions**

[ABSREF](ABSREF.md)&nbsp;&nbsp;&nbsp;Returns the absolute reference of a range of
cells to another range

[ACTIVE.CELL](ACTIVE.CELL.md)&nbsp;&nbsp;&nbsp;Returns the reference of the active cell

[GET.FORMULA](GET.FORMULA.md)&nbsp;&nbsp;&nbsp;Returns the contents of a cell

[GET.NAME](GET.NAME.md)&nbsp;&nbsp;&nbsp;Returns the definition of a name

[GET.NOTE](GET.NOTE.md)&nbsp;&nbsp;&nbsp;Returns characters from a note

[RELREF](RELREF.md)&nbsp;&nbsp;&nbsp;Returns a relative reference



Return to [README](README.md#G)

