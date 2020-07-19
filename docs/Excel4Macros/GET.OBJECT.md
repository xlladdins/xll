# GET.OBJECT

Returns information about the specified object. Use GET.OBJECT to return
information you can use in other macro formulas that manipulate objects.

**Syntax**

**GET.OBJECT**(**type\_num**, object\_id\_text, start\_num, count\_num,
item\_index)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the type of
information you want returned about an object. GET.OBJECT returns the
\#VALUE\! error value (and the macro is halted) if an object isn't
specified or if more than one object is selected.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>2</td>
<td>If the object is locked, returns TRUE; otherwise FALSE.</td>
</tr>
<tr class="odd">
<td>3</td>
<td>Z-order position (layering) of the object; that is, the relative position of the overlapping objects, starting with 1 for the object that is most under the others.</td>
</tr>
<tr class="even">
<td>4</td>
<td>Reference of the cell under the upper-left corner of the object as text in R1C1 reference style; for a line or arc, returns the start point.</td>
</tr>
<tr class="odd">
<td>5</td>
<td>X offset from the upper-left corner of the cell under the upper-left corner of the object, measured in points.</td>
</tr>
<tr class="even">
<td>6</td>
<td>Y offset from the upper-left corner of the cell under the upper-left corner of the object, measured in points.</td>
</tr>
<tr class="odd">
<td>7</td>
<td>Reference of the cell under the lower-right corner of the object as text in R1C1 reference style; for a line or arc, returns the end point.</td>
</tr>
<tr class="even">
<td>8</td>
<td>X offset from the upper-left corner of the cell under the lower-right corner of the object, measured in points.</td>
</tr>
<tr class="odd">
<td>9</td>
<td>Y offset from the upper-left corner of the cell under the lower-right corner of the object, measured in points.</td>
</tr>
<tr class="even">
<td>10</td>
<td>Name, including the filename, of the macro assigned to the object. If no macro is assigned, returns FALSE.</td>
</tr>
<tr class="odd">
<td>11</td>
<td>Number indicating how the object moves and sizes:<br />
1 = Object moves and sizes with cells<br />
2 = Object moves with cells<br />
3 = Object is fixed</td>
</tr>
</tbody>
</table>

Values 12 to 21 for type\_num apply only to text boxes and buttons. If
another type of object is selected, GET.OBJECT returns the \#VALUE\!
error value.

|               |             |
| ------------- | ----------- |
| **Type\_num** | **Returns** |

|    |                                                                                                                                                                                                                                                                     |
| -- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 12 | Text starting at start\_num for count\_num characters.                                                                                                                                                                                                              |
| 13 | Font name of all text starting at start\_num for count\_num characters. If the text contains more than one font name, returns the \#N/A error value.                                                                                                                |
| 14 | Font size of all text starting at start\_num for count\_num characters. If the text contains more than one font size, returns the \#N/A error value.                                                                                                                |
| 15 | If all text starting at start\_num for count\_num characters is bold, returns TRUE. If text contains only partial bold formatting, returns the \#N/A error value.                                                                                                   |
| 16 | If all text starting at start\_num for count\_num characters is italic, returns TRUE. If text contains only partial italic formatting, returns the \#N/A error value.                                                                                               |
| 17 | If all text starting at start\_num for count\_num characters is underlined, returns TRUE. If text contains only partial underline formatting, returns the \#N/A error value.                                                                                        |
| 18 | If all text starting at start\_num for count\_num characters is struck through, returns TRUE. If text contains only partial struck-through formatting, returns the \#N/A error value.                                                                               |
| 19 | In Microsoft Excel for the Macintosh, if all text starting at start\_num for count\_num characters is outlined, returns TRUE. If text contains only partial outline formatting, returns the \#N/A error value. Always returns FALSE in Microsoft Excel for Windows. |
| 20 | In Microsoft Excel for the Macintosh, if all text starting at start\_num for count\_num characters is shadowed, returns TRUE. If text contains only partial shadow formatting, returns the \#N/A error value. Always returns FALSE in Microsoft Excel for Windows.  |
| 21 | Number from 0 to 56 indicating the color of all text starting at start\_num for count\_num characters; if color is automatic, returns 0. If more than one color is used, returns the \#N/A error value.                                                             |

Values 22 to 25 for type\_num also apply only to text boxes and buttons.
If another type of object is selected, GET.OBJECT returns the \#N/A
error value.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>23</td>
<td>Number indicating the vertical alignment of text:<br />
1 = Top<br />
2 = Center<br />
3 = Bottom<br />
4 = Justified</td>
</tr>
<tr class="odd">
<td>24</td>
<td>Number indicating the orientation of text:<br />
0 = Horizontal<br />
1 = Vertical<br />
2 = Upward<br />
3 = Downward</td>
</tr>
<tr class="even">
<td>25</td>
<td>If button or text box is set to automatic sizing, returns TRUE; otherwise FALSE.</td>
</tr>
</tbody>
</table>

The following values for type\_num apply to all objects, except where
indicated.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>27</td>
<td>Number indicating the type of the border or line:<br />
0 = Custom<br />
1 = Automatic<br />
2 = None</td>
</tr>
<tr class="odd">
<td>28</td>
<td>Number indicating the style of the border or line as shown in the Patterns tab in the Format Objects dialog box:<br />
0 = None<br />
1 = Solid line<br />
2 = Dashed line<br />
3 = Dotted line<br />
4 = Dashed dotted line<br />
5 = Dashed double-dotted line<br />
6 = 50% gray line<br />
7 = 75% gray line<br />
8 = 25% gray line</td>
</tr>
<tr class="even">
<td>29</td>
<td>Number from 0 to 56 indicating the color of the border or line; if the border is automatic, returns 0.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>Number indicating the weight of the border or line:<br />
1 = Hairline<br />
2 = Thin<br />
3 = Medium<br />
4 = Thick</td>
</tr>
<tr class="even">
<td>31</td>
<td>Number indicating the type of fill:<br />
0 = Custom<br />
1 = Automatic<br />
2 = None</td>
</tr>
<tr class="odd">
<td>32</td>
<td>Number from 1 to 18 indicating the fill pattern as shown in the Format Object dialog box.</td>
</tr>
<tr class="even">
<td>33</td>
<td>Number from 0 to 56 indicating the foreground color of the fill pattern; if the fill is automatic, returns 0. If the object is a line, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>Number from 0 to 56 indicating the background color of the fill pattern; if the fill is automatic, returns 0. If the object is a line, returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>35</td>
<td>Number indicating the width of the arrowhead:<br />
1 = Narrow<br />
2 = Medium<br />
3 = Wide<br />
If the object is not a line, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>Number indicating the length of the arrowhead:<br />
1 = Short<br />
2 = Medium<br />
3 = Long<br />
If the object is not a line, returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>37</td>
<td>Number indicating the style of the arrowhead:<br />
1 = No head<br />
2 = Open head<br />
3 = Closed head<br />
4 = Open double-ended head<br />
5 = Closed double-ended head<br />
If the object is not a line, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>If the border has round corners, returns TRUE; if the corners are square, returns FALSE. If the object is a line, returns the #N/A error value.</td>
</tr>
<tr class="even">
<td>39</td>
<td>If the border has a shadow, returns TRUE; if the border has no shadow, returns FALSE. If the object is a line, returns the #N/A error value.</td>
</tr>
<tr class="odd">
<td>40</td>
<td>If the Lock Text check box in the Protection Tab of the Format Object dialog box is selected, returns TRUE; otherwise FALSE.</td>
</tr>
<tr class="even">
<td>41</td>
<td>If objects are set to be printed, returns TRUE; otherwise FALSE.</td>
</tr>
<tr class="odd">
<td>42</td>
<td>The horizontal distance, measured in points, from the left edge of the active window to the left edge of the object. May be a negative number if the window is scrolled beyond the object.</td>
</tr>
<tr class="even">
<td>43</td>
<td>The vertical distance, measured in points, from the top edge of the active window to the top edge of the object. May be a negative number if the window is scrolled beyond the object.</td>
</tr>
<tr class="odd">
<td>44</td>
<td>The horizontal distance, measured in points, from the left edge of the active window to the right edge of the object. May be a negative number if the window is scrolled beyond the object.</td>
</tr>
<tr class="even">
<td>45</td>
<td>The vertical distance, measured in points, from the top edge of the active window to the bottom edge of the object. May be a negative number if the window is scrolled beyond the object.</td>
</tr>
<tr class="odd">
<td>46</td>
<td>The number of vertices in a polygon, or the #N/A error value if the object is not a polygon.</td>
</tr>
<tr class="even">
<td>47</td>
<td>A count_num by 2 array of vertex coordinates starting at start_num in a polygon's array of vertices.</td>
</tr>
<tr class="odd">
<td>48</td>
<td>If the object is a text box, returns the cell reference that the text box is linked to. If the object is a control on a worksheet, returns the cell reference that the control's value is linked to. This information is returned as a string.</td>
</tr>
<tr class="even">
<td>49</td>
<td>Returns the ID number of the object. For example, "Rectangle 5" returns 5. Note that the name of the object may not have this index in it if the object has been renamed by the user.</td>
</tr>
<tr class="odd">
<td>50</td>
<td>Returns the object's classname. For example, "Rectangle".</td>
</tr>
<tr class="even">
<td>51</td>
<td>Returns the object name. By default, object names are the classname followed by the ID. For example, "Rectangle 1" is an object name, of which "Rectangle" is the classname, and 1 is the ID number. The object can also be renamed, in which case the name picked by the user is returned.</td>
</tr>
<tr class="odd">
<td>52</td>
<td>Returns the distance from cell A1 to the Left of the object bounding rectangle in points</td>
</tr>
<tr class="even">
<td>53</td>
<td>Returns the distance from Cell A1 to the top of the object bounding rectangle in points</td>
</tr>
<tr class="odd">
<td>54</td>
<td>Returns the width of object bounding rectangle in points</td>
</tr>
<tr class="even">
<td>55</td>
<td>Returns the height of object bounding rectangle in points</td>
</tr>
<tr class="odd">
<td>56</td>
<td>If the object is enabled, returns TRUE; otherwise, it returns FALSE.</td>
</tr>
<tr class="even">
<td>57</td>
<td>Returns the shortcut key assignment for the control object, as text.</td>
</tr>
<tr class="odd">
<td>58</td>
<td>Returns TRUE is the button control on a dialog sheet is the default button of the dialog; otherwise, returns FALSE</td>
</tr>
<tr class="even">
<td>59</td>
<td>Returns TRUE if the button control on the dialog sheet is clicked when the user presses the ESCAPE Key; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>60</td>
<td>Returns TRUE if the button control on a dialog sheet will close the dialog box when pressed; otherwise, returns FALSE</td>
</tr>
<tr class="even">
<td>61</td>
<td>Returns TRUE if the button control on a dialog sheet will be clicked when the user presses F1.</td>
</tr>
<tr class="odd">
<td>62</td>
<td>Returns the value of the control. For a check box or radio button, Returns 1 if it is selected, zero if it is not selected, or 2 if mixed. For a List box or dropdown box, returns the index number of the selected item, or zero if no item is selected. For a scroll bar, returns the numeric value of the scroll bar.</td>
</tr>
<tr class="even">
<td>63</td>
<td>Returns the minimum value that a scroll bar or spinner button can have</td>
</tr>
<tr class="odd">
<td>64</td>
<td>Returns the maximum value that a scroll bar or spinner button can have</td>
</tr>
<tr class="even">
<td>65</td>
<td>Returns the step increment value added or subtracted from the value of a scroll bar or spinner. This value is used when the arrow buttons are pressed on the control.</td>
</tr>
<tr class="odd">
<td>66</td>
<td>Returns the large, or "page" step increment value added or subtracted from the value of a scroll bar when it is clicked in the region between the thumb and the arrow buttons.</td>
</tr>
<tr class="even">
<td>67</td>
<td>Returns the input type allowed in an edit box control:<br />
1 = Text<br />
2 = Integer<br />
3 = Number (what type)<br />
4 = Cell reference<br />
5 = Formula</td>
</tr>
<tr class="odd">
<td>68</td>
<td>Returns TRUE if the edit box control allows multi-line editing with wrapped text; otherwise, it returns FALSE.</td>
</tr>
<tr class="even">
<td>69</td>
<td>Returns TRUE if the edit box has a vertical scroll bar; otherwise, it returns FALSE.</td>
</tr>
<tr class="odd">
<td>70</td>
<td>Returns the object ID of the object that is linked to a list box or edit box. For a dropdown combo box that has an editable entry field, returns the object ID of itself. A dropdown box that can't be edited, returns FALSE.</td>
</tr>
<tr class="even">
<td>71</td>
<td>Returns the number of entries in a List box, dropdown List box, or dropdown combo box.</td>
</tr>
<tr class="odd">
<td>72</td>
<td>Returns the text of the selected entry in a List box, dropdown List box, or dropdown combo box.</td>
</tr>
<tr class="even">
<td>73</td>
<td>Returns the range used to fill the entries in a List box, dropdown List box, or dropdown combo box, as text. If an empty string is returned, then the control isn't filled from a range.</td>
</tr>
<tr class="odd">
<td>74</td>
<td>Returns the number of list lines displayed when a dropdown control is dropped.</td>
</tr>
<tr class="even">
<td>75</td>
<td>Returns TRUE the object is displayed as 3-D; otherwise, it returns FALSE.</td>
</tr>
<tr class="odd">
<td>76</td>
<td>Returns the Far East phonetic accelerator key as text. Used for Far East versions of Microsoft Excel.</td>
</tr>
<tr class="even">
<td>77</td>
<td>Returns the select status of the list box:<br />
0 = single<br />
1 = simple multi-select<br />
2 = extended multi-select</td>
</tr>
<tr class="odd">
<td>78</td>
<td>Returns an array of TRUE and FALSE values indicating which items are selected in a list box. If TRUE, the item is selected; If FALSE, the item is not selected.</td>
</tr>
<tr class="even">
<td>79</td>
<td>Returns TRUE if the add indent attribute is on for alignment. Returns FALSE if the add indent attribute is off for alignment. Used for only Far East versions of Microsoft Excel.</td>
</tr>
</tbody>
</table>

Object\_id\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name and number, or
number alone, of the object you want information about. Object\_id\_text
is the text displayed in the reference area when the object is selected.
If object\_id\_text is omitted, it is assumed to be the selected object.
If object\_id\_text is omitted and no object is selected, GET.OBJECT
returns the \#REF\! error value and interrupts the macro.

Start\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number of the first character
in the text box or button or the first vertex in a polygon you want
information about. Start\_num is ignored unless a text box, button, or
polygon is specified by type\_num and object\_id\_text. If start\_num is
omitted, it is assumed to be 1.

Count\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number of characters in a text
box or button, or the number of vertices in a polygon, starting at
start\_num, that you want information about. Count\_num is ignored
unless a text box, button, or polygon is specified by type\_num and
object\_id\_text. If count\_num is omitted, it is assumed to be 255.

Item\_index&nbsp;&nbsp;&nbsp;&nbsp;is the index number or position of
the item in the list box or drop-down box that you want information
about, ranging from 1 to the number of items in the list box or
drop-down box.

**Tip**&nbsp;&nbsp;&nbsp;Use GET.OBJECT(45) - GET.OBJECT(43) to
determine the height of an object and GET.OBJECT(44) - GET.OBJECT(42) to
determine the width.

**Examples**

The following macro formula returns the reference of the cell under the
upper-left corner of the object Oval 3 (assume the cell is E2):

GET.OBJECT(4, "Oval 3") returns "R2C5"

The following macro formula changes the protection status of the object
Rectangle 2 if it is locked:

IF(GET.OBJECT(2, "Rectangle 2"), OBJECT.PROTECTION(FALSE))

The following macro formula returns characters 25 through 185 from the
object Text 5:

GET.OBJECT(12, "Text 5", 25, 160)

**Related Functions**

[CREATE.OBJECT](CREATE.OBJECT.md)&nbsp;&nbsp;&nbsp;Creates an object

[FONT.PROPERTIES](FONT.PROPERTIES.md)&nbsp;&nbsp;&nbsp;Applies a font to the selection

[OBJECT.PROTECTION](OBJECT.PROTECTION.md)&nbsp;&nbsp;&nbsp;Controls how an object is protected

[PLACEMENT](PLACEMENT.md)&nbsp;&nbsp;&nbsp;Determines an object's relationship to
underlying cells



Return to [README](README.md)

