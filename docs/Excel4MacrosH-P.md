<span id="H" class="anchor"></span>This document contains reference
information on the following Excel macro functions:

# H

[HALT](#halt), [HELP](#help), [HIDE](#hide),
[HIDE.DIALOG](#hide.dialog), [HIDE.OBJECT](#hide.object),
[HISTOGRAM](#histogram), [HLINE](#hline), [HPAGE](#hpage),
[HSCROLL](#hscroll)

# I

[IF](#if), [INITIATE](#initiate), [INPUT](#input), [INSERT](#insert),
[INSERT.OBJECT](#insert.object), [INSERT.PICTURE](#insert.picture),
[INSERT.TITLE](#insert.title)

# J

[JUSTIFY](#justify)

# L

[LABEL.PROPERTIES](#label.properties), [LAST.ERROR](#last.error),
[LEGEND](#legend), [LINE.PRINT](#line.print), [LINK.COMBO](#link.combo),
[LINK.FORMAT](#link.format), [LINKS](#links),
[LISTBOX.PROPERTIES](#listbox.properties), [LIST.NAMES](#list.names)

# M

[MACRO.OPTIONS](#macro.options), [MAIL.ADD.MAILER](#mail.add.mailer),
[MAIL.DELETE.MAILER](#mail.delete.mailer),
[MAIL.EDIT.MAILER](#mail.edit.mailer), [MAIL.FORWARD](#mail.forward),
[MAIL.LOGOFF](#mail.logoff), [MAIL.LOGON](#mail.logon),
[MAIL.NEXT.LETTER](#mail.next.letter), [MAIL.REPLY](#mail.reply),
[MAIL.REPLY.ALL](#mail.reply.all),
[MAIL.SEND.MAILER](#mail.send.mailer), [MAIN.CHART](#main.chart),
[MAIN.CHART.TYPE](#main.chart.type), [MCORREL](#mcorrel),
[MCOVAR](#mcovar), [MENU.EDITOR](#menu.editor),
[MERGE.STYLES](#merge.styles), [MESSAGE](#message), [MOVE](#move),
[MOVEAVG](#moveavg), [MOVE.TOOL](#move.tool)

# N

[NAMES](#names), [NEW](#new), [NEW.WINDOW](#new.window), [NEXT](#next),
[NOTE](#note)

# O

[OBJECT.PROPERTIES](#object.properties),
[OBJECT.PROTECTION](#object.protection), [ON.DATA](#on.data),
[ON.DOUBLECLICK](#on.doubleclick), [ON.ENTRY](#on.entry),
[ON.KEY](#on.key), [ON.RECALC](#on.recalc), [ON.SHEET](#on.sheet),
[ON.TIME](#on.time), [ON.WINDOW](#on.window), [OPEN](#open),
[OPEN.DIALOG](#open.dialog), [OPEN.LINKS](#open.links),
[OPEN.MAIL](#open.mail), [OPEN.TEXT](#open.text),
[OPTIONS.CALCULATION](#options.calculation),
[OPTIONS.CHART](#options.chart), [OPTIONS.EDIT](#options.edit),
[OPTIONS.GENERAL](#options.general),
[OPTIONS.LISTS.ADD](#options.lists.add),
[OPTIONS.LISTS.DELETE](#options.lists.delete),
[OPTIONS.LISTS.GET](#options.lists.get),
[OPTIONS.TRANSITION](#options.transition),
[OPTIONS.VIEW](#options.view), [OUTLINE](#outline), [OVERLAY](#overlay)

# P

[PAGE.SETUP](#page.setup), [PARSE](#parse), [PASTE](#paste),
[PASTE.LINK](#paste.link), [PASTE.PICTURE](#paste.picture),
[PASTE.PICTURE.LINK](#paste.picture.link),
[PASTE.SPECIAL](#paste.special), [PASTE.SPECIAL Syntax
1](#paste.special-syntax-1), [PASTE.SPECIAL Syntax
2](#paste.special-syntax-2), [PASTE.SPECIAL Syntax
3](#paste.special-syntax-3), [PASTE.SPECIAL Syntax
4](#paste.special-syntax-4), [PASTE.TOOL](#paste.tool),
[PATTERNS](#patterns), [PAUSE](#pause),
[PIVOT.ADD.DATA](#pivot.add.data),
[PIVOT.ADD.FIELDS](#pivot.add.fields), [PIVOT.FIELD](#pivot.field),
[PIVOT.FIELD.GROUP](#pivot.field.group),
[PIVOT.FIELD.PROPERTIES](#pivot.field.properties),
[PIVOT.FIELD.UNGROUP](#pivot.field.ungroup), [PIVOT.ITEM](#pivot.item),
[PIVOT.ITEM.PROPERTIES](#pivot.item.properties),
[PIVOT.REFRESH](#pivot.refresh), [PIVOT.SHOW.PAGES](#pivot.show.pages),
[PIVOT.TABLE.WIZARD](#pivot.table.wizard), [PLACEMENT](#placement),
[POKE](#poke), [PRECISION](#precision), [PREFERRED](#preferred),
[PRESS.TOOL](#press.tool), [PRINT](#print),
[PRINTER.SETUP](#printer.setup), [PRINT.PREVIEW](#print.preview),
[PROMOTE](#promote), [PROTECT.DOCUMENT](#protect.document),
[PTTESTM](#pttestm), [PTTESTV](#pttestv),
[PUSHBUTTON.PROPERTIES](#pushbutton.properties)

# HALT

Stops all macros from running. Use HALT instead of RETURN to prevent a
macro from returning to the macro that called it.

**Syntax**

**HALT**(cancel\_close)

Cancel\_close    is a logical value that specifies whether a macro
sheet, when encountering the HALT function in an Auto\_Close macro, is
closed.

  - > If cancel\_close is TRUE, Microsoft Excel halts the macro and
    > prevents the workbook from being closed.

  - > If cancel\_close is FALSE or omitted, Microsoft Excel halts the
    > macro and allows the workbook to be closed.

  - > If cancel\_close is specified in a macro that is not an
    > Auto\_Close macro, it is ignored and the HALT function simply
    > stops the current macro.

>  

**Remarks**

You can prevent an Auto\_Close or Auto\_Open macro from running by
holding down the SHIFT key while opening or closing the workbook.

**Examples**

If A1 contains the \#N/A error value, then when the following macro
formula is calculated, the macro halts:

IF(ISERROR(A1), HALT(), GOTO(D4))

The following macro formula at the end of an Auto\_Close macro ends the
macro and prevents the workbook from being closed:

HALT(TRUE)

**Related Functions**

BREAK   Interrupts a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

RETURN   Ends the currently running macro

Return to [top](#H)

# HELP

Starts or switches to Help and displays the specified custom Help topic.
Use HELP with custom Help files to create your own Help system, which
can be used just like the built-in Microsoft Excel Help.

**Syntax**

**HELP**(help\_ref)

Help\_ref    is a reference to a topic in a Help file, in the form
"filename\!topic\_number".

  - > Help\_ref must be given as text.

>  

**Remarks**

  - > Microsoft Excel for Windows does not support the use of Help files
    > in the text file format for custom Help.

  - > In Microsoft Excel for the Macintosh, custom Help files are plain
    > text files or text files with line breaks.

>  

**Tips**

  - > In Microsoft Excel for Windows, the following macro formula
    > switches back to Microsoft Excel when Help is active:

APP.ACTIVATE()

  - > The following macro formula closes Help when Help is active:

SEND.KEYS("%{F4}")

**Examples**

In Microsoft Excel for Windows, the following macro formula displays the
Help topic numbered 101 in the file CUSTHELP.DOC. The Help window
remains open if the user switches to another window or application.

HELP("CUSTHELP.DOC\!101")

If the custom Help file is not in the current directory, specify the
full path along with the name of the file. For example:

HELP("C:\\EXCEL\\CUSTHELP.DOC\!101")

In Microsoft Excel for the Macintosh, the following macro formula
displays the Help topic numbered 101 in the file CUSTOM HELP:

HELP("CUSTOM HELP\!101")

If the custom Help file is not in the current folder, specify the full
path along with the name of the file. For example:

HELP("HARD DISK:EXCEL:HELP:CUSTOM HELP\!101")

Return to [top](#H)

# HIDE

Equivalent to clicking the Hide command on the Window menu. Hides the
active window.

**Syntax**

**HIDE**( )

**Tip**   Hiding windows can speed up your macros. You can switch to
hidden windows with the ACTIVATE function. You can continue to use
functions that refer to specific sheets, such as FORMULA and the GET
functions, even when those sheets are hidden.

**Related Function**

UNHIDE   Displays a hidden window

Return to [top](#H)

# HIDE.DIALOG

Closes the dialog box that has the current focus.

**Syntax**

**HIDE.DIALOG**(cancel\_logical)

Cancel\_logical    is a logical value that specifies whether to close
the dialog box and validate any edit boxes. If FALSE, the dialog box is
closed and edit boxes are validated, or checked to determine if they
contain a valid data type. If TRUE, the dialog box is closed and the
edit boxes are not validated.

**Remarks**

If the edit box does not contain a valid data type when the dialog box
is closed, the dialog will remain open. For example, if the edit box is
supposed to contain integer values, and a text value is entered, the
dialog box will not close. This applies to only those dialog boxes that
must be closed before any further user action can happen.

**Examples**

HIDE.DIALOG(FALSE) closes the dialog box and checks to see if the edit
box contains a valid data type (validated)

**Related Functions**

EDITBOX.PROPERTIES   Sets the properties of an edit box on a worksheet
or dialog sheet

SHOW.DIALOG   Runs a dialog on a dialog sheet

Return to [top](#H)

# HIDE.OBJECT

Hides or displays the specified object.

**Syntax**

**HIDE.OBJECT**(object\_id\_text, hide)

Object\_id\_text    is the name and number, or number alone, of the
object, as text, as it appears in the reference area when the object is
selected. The name of the object is also the text returned by the
CREATE.OBJECT function, so object\_id\_text can be a reference to a cell
containing CREATE.OBJECT. To give the name of more than one object, use
the following format for object\_id\_text:

"oval 3, text 2, arc 5"

 

If object\_id\_text is omitted, the function operates on all selected
objects. If no object is selected or if the object specified by
object\_id\_text does not exist, HIDE.OBJECT returns the \#VALUE\! error
value.

Hide    is a logical value that specifies whether to hide or display the
specified object. If hide is TRUE or omitted, Microsoft Excel hides the
object; if FALSE, Microsoft Excel displays the object.

**Remarks**

Objects are not automatically selected after they are unhidden.

**Examples**

The following macro formula hides the selected object:

HIDE.OBJECT(, TRUE)

The following macro formula displays the object named Oval 3:

HIDE.OBJECT("Oval 3", FALSE)

The following macro formula displays the three specified objects:

HIDE.OBJECT("oval 3, text 2, arc 5", FALSE)

**Related Functions**

CREATE.OBJECT   Creates an object

DISPLAY   Controls how an object is displayed

Return to [top](#H)

# HISTOGRAM

Calculates individual and cumulative percentages for a range of data and
a corresponding range of data bins.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**HISTOGRAM**(**inprng**, outrng, binrng, pareto, chartc, chart, labels)

**HISTOGRAM**?(inprng, outrng, binrng, pareto, chartc, chart, labels)

Inprng    is the input range.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Binrng    is an optional set of numbers that define the bin ranges. The
values must be in ascending order. The values are interpreted as more
than value A up to value B, more than value B up to value C, and so on.
One additional bin is created for values for less than the minimum value
specified in binrng.

Pareto    is a logical value.

  - > If pareto is TRUE, data in the output table is presented in both
    > ascending-bin order and descending-frequency order.

  - > If pareto is FALSE or omitted, data in the output table is
    > presented in ascending-bin order only.

>  

Chartc    is a logical value. If chartc is TRUE, HISTOGRAM generates a
cumulative percentages column in the output table. If both chartc and
chart are TRUE, HISTOGRAM also includes a cumulative percentage line in
the histogram chart. If omitted, chartc is FALSE.

Chart    is a logical value. If chart is TRUE, HISTOGRAM generates a
histogram chart in addition to the output table. If omitted, chart is
FALSE.

Labels    is a logical value.

  - > If labels is TRUE, then the first row or column of inprng contains
    > labels.

  - > If labels is FALSE or omitted, all cells in inprng are considered
    > data. Microsoft Excel generates appropriate data labels for the
    > output table.

Return to [top](#H)

# HLINE

Scrolls through the active window by a specific number of columns.
Returns the \#VALUE\! error value if the active sheet is a chart.

**Syntax**

**HLINE**(**num\_columns**)

Num\_columns    is the number of columns in the active worksheet or
macro sheet you want to scroll through horizontally.

  - > If num\_columns is positive, HLINE scrolls to the right.

  - > If num\_columns is negative, HLINE scrolls to the left.

  - > Num\_columns must be between -256 and 256, inclusive.

>  

**Example**

The following function scrolls the active window by one-half window to
the right:

HLINE(GET.WINDOW(15)/2)

**Related Functions**

HPAGE   Horizontally scrolls through the active window one window at a
time

HSCROLL   Horizontally scrolls through a sheet by percentage or by
column number

VLINE   Vertically scrolls through the active window by rows

VPAGE   Vertically scrolls through the active window one window at a
time

VSCROLL   Vertically scrolls through a sheet by percentage or by row
number

Return to [top](#H)

# HPAGE

Horizontally scrolls through the active window one window at a time. Use
HPAGE to change the displayed area of a worksheet or macro sheet.

**Syntax**

**HPAGE**(**num\_windows**)

Num\_windows    specifies the number of windows to scroll through the
active window horizontally. A window is defined as the number of visible
columns. If three columns are visible in the window, HPAGE scrolls
through in increments of three columns.

  - > If num\_windows is positive, HPAGE scrolls to the right.

  - > If num\_windows is negative, HPAGE scrolls to the left.

>  

**Related Functions**

HLINE   Horizontally scrolls through the active window by columns

HSCROLL   Horizontally scrolls through a worksheet by percentage or by
column number

VLINE   Vertically scrolls through the active window by rows

VPAGE   Vertically scrolls through the active window one window at a
time

VSCROLL   Vertically scrolls through a worksheet by percentage or by row
number

Return to [top](#H)

# HSCROLL

Horizontally scrolls through the active sheet by percentage or by column
number.

**Syntax**

**HSCROLL**(**position**, col\_logical)

Position    specifies the column you want to scroll to. Position can be
an integer representing the column number or a fraction or percentage
representing the horizontal position of the column in the sheet. If
position is 0, HSCROLL scrolls through your sheet to its leftmost edge.
If position is 1, HSCROLL scrolls through your sheet to its rightmost
edge. For charts that do not size with the window, use a fraction or
percentage.

Col\_logical    is a logical value specifying how the function scrolls.

  - > If col\_logical is TRUE, HSCROLL scrolls through the sheet to
    > column position.

  - > If col\_logical is FALSE or omitted, then HSCROLL scrolls through
    > the sheet to the horizontal position represented by the fraction
    > position.

>  

**Remarks**

  - > To scroll to a specific column n, either use HSCROLL(n, TRUE) or
    > use HSCROLL(n/256). To scroll to column 38, for example, use
    > HSCROLL(38, TRUE) or HSCROLL(38/256).

  - > If you are recording a macro and move the scroll box several times
    > in a row, the recorder only records the final location of the
    > scroll box, omitting any intermediate steps. Remember that
    > scrolling does not change the active cell or the selection.

>  

**Related Functions**

HLINE   Horizontally scrolls through the active window by columns

HPAGE   Horizontally scrolls through the active window one window at a
time

VLINE   Vertically scrolls through the active window by rows

VPAGE   Vertically scrolls through the active window one window at a
time

VSCROLL   Vertically scrolls through a sheet by percentage or row number

Return to [top](#H)

# IF

Used with ELSE, ELSE.IF, and END.IF to control which formulas in a macro
are executed. There are two syntax forms of the IF function. The
following syntax can be used on macro sheets only; use it when you want
your macro to branch to a particular set of functions based on the
outcome of a logical test. The worksheet form of this function can be
used on worksheets and macro sheets.

**Syntax**

**IF**(**logical\_test**)

Logical\_test    is a logical value that IF uses to determine which
functions to carry out next—that is, where to branch.

  - > If logical\_test is TRUE, Microsoft Excel carries out the
    > functions between the IF function and the next ELSE, ELSE.IF, or
    > END.IF function. Instructions between ELSE.IF or ELSE and END.IF
    > are not carried out.

  - > If logical\_test is FALSE, Microsoft Excel immediately branches to
    > the next ELSE.IF, ELSE, or END.IF function.

  - > If logical\_test produces an error, the macro halts.

# Tips

  - > Use IF with ELSE, ELSE.IF, and END.IF when you want to perform
    > multiple actions based on a condition. This method is preferable
    > to using GOTO because it makes your macros more structured.

  - > If your macro ends with an error at a cell containing this form of
    > the IF function, make sure there is a corresponding END.IF
    > function.

**Example**

The following macro runs the macro CompleteEntry if the user clicks OK:

IF(ALERT("Are you done with this entry?", 1), CompleteEntry(), )

**Tip**   You can indent formulas in a macro. To indent a formula, type
as many spaces as you want between the equal sign and the first letter
of the formula.

**Related Functions**

ELSE   Specifies an action to take if an IF function returns FALSE

ELSE.IF   Specifies an action to take if an IF or another ELSE.IF
function returns FALSE

END.IF   Ends a group of macro functions started with an IF statement

ERROR   Specifies what action to take if an error occurs while a macro
is running

Return to [top](#H)

# INITIATE

Opens a dynamic data exchange (DDE) channel to an application and
returns the number of the open channel. Once you have opened a channel
to another application with INITIATE, you can use EXECUTE and SEND.KEYS
to control the other application from a Microsoft Excel macro.
(SEND.KEYS is available only with Microsoft Excel for Windows.) If
INITIATE is successful, it returns the number of the open channel. All
the subsequent DDE macro functions use this number to specify the
channel. If INITIATE is unsuccessful, FALSE is returned.

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for this function.

**Syntax**

**INITIATE**(**app\_text, topic\_text**)

App\_text    is the DDE name of the application with which you want to
begin a DDE session, in text form. The form of app\_text depends on the
application you are accessing. The DDE name of Microsoft Excel, for
example, is "Excel".

Topic\_text    describes something, such as a document or a record in a
database, in the application that you are accessing; the form of
topic\_text depends on the application you are accessing. Microsoft
Excel accepts the names of the current documents as topic\_text, as well
as the name "System".

**Remarks**

  - > You can specify an instance of an application by appending the
    > application's task ID number to the app\_text argument. If you
    > start an application by using the EXEC function, EXEC returns the
    > task ID number for that instance of the application.

  - > If more than one instance of an application is running and you do
    > not specify which instance you would like to open a channel to,
    > INITIATE displays a dialog box from which you can choose the
    > instance you want. You can prevent this dialog box from appearing
    > by disabling or redirecting errors with the ERROR function.

>  

**Example**

The following macro formula opens a channel to the document named MEMO
in the application named WORD:

INITIATE("WORD", "MEMO")

**Related Functions**

POKE   Sends data to another application

REQUEST   Returns data from another application

TERMINATE   Closes a channel to another application

EXECUTE   Carries out a command in another application

EXEC   Starts a separate program

Return to [top](#H)

# INPUT

Displays a dialog box for user input. Returns the information entered in
the dialog box. Use INPUT to display a simple dialog box for the user to
enter information to be used in a macro.

The dialog box has an OK and a Cancel button. If you click the OK
button, INPUT returns the default value specified or the value typed in
the edit box. If you click the Cancel button, INPUT returns FALSE.

**Syntax**

**INPUT**(**message\_text**, type\_num, title\_text, default, x\_pos,
y\_pos, help\_ref)

Message\_text    is the text to be displayed in the dialog box.
Message\_text must be enclosed in quotation marks.

Type\_num    is a number specifying the type of data to be entered.

|               |               |
| ------------- | ------------- |
| **Type\_num** | **Data type** |
| 0             | Formula       |
| 1             | Number        |
| 2             | Text          |
| 4             | Logical       |
| 8             | Reference     |
| 16            | Error         |
| 64            | Array         |

You can also use the sum of the allowable data types for type\_num. For
example, for an input box that can accept formulas, text, or numbers,
set type\_num equal to 3 (the sum of 0, 1, and 2, which are the type
specifiers for formula, number, and text). If type\_num is omitted, it
is assumed to be 2.

  - > If type\_num is 0, INPUT returns the formula in the form of text,
    > for example, "=2\*PI()/360".

  - > To enter a formula, include an equal sign at the beginning of the
    > formula; otherwise the formula is returned as text.

  - > If the formula contains references, they are returned as
    > R1C1-style references, for example, "=RC\[-1\]\*(1+R1C1)".

  - > If type\_num is 8, INPUT returns an absolute reference to the
    > specified cells.

  - > If you enter a single-cell reference in the dialog box, the value
    > in that cell is returned by the INPUT function.

  - > If the information entered in the dialog box is not of the correct
    > data type, Microsoft Excel attempts to convert it to the specified
    > type. If the information can't be converted, Microsoft Excel
    > displays an error message.

>  

Title\_text    is text specifying a title to be displayed in the title
bar of the dialog box. If title\_text is omitted, it is assumed to be
"Input".

Default    specifies a value to be shown in the edit box when the dialog
box is initially displayed. If default is omitted, the edit box is left
empty.

X\_pos, y\_pos    specify the horizontal and vertical position, in
points, of the dialog box. A point is 1/72nd of an inch. If either or
both arguments are omitted, the dialog box is centered in the
corresponding direction.

Help\_ref    is a reference to a custom online Help topic in a text
file, in the form "filename\!topic\_number".

  - > If help\_ref is present, a Help button appears in the lower-right
    > corner of the dialog box. Clicking the Help button starts Help and
    > displays the specified topic.

  - > If help\_ref is omitted, no Help button appears.

  - > Help\_ref must be given as text.

>  

For more information about custom Help topics, see HELP.

**Remarks**

Relative references entered in formulas in the INPUT dialog box are
relative to the active cell at the time the INPUT function is
calculated. If you are using the reference entered into the dialog box
in a cell other than the active cell, it may not refer to the cells you
intend it to. For example, if the active cell is A3 and you enter the
formula "=A1+A2" in an INPUT dialog box, intending to add the values in
those cells, and then use the FORMULA function to enter the formula in
cell B3, the formula in cell B3 will read "=B1+B2" because you gave a
relative reference. You can use FORMULA.CONVERT to solve this problem.

**Examples**

In Microsoft Excel for Windows, the following macro formula displays the
following dialog box:

INPUT("Enter the inflation rate:", 1, "Inflation Rate", , , ,
"CUSTHELP.DOC\!101")

If you then enter 12%, INPUT returns the value 0.12.

In Microsoft Excel for the Macintosh, the following macro formula
displays the following dialog box:

INPUT("Enter the inflation rate:", 1, "Inflation Rate", , , , "CUSTOM
HELP\!101")

If you then enter 12%, INPUT returns the value 0.12.

If the active cell is C2 and you enter the formula =B2\*(1+$A$1) in
response to the following macro formula:

INPUT("Enter your monthly increase formula:", 0)

INPUT returns "=RC\[-1\]\*(1+R1C1)"

If you select the range $A$2:$A$8 in the INPUT dialog box:

REFTEXT|USA|002|001|001|common|UREFTEXT(INPUT("Please make your
selection.", 8)) returns R2C1:R8C1

**Related Functions**

ALERT   Displays a dialog box and a message

DIALOG.BOX   Displays a custom dialog box

FORMULA.CONVERT   Changes the style and type of references in a formula

HELP   Displays a custom Help topic

Return to [top](#H)

# INSERT

Inserts a blank cell or range of cells or pastes cells from the
Clipboard into a sheet. Shifts the selected cells to accommodate the new
ones. The size and shape of the inserted range are the same as those of
the current selection.

**Syntax**

**INSERT**(shift\_num)

**INSERT**?(shift\_num)

Shift\_num    is a number from 1 to 4 specifying which way to shift the
cells. If an entire row or column is selected, shift\_num is ignored. If
shift\_num is omitted, Microsoft Excel shifts cells in the logical
direction based on the selection.

|                |                     |
| -------------- | ------------------- |
| **Shift\_num** | **Direction**       |
| 1              | Shift cells right   |
| 2              | Shift cells down    |
| 3              | Shift entire row    |
| 4              | Shift entire column |

**Remarks**

If you have just cut or copied information to the Clipboard, INSERT
performs both an insert and a paste operation. First, Microsoft Excel
inserts new blank cells into the sheet; then, Microsoft Excel pastes
information from the Clipboard into the newly inserted cells. If you
have used the INSERT function in macros written for Microsoft Excel
version 2.2 or earlier, make sure you consider this feature when you use
your old macros with later versions of Microsoft Excel.

**Related Functions**

COPY   Copies and pastes data or objects

CUT   Cuts or moves data or objects

EDIT.DELETE   Removes cells from a sheet

PASTE   Pastes cut or copied data

Return to [top](#H)

# INSERT.OBJECT

Equivalent to choosing the Object command from the Insert menu, and then
selecting an object type and choosing the OK button. Creates an embedded
object whose source data is supplied by another application. Also starts
an application of the appropriate class for the specified object type.

**Syntax**

**INSERT.OBJECT**(**object\_class**, file\_name, link\_logical,
display\_icon\_logical, icon\_file, icon\_number, icon\_label)

**INSERT.OBJECT**?(object\_class, file\_name, link\_logical,
display\_icon\_logical, icon\_file, icon\_number, icon\_label)

Object\_class    is a text string containing the classname for the
object you want to create.

  - > Object\_class is the classname corresponding to the Object Type
    > selection in the Insert Object dialog box.

  - > For more information about object classnames, consult the
    > documentation for your source application to see how it supports
    > object linking and embedding (OLE).

>  

File\_name    is a text string specifying the file from which to create
an OLE object.

Link\_logical    is a logical value indicating whether the new object
based on file\_name should be linked to file\_name. If it is not linked,
the object is created as a copy or the file. Link\_logical is ignored if
file\_name is not specified. If link\_logical is FALSE or omitted, then
no link is established.

Display\_icon\_logical    is a logical value corresponding to the
Display as Icon checkbox. If it is FALSE or omitted, then the regular
picture for the object is displayed. If it is TRUE, then the icon
icon\_number found in icon\_file is displayed with the label
icon\_label. If display\_icon\_logical is not TRUE, then icon\_file,
icon\_number, and icon\_label are ignored.

Icon\_file    is the name of the file where the icon to display is
located.

Icon\_number    is the number of the icon within icon\_file that should
be used.

Icon\_label    is a text string indicating a label to display beneath
the icon. If the parameter is an empty string ("") or is omitted, no
label is displayed.

**Remarks**

  - > If INSERT.OBJECT starts another application, your macro pauses.
    > Your macro resumes when you return to Microsoft Excel.

  - > Although you will not normally use Microsoft Excel class names in
    > a Microsoft Excel macro, you may need them in macros written for
    > other applications. Microsoft Excel uses classnames
    > "Excel.Sheet.5" and "Excel.Chart.5".

>  

**Related Function**

EDIT.OBJECT   Edits an object

Return to [top](#H)

# INSERT.PICTURE

Equivalent to clicking the Picture command on the Insert menu. This
function is available for Microsoft Excel for Windows only .

**Syntax**

**INSERT.PICTURE**(file\_name, filter\_number)

**INSERT.PICTURE**?(file\_name, filter\_number)

File\_name    is the name, as text, of the file containing the picture
that you want to insert into your workbook.

Filter\_number    is a number specifying which converter Microsoft Excel
will use to open the file.

|                   |                                      |
| ----------------- | ------------------------------------ |
| **Convert\_type** | **Converter and filename extension** |
| 1                 | Windows Bitmaps (bmp)                |
| 2                 | Windows Metafile (wmf)               |
| 3                 | DrawPerfect (wpg)                    |
| 4                 | Micrografix Designer/Draw (drw)      |
| 5                 | AutoCAD Format 2-D (dxf)             |
| 6                 | HP Graphics Language (hgl)           |
| 7                 | Computer Graphics Metafile (cgm)     |
| 8                 | Encapsulated Postscript (eps)        |
| 9                 | Tagged Image Format (tif)            |
| 10                | PC PaintBrush (pcx)                  |
| 11                | Lotus 1-2-3 Graphic (pic)            |
| 12                | AutoCAD Plot Files (plt)             |
| 13                | Macintosh PICT (pct)                 |

Return to [top](#H)

# INSERT.TITLE

Attaches text to various parts of a chart.

**Syntax**

2-D charts

**INSERT.TITLE**(chart, y\_primary, x\_primary, y\_secondary,
x\_secondary)

3-D charts

**INSERT.TITLE**(chart, z\_primary, x\_primary, y\_primary)

Chart    is a logical value specifying whether to attach a title to the
chart.

Y\_primary    is a logical value specifying whether to attach a title to
the value (y) axis of a 2-D chart or the series (y) axis of a 3-D chart.

X\_primary    is a logical value specifying whether to attach a title to
the category (x) axis of the chart.

Z\_primary    is a logical value specifying whether to attach a title to
the value (z) axis of a 3-D chart.

Y\_secondary    is a logical value specifying whether to attach a title
to the second value (y) axis of a chart containing more than one chart
type.

X\_secondary    is a logical value specifying whether to attach a title
to the second category (x) axis of a chart containing more than one
chart type.

**Remarks**

To change the text in a selected title, use the FORMULA function.

**Related Function**

FORMULA   Enters formulas in a chart

Return to [top](#H)

# JUSTIFY

Equivalent to clicking the Justify command on the Fill submenu of the
Edit menu. Rearranges the text in a range so that it fills the range
evenly.

**Syntax**

**JUSTIFY**()

**Related Function**

ALIGNMENT   Aligns the contents of the selected cells

Return to [top](#H)

# LABEL.PROPERTIES

Sets the accelerator property of the label and group box controls on a
worksheet or dialog sheet.

**Syntax**

**LABEL.PROPERTIES**(accel\_text, accel\_text2, 3d\_shading)

**LABEL.PROPERTIES**?(accel\_text, accel\_text2, 3d\_shading)

Accel\_text    is a text string containing the character to use as the
label's accelerator key on a dialog sheet. The character is matched
against the text of the control, and the first matching character is
underlined. When the user presses ALT+accel\_text in Microsoft Excel for
Windows or COMMAND+accel\_text in Microsoft Excel for the Macintosh, the
control is clicked. This argument is ignored for controls on worksheets.

Accel\_text2    is a text string containing the second accelerator key
on a dialog sheet. This argument is for only Far East versions of
Microsoft Excel.

3d\_shading    is a logical value that specifies whether the list box
appears as 3-D. If TRUE, the list box will appear as 3-D. If FALSE or
omitted, the list box will not be 3-D. This argument is available for
only worksheets.

**Related Functions**

CHECKBOX.PROPERTIES   Sets various properties of check box and option
box controls

SCROLLBAR.PROPERTIES   Sets the properties of the scroll bar and spinner
controls

PUSHBUTTON.PROPERTIES   Sets the properties of the push button control

Return to [top](#H)

# LAST.ERROR

Returns the reference to the cell where the last macro sheet error
occurred. If no error has occurred, LAST.ERROR returns the \#N/A error
value. Use LAST.ERROR in conjunction with the REFTEXT function to
quickly locate errors.

**Syntax**

**LAST.ERROR**( )

**Related Function**

ERROR   Specifies what action to take if an error is encountered while a
macro is running

Return to [top](#H)

# LEGEND

Adds a legend to or removes a legend from a chart. This is also
equivalent to clicking the Legend button on the Chart toolbar when a
chart is active.

**Syntax**

**LEGEND**(logical)

Logical    is a logical value specifying which command LEGEND is
equivalent to.

  - > If logical is TRUE or omitted, LEGEND is equivalent to the Legend
    > command on the Insert menu.

  - > If logical is FALSE, LEGEND is equivalent to the Delete command on
    > the Edit menu.

  - > If logical is FALSE and the active chart has no legend, LEGEND
    > takes no action.

>  

**Related Function**

FORMAT.LEGEND   Determines the position and orientation of the legend on
a chart

Return to [top](#H)

# LINE.PRINT

Prints the active worksheet using methods compatible with those of Lotus
1-2-3. LINE.PRINT does not use the Microsoft Windows printer drivers.
Unless you have a specific need for the LINE.PRINT function, use the
PRINT function instead.

**Note   **This function is only available in Microsoft Excel for
Windows.

**Syntax 1**

Go, Line, Page, Align, and Clear

**LINE.PRINT**(**command**, file, append)

**Syntax 2**

Worksheet settings

**LINE.PRINT**(**command**, setup\_text, leftmarg, rightmarg, topmarg,
botmarg, pglen, formatted)

**Syntax 3**

Global settings

**LINE.PRINT**(**command**, setup\_text, leftmarg, rightmarg, topmarg,
botmarg, pglen, wait, autolf, port, update)

Command    is a number corresponding to the command you want LINE.PRINT
to carry out. For syntax 2, command must be 5. For syntax 3, command
must be 6.

|             |                                           |
| ----------- | ----------------------------------------- |
| **Command** | **Command that is carried out**           |
| 1           | Go                                        |
| 2           | Line                                      |
| 3           | Page                                      |
| 4           | Align                                     |
| 5           | Worksheet settings                        |
| 6           | Global settings (saved in EXCEL5.INI)     |
| 7           | Clear (change to current global settings) |

File    is the name of a file to which you want to print. If omitted,
Microsoft Excel prints to the printer port determined by the current
global settings.

Append    is a logical value specifying whether to append text to file.
If TRUE, the file you are printing is appended to file; if FALSE or
omitted, the file you are printing overwrites the contents of file.

Setup\_text    is text that includes a printer initialization sequence
or other control codes to prepare your printer for printing. If omitted,
no setup text is used.

Leftmarg    is the size of the left margin measured in characters from
the left side of the page. If omitted, it is assumed to be 4.

Rightmarg    is the size of the right margin measured in characters from
the left side of the page. If omitted, it is assumed to be 76.

Topmarg    is the size of the top margin measured in lines from the top
of the page. If omitted, it is assumed to be 2.

Botmarg    is the size of the bottom margin measured in lines from the
bottom of the page. If omitted, it is assumed to be 2.

Pglen    is the number of lines on one page. If omitted, it is assumed
to be 66 (11 inches with 6 lines per inch). If you're using an HP
LaserJet or compatible printer, set pglen to 60 (the printer reserves
six lines).

Formatted    is a logical value specifying whether to format the output.
If TRUE or omitted, the output is formatted; if FALSE, it is not
formatted.

Wait    is a logical value specifying whether to wait after printing a
page. If TRUE, Microsoft Excel waits; if FALSE or omitted, Microsoft
Excel continues printing.

Autolf    is a logical value specifying whether your printer has
automatic line feeding. If TRUE, Microsoft Excel prints lines normally;
if FALSE or omitted, Microsoft Excel sends an additional line feed
character after printing each line.

Port    is a number from 1 to 8 specifying which port to use when
printing.

|              |                             |
| ------------ | --------------------------- |
| **Port**     | **Port used when printing** |
| 1 or omitted | LPT1                        |
| 2            | COM1                        |
| 3            | LPT2                        |
| 4            | COM2                        |
| 5            | LPT1                        |
| 6            | LPT2                        |
| 7            | LPT3                        |
| 8            | LPT4                        |

Update    is a logical value specifying whether to update and save
global settings. If TRUE, the settings are saved in the EXCEL5.INI file;
if FALSE or omitted, the global settings are not saved.

**Remarks**

The default values for print settings on your worksheet are determined
by the current global settings.

**Example**

The following macro formula prints the currently defined print area to
the currently defined printer port:

LINE.PRINT(1)

**Related Function**

PRINT   Prints the active sheet

Return to [top](#H)

# LINK.COMBO

Links an edit box and a list box control into a linked combination box
group. The resulting linked controls track each other's selection and
contents. Linked edit and list box combinations are similar to an
editable drop-down list box, except that the list box is permanently
visible and dropped down.

**Syntax**

**LINK.COMBO**(**link\_logical**)

Link\_logical    is a logical value that specifies whether the controls
are linked or unlinked. If TRUE, the controls will become linked. If
FALSE, the controls will be unlinked.

**Remarks**

To use this function, first select the list box and edit box to be
linked or unlinked. You can do this with SELECT("list box 1,Edit box
2").

**Examples**

LINK.COMBO(FALSE) will unlink a list box and an edit box.

**Related Functions**

ADD.LIST.ITEM   Adds an item in a list box or drop-down control on a
worksheet or dialog sheet control

SELECT.LIST.ITEM   Selects an item in a list box or in a group box

Return to [top](#H)

# LINK.FORMAT

Links the number format of the selected data label to the worksheet cell
or range containing the data label text.

**Syntax**

**LINK.FORMAT**( )

Return to [top](#H)

# LINKS

Returns, as a horizontal array of text values, the names of all
workbooks referred to by external references in the workbook specified.
Use LINKS with OPEN.LINKS to open supporting workbooks.

**Syntax**

**LINKS**(document\_text, type\_num)

Document\_text    is the name of a workbook, including its path. If
document\_text is omitted, LINKS operates on the active workbook. If the
workbook specified by document\_text is not open, LINKS returns the
\#N/A error value.

Type\_num    is a number from 1 to 6 specifying the type of linked
workbooks to return.

|               |                                                |
| ------------- | ---------------------------------------------- |
| **Type\_num** | **Returns**                                    |
| 1 or omitted  | Microsoft Excel link                           |
| 2             | DDE/OLE link (Microsoft Excel for Windows)     |
| 3             | Reserved                                       |
| 4             | Not applicable                                 |
| 5             | Publisher (Microsoft Excel for the Macintosh)  |
| 6             | Subscriber (Microsoft Excel for the Macintosh) |

**Remarks**

  - > If the active workbook contains no external references, LINKS
    > returns the \#N/A error value.

  - > With the INDEX function, you can select individual workbook names
    > from the array for use in other functions that take workbook names
    > as arguments.

  - > The names of the workbook are always returned in alphabetic order.
    > If supporting workbooks are open, LINKS returns the names of the
    > workbooks; if supporting workbooks are closed, LINKS includes the
    > full path of each workbook.

  - > If type\_num is 5 or 6, LINKS returns a two-row array in which the
    > first row contains the edition name and the second row contains
    > the reference.

>  

**Examples**

If a chart named Chart1 is open and contains links to workbook Data1 and
Data2, and the LINKS function shown below is entered as an array into a
two-cell horizontal range:

LINKS("Chart1") equals "Data1" in the first cell of the range and
"Data2" in the second cell.

In Microsoft Excel for Windows, if the chart named VARIANCE.XLS is open
and contains data series that refer to workbook named BUDGET.XLS and
ACTUAL.XLS, then:

OPEN.LINKS(LINKS("VARIANCE.XLS")) opens BUDGET.XLS and ACTUAL.XLS.

In Microsoft Excel for the Macintosh, if the workbook named SALES 1991
is open and contains references to the workbook WEST SALES, SOUTH SALES,
and EAST SALES, then:

OPEN.LINKS(LINKS("SALES 1991")) opens WEST SALES, SOUTH SALES, and EAST
SALES.

**Related Functions**

CHANGE.LINK   Changes supporting workbook links

GET.LINK.INFO   Returns information about a link

OPEN.LINKS   Opens specified supporting workbook

UPDATE.LINK   Updates a link to another workbook

Return to [top](#H)

# LISTBOX.PROPERTIES

Sets the properties of a list box and drop-down controls on a worksheet
or dialog sheet.

**Syntax**

**LISTBOX.PROPERTIES**(range, link, drop\_size, multi\_select,
3d\_shading)

**LISTBOX.PROPERTIES**?(range, link, drop\_size, multi\_select,
3d\_shading)

Range    is the cell range that the initial contents of the list box are
taken from. If blank (empty text), the list box is initially unfilled.

Link    is the cell on the sheet to which the list box is linked, and
indicates the numeric position of the currently selected item in the
list box. Whenever an item in the list box is selected, its numeric
position is entered into the linked cell on the sheet.

Drop\_size    is the number of lines shown when a drop-down control is
dropped. This value is ignored when applied to a non-drop-down list box.

Multi\_select    is a number that specifies the selection mode of the
list box. Zero is single selection. 1 is simple multi-select. 2 is
extended multi-select.

3d\_shading    is a logical value that specifies whether the list box
appears as 3-D. If TRUE, the list box will appear as 3-D. If FALSE or
omitted, the list box will not be 3-D. This argument is available for
only worksheets.

**Related Functions**

ADD.LIST.ITEM   Adds an item in a list box or drop-down control on a
worksheet or dialog sheet control

SELECT.LIST.ITEM   Selects an item in a list box or in a group box

CHECKBOX.PROPERTIES   Sets various properties of check box and option
box controls

SCROLLBAR.PROPERTIES   Sets the properties of the scroll bar and spinner
controls

PUSHBUTTON.PROPERTIES   Sets the properties of the push button control

Return to [top](#H)

# LIST.NAMES

Equivalent to clicking the Paste command on the Name submenu of the
Insert menu and selecting the Paste List option button. Lists all names
(except hidden names) defined in your workbook. LIST.NAMES also lists
the cells to which the names refer; whether a macro corresponding to a
particular name is a command macro or a custom function; the shortcut
key for each command macro; and the category of the custom functions.

**Syntax**

**LIST.NAMES**( )

**Remarks**

  - > If the current selection is a single cell or five or more columns
    > wide, LIST.NAMES pastes all five types of information about
    > worksheet names into five columns. The first column contains cell
    > names. The second column contains the corresponding cell
    > references. The third column contains the number 1 if the name
    > refers to a custom function, the number 2 if it refers to a
    > command macro, or 0 if it refers to anything else. The fourth
    > column lists the shortcut keys for command macros. The fifth
    > column contains a category name for custom functions or the number
    > of the built-in category.

  - > If the selection includes fewer than five columns, LIST.NAMES
    > omits the information that would have been pasted into the other
    > columns.

  - > When you use LIST.NAMES, Microsoft Excel completely replaces the
    > contents of the cells it pastes into.

>  

**Related Functions**

GET.DEF   Returns a name matching a definition

GET.NAME   Returns the definition of a name

NAMES   Returns the names defined in a workbook

Return to [top](#H)

# MACRO.OPTIONS

Equivalent to clicking the Options button in the Macro dialog box, which
appears when you click the Macros command (Tools menu, Macro submenu).

**Syntax**

MACRO.OPTIONS(**macro\_name**, description, menu\_on, menu\_text,
shortcut\_on, shortcut\_key, function\_category, status\_bar\_text,
help\_id, help\_file)

Macro\_name    is the name of the macro that you want to set options
for, including the name of the workbook and sheet containing the macro.

Description    is the description of the macro displayed in the Macro
dialog box.

Menu\_on    is a logical value indicating whether a menu item is
automatically added for this macro. If TRUE, menu\_text must be
specified. If FALSE or omitted, no menu item is added. If the macro
already has a menu item, setting this argument to FALSE removes the menu
item.

Menu\_text    is the text of the menu item to be added for the macro.
Ignored unless menu\_on is TRUE.

Shortcut\_on    is a logical value indicating whether a shortcut key is
assigned to the macro. If TRUE, shortcut\_key must be specified. If
FALSE or omitted, no shortcut key is assigned. If the macro already has
a shortcut key, setting this argument to FALSE removes the shortcut key.

Shortcut\_key    is the letter of the shortcut key for the macro.
Ignored if shortcut\_key is FALSE.

Function\_category    is the number of the category in the Paste
Function dialog box that the macro is assigned to. Categories are
numbered starting at 1 for the category at the top of the list in the
Paste Function dialog box.

Status\_bar\_text    the text displayed in the status bar when a menu
item or toolbar button assigned to this macro is clicked on. Be sure to
enclose the text in quotes.

Help\_id    is the numerical ID for the help topic associated with this
macro and any related menu items or toolbar buttons.

Help\_file    is the pathname of the help file for the macro.

Return to [top](#H)

# MAIL.ADD.MAILER

Equivalent to clicking the Add Mailer command on the Mail submenu of the
File menu. Adds a new PowerTalk mailer to the active workbook. Use this
command to add addressing or subject information to a workbook that you
want to send to another user.

**Note   **This function is available on Macintosh computers with
Microsoft Excel and Apple PowerTalk only.

**Syntax**

**MAIL.ADD.MAILER**( )

**Remarks**

If there is already a mailer, this command fails and returns the
\#VALUE\! error value.

**Related Function**

MAIL.DELETE.MAILER   Deletes an existing mailer from the active workbook

Return to [top](#H)

# MAIL.DELETE.MAILER

Equivalent to clicking the Delete Mailer command on the Mail submenu of
the File menu. Deletes an existing mailer from the active workbook.

**Note   **This function is available on Macintosh computers with
Microsoft Excel and Apple PowerTalk only.

**Syntax**

**MAIL.DELETE.MAILER**( )

**Remarks**

If there is no mailer, returns the \#VALUE\! error value.

**Related Function**

MAIL.ADD.MAILER   Adds a new PowerTalk mailer to the active workbook

Return to [top](#H)

# MAIL.EDIT.MAILER

Equivalent to clicking the Mailer button when mailer is attached to the
current workbook. Allows you to edit a PowerTalk mailer attached to the
active workbook

**Note   **This function is available on Macintosh computers with
Microsoft Excel and Apple PowerTalk only.

**Syntax**

**MAIL.EDIT.MAILER**(to\_recipients, cc\_recipients, bcc\_recipients,
subject, enclosures, which\_address)

**MAIL.EDIT.MAILER**?(to\_recipients, cc\_recipients, bcc\_recipients,
subject, enclosures, which\_address)

To\_recipients    is the name of the person to whom you want to send the
mail. The name should be given as text. To specify more than one name,
give the list of names as an array.

Cc\_recipients    is the name of those recipients to be carbon copied. A
single name should be given as text. To specify more than one name, give
the list of names as an array.

Bcc\_recipients    is the name of the recipients to be added as blind
carbon copies.

Subject     is a text string containing the subject text for the mail
messages.

Enclosures    is an array of strings specifying enclosures as file
names.

Which\_address    indicates which type of address to use, as a text
string, specifying the address type for all recipients. For example,
"Fax".

**Remarks**

If there is no mailer, returns the \#VALUE\! error value.

**Related Functions**

MAIL.DELETE.MAILER   Adds a new PowerTalk mailer to the active workbook

MAIL.ADD.MAILER   Adds a new PowerTalk mailer to the active workbook

Return to [top](#H)

# MAIL.FORWARD

Equivalent to clicking the Forward command on the Mail submenu of the
File menu. Creates a new mailer to replace the previous version and
brings up the mailer dialog.

**Note   **This function is available on Macintosh computers with
Microsoft Excel and Apple PowerTalk only.

**Syntax**

**MAIL.FORWARD**( )

**Remarks**

  - > Returns the \#VALUE\! error value or \#N/A if the current workbook
    > has no mailer.

  - > This function is available only when the current workbook is open
    > and has been received by PowerTalk with a piece of mail to
    > forward.

**Related Functions**

MAIL.EDIT.MAILER   Allows you to edit a PowerTalk mailer attached to the
active workbook

MAIL.DELETE.MAILER   Deletes a new PowerTalk mailer to the active
workbook

MAIL.ADD.MAILER   Adds a new PowerTalk mailer to the active workbook

Return to [top](#H)

# MAIL.LOGOFF

Ends the current mail session.

**Important   **To use MAIL.LOGOFF in Microsoft Excel for Windows, you
must be using a mail client that supports the Messaging Applications
Programming Interface (MAPI) or Vendor-Independent Messaging (VIM). The
function is available for only Microsoft Excel for Windows.

**Syntax**

**MAIL.LOGOFF**( )

**Remarks**

Returns TRUE if the session was ended, or \#VALUE\! if there was no
session.

Return to [top](#H)

# MAIL.LOGON

Starts a mail session.

**Important   **To use MAIL.LOGON in Microsoft Excel for Windows, you
must be using a mail client that supports the Messaging Applications
Programming Interface (MAPI) or Vendor-Independent Messaging (VIM). The
function is available for only Microsoft Excel for Windows.

**Syntax**

**MAIL.LOGON**(name\_text, password\_text, download\_logical)

**MAIL.LOGON**?(name\_text, password\_text, download\_logical)

Name\_text    is the username of the mail account or Microsoft Exchange
profile name. If omitted, prompts for username.

Password\_text    is the password for the account. If omitted, prompts
for password. Ignored when the dialog box form is used. This argument is
ignored in Microsoft Exchange.

Download\_logical    specifies whether to download new mail. Use TRUE to
download new mail; use FALSE or leave blank to skip downloading new
mail.

**Remarks**

Returns FALSE if you cancel the dialog box or \#VALUE\! if you can't log
on.

If you omit both name\_text and password\_text, Microsoft Excel tries to
log on using an existing mail session.

**Related Function**

MAIL.LOGOFF   Ends the current mail session

Return to [top](#H)

# MAIL.NEXT.LETTER

Equivalent to clicking the Next Letter command on the Mail submenu of
the File menu. Opens the oldest unread Microsoft Excel workbook from the
In Tray as a new window.

**Note   **This function is available on Macintosh computers with
Microsoft Excel and Apple PowerTalk only.

**Syntax**

**MAIL.NEXT.LETTER**( )

**Remarks**

Returns \#VALUE\! on error, and \#N/A if there are no more letters in
the In Tray to open.

**Related Functions**

MAIL.EDIT.MAILER   Allows you to edit a PowerTalk mailer attached to the
active workbook

MAIL.DELETE.MAILER   Adds a new PowerTalk mailer to the active workbook

MAIL.ADD.MAILER   Adds a new PowerTalk mailer to the active workbook

Return to [top](#H)

# MAIL.REPLY

Equivalent to clicking the Reply command on the Mail submenu of the File
menu. Replies to the sender of the current letter.

**Note   **This function is available on Macintosh computers with
Microsoft Excel and Apple PowerTalk only.

**Syntax**

**MAIL.REPLY**( )

**Remarks**

  - > Returns the \#VALUE\! error value or \#N/A if the current workbook
    > has no mailer.

  - > The letter must currently be open.

**Related Functions**

MAIL.EDIT.MAILER   Allows you to edit a PowerTalk mailer attached to the
active workbook

MAIL.DELETE.MAILER   Deletes a new PowerTalk mailer to the active
workbook

MAIL.ADD.MAILER   Adds a new PowerTalk mailer to the active workbook

Return to [top](#H)

# MAIL.REPLY.ALL

Equivalent to clicking the Reply All command on the Mail submenu of the
File menu. Replies to the sender and all recipients of the current
letter.

**Note   **This function is available on Macintosh computers with
Microsoft Excel and Apple PowerTalk only.

**Syntax**

**MAIL.REPLY.ALL**( )

**Remarks**

Returns the \#VALUE\! error value or \#N/A if the current workbook has
no mailer.

**Related Functions**

MAIL.EDIT.MAILER   Allows you to edit a PowerTalk mailer attached to the
active workbook

MAIL.DELETE.MAILER   Deletes a new PowerTalk mailer to the active
workbook

MAIL.ADD.MAILER   Adds a new PowerTalk mailer to the active workbook

Return to [top](#H)

# MAIL.SEND.MAILER

Equivalent to clicking the Send Mailer command on the Mail submenu of
the File menu. Sends a PowerTalk mailer.

**Note   **This function is available on Macintosh computers with
Microsoft Excel and Apple PowerTalk only.

**Syntax**

**MAIL.SEND.MAILER**( )

Return to [top](#H)

# MAIN.CHART

Equivalent to clicking the Main Chart command on the Format menu when a
chart sheet is active in Microsoft Excel version 2.2 or earlier. This
function is included only for macro compatibility.

**Syntax**

**MAIN.CHART**(**type\_num**, stack, 100, vary, overlap, drop, hilo,
overlap%,  
cluster, angle)

**MAIN.CHART**?(type\_num, stack, 100, vary, overlap, drop, hilo,
overlap%,  
cluster, angle)

Return to [top](#H)

# MAIN.CHART.TYPE

Equivalent to clicking the Main Chart Type command on the Chart menu in
Microsoft Excel for the Macintosh version 1.5 or earlier. This function
is included only for macro compatibility.

**Syntax**

**MAIN.CHART.TYPE**(**type\_num**)

Return to [top](#H)

# MCORREL

Returns a correlation matrix that measures the correlation between two
or more data sets that are scaled to be independent of the unit of
measurement.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**MCORREL**(**inprng**, outrng, grouped, labels)

**MCORREL**?(inprng, outrng, grouped, labels)

Inprng    is the input range.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Grouped    is a text character that indicates whether the data in the
input range is organized by row or column.

  - > If grouped is "C" or omitted, then the data is organized by
    > column.

  - > If grouped is "R", then the data is organized by row.

>  

Labels    is a logical value that describes where the labels are located
in the input range, as shown in the following table:

|                  |             |                                                   |
| ---------------- | ----------- | ------------------------------------------------- |
| **Labels**       | **Grouped** | **Labels are in**                                 |
| TRUE             | "C"         | First row of the input range.                     |
| TRUE             | "R"         | First column of the input range.                  |
| FALSE or omitted | (ignored)   | No labels. All cells in the input range are data. |

**Related Function**

MCOVAR   Returns the covariance between two or more data sets

Return to [top](#H)

# MCOVAR

Returns a covariance matrix that measures the covariance between two or
more data sets.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**MCOVAR**(**inprng**, outrng, grouped, labels)

**MCOVAR**?(inprng, outrng, grouped, labels)

Inprng    is the input range.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Grouped    is a text character that indicates whether the data in the
input range is organized by row or column.

  - > If grouped is "C" or omitted, then the data is organized by
    > column.

  - > If grouped is "R", then the data is organized by row.

>  

Labels    is a logical value that describes where the labels are located
in the input range, as shown in the following table:

|                  |             |                                                   |
| ---------------- | ----------- | ------------------------------------------------- |
| **Labels**       | **Grouped** | **Labels are in**                                 |
| TRUE             | "C"         | First row of the input range                      |
| TRUE             | "R"         | First column of the input range                   |
| FALSE or omitted | (ignored)   | No labels. All cells in the input range are data. |

**Related Function**

MCORREL   Returns the correlation coefficient of two or more data sets
that are scaled to be independent of the unit of measurement

Return to [top](#H)

# MENU.EDITOR

This function should not be used in Microsoft Excel 97 or later because
the Menu Editor is only available in Microsoft Excel 95 and Microsoft
Excel version 5.0.

Return to [top](#H)

# MERGE.STYLES

Equivalent to clicking the Merge button in the Style dialog box, which
appears when you click the Style command on the Format menu. Merges all
the styles from another workbook into the active workbook. Use
MERGE.STYLES when you want to import styles from another sheet in
another workbook.

**Syntax**

**MERGE.STYLES**(document\_text)

Document\_text is the name of a sheet in a workbook from which you want
to merge styles into the active workbook.

**Remarks**

  - > If any styles from the workbook being merged have the same name as
    > styles in the active workbook, a dialog box appears asking if you
    > want to replace the existing definitions of the styles with the
    > "merged" definitions of the styles. If you click the Yes button,
    > all the definitions are replaced; if you click the No button, all
    > the original definitions in the active workbook are retained.

  - > When you move a sheet with styles to another workbook with styles,
    > any styles with identical names but conflicting definitions have
    > the sheet name added to the style name.

>  

**Related Functions**

DEFINE.STYLE   Creates or changes a cell style

DELETE.STYLE   Deletes a cell style

Return to [top](#H)

# MESSAGE

Displays and removes messages in the message area of the status bar.
MESSAGE is useful for displaying text that doesn't need a response, such
as descriptions of commands in user-defined menus.

**Syntax**

**MESSAGE**(**logical**, text)

Logical    is a logical value specifying whether to display or remove a
message.

  - > If logical is TRUE, Microsoft Excel displays text in the message
    > area of the status bar.

  - > If logical is FALSE, Microsoft Excel removes any messages, and the
    > status bar is returned to normal (that is, command help messages
    > are displayed).

Text    is the message you want to display in the status bar. If text is
"" (empty text), Microsoft Excel removes any messages currently
displayed in the status bar.

**Remarks**

  - > Only one message can be displayed in the status bar at a time.
    > Messages are always displayed in the same place.

  - > MESSAGE works the same way whether the status bar is displayed or
    > not. You can, for example, use MESSAGE while the status bar isn't
    > displayed. As soon as you display the status bar, you see your
    > message.

  - > If you display any message (even empty text) and don't remove it
    > with MESSAGE(FALSE), that message is displayed until you quit
    > Microsoft Excel.

  - > You can also use the ALERT function to get the user's attention;
    > however, this interrupts the macro and requires the user's
    > intervention before the macro can continue.

>  

**Example**

The following lines in a macro display a message warning that you must
wait for a moment while the macro calls a subroutine.

MESSAGE(TRUE, "One moment please...")

**Related Functions**

ALERT   Displays a dialog box and a message

BEEP   Sounds a tone

Return to [top](#H)

# MOVE

Equivalent to moving a window by dragging its title bar in Microsoft
Excel version 3.0 or earlier. MOVE is also equivalent to choosing the
Move command from the Control menu in Microsoft Windows. This function
is included only for macro compatibility and will be converted to
WINDOW.MOVE when you open older macro sheets. For more information, see
WINDOW.MOVE.

**Syntax**

**MOVE**(x\_pos, y\_pos, window\_text)

**MOVE**?(x\_pos, y\_pos, window\_text)

**Related Functions**

WINDOW.MOVE   Sizes a window

WINDOW.SIZE   Moves a window

Return to [top](#H)

# MOVEAVG

Projects values in a forecast period, based on the average value of the
variable over a specific number of preceding periods.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**MOVEAVG**(**inprng**, outrng, interval, stderrs, chart, labels)

**MOVEAVG**?(inprng, outrng, interval, stderrs, chart, labels)

Inprng    is the input range.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Interval    is the number of values to include in the moving average. If
omitted, interval is 3.

Stderrs    is a logical value.

  - > If stderrs is TRUE, standard error values are included in the
    > output table.

  - > If stderrs is FALSE or omitted, standard errors are not included
    > in the output table.

>  

Chart    is a logical value.

  - > If chart is TRUE, then MOVEAVG generates a chart for the actual
    > and forecast values.

  - > If chart is FALSE or omitted, the chart is not generated.

>  

Labels    is a logical value.

  - > If labels is TRUE, then the first row or column of inprng contains
    > labels.

  - > If labels is FALSE or omitted, all cells in inprng are considered
    > data. Microsoft Excel generates appropriate data labels for the
    > output table.

Return to [top](#H)

# MOVE.TOOL

Moves or copies a button from one toolbar to another.

**Syntax**

**MOVE.TOOL**(**from\_bar\_id, from\_bar\_position**, to\_bar\_id,
to\_bar\_position, copy, width)

From\_bar\_id    specifies the number or name of a toolbar from which
you want to move or copy the button. For detailed information, see the
description of bar\_id in ADD.TOOL.

From\_bar\_position    specifies the current position of the button
within the toolbar. From\_bar\_position starts with 1 at the left side
(if horizontal) or at the top (if vertical).

To\_bar\_id    specifies the number or name of a toolbar to which you
want to move or paste the button. For detailed information, see the
description of bar\_id in ADD.TOOL. To\_bar\_id is optional if you are
moving a button within the same toolbar.

To\_bar\_position    specifies where you want to move or paste the
button within the toolbar. To\_bar\_position starts with 1 at the left
side (if horizontal) or at the top (if vertical). To\_bar\_position is
optional if you are only adjusting the width of a drop-down list.

Copy    is a logical value specifying whether to copy the button. If
copy is TRUE, the button is copied; if FALSE or omitted, the button is
moved.

Width    is the width, measured in points, of a drop-down list. If the
button you are moving is not a drop-down list, width is ignored.

**Related Functions**

ADD.TOOL   Adds one or more buttons to a toolbar

COPY.TOOL   Copies a button face to the Clipboard

GET.TOOL   Returns information about a button or buttons on a toolbar

Return to [top](#H)

# NAMES

Returns, as a horizontal array of text, the specified names defined in
the specified workbook. The returned array lists the names in alphabetic
order. Use NAMES instead of LIST.NAMES when you want to return the names
to the macro sheet instead of to the active worksheet.

**Syntax**

**NAMES**(document\_text, type\_num, match\_text)

Document\_text    is text that specifies the workbook whose names you
want returned. If document\_text is omitted, it is assumed to be the
active workbook.

Type\_num    is a number from 1 to 3 that specifies whether to include
hidden names in the returned array.

|                     |                   |
| ------------------- | ----------------- |
| **If type\_num is** | **NAMES returns** |
| 1 or omitted        | Normal names only |
| 2                   | Hidden names only |
| 3                   | All names         |

Match\_text    is text that specifies the names you want returned and
can include wildcard characters. If match\_text is omitted, all names
are returned.

**Remarks**

  - > Hidden names are defined using the DEFINE.NAME macro function and
    > do not appear in the Paste Name, Define Name, or Go To dialog
    > boxes.

  - > NAMES returns a horizontal array, so you will normally enter this
    > function as an array in several horizontal cells or define a name
    > to refer to the array that NAMES returns. If you want the names in
    > a vertical array instead, use the TRANSPOSE function.

  - > You can use the COLUMNS function to count the number of entries in
    > the horizontal array.

>  

**Example**

The following macro formula returns all names on the active workbook
starting with the letter P.

NAMES(, 3, "P\*")

**Related Functions**

DEFINE.NAME   Defines a name on the active worksheet or macro sheet

DELETE.NAME   Deletes a name

GET.DEF   Returns a name matching a definition

GET.NAME   Returns the definition of a name

LIST.NAMES   Lists names and their associated information

SET.NAME   Defines a name as a value

Return to [top](#H)

# NEW

Equivalent to clicking the New command on the File menu. Creates a new
Microsoft Excel workbook or opens a template.

**Syntax**

**NEW**(type\_num, xy\_series, add\_logical)

**NEW**?(type\_num, xy\_series, add\_logical)

Type\_num    specifies the type of workbook to create, as shown in the
following table. Type\_num is most often 5 or quoted text; other values
are mainly for compatibility with Microsoft Excel version 4.0.

|               |                                                                               |
| ------------- | ----------------------------------------------------------------------------- |
| **Type\_num** | **Workbook**                                                                  |
| Omitted       | New workbook with a single worksheet of the same type as the active worksheet |
| 1             | New workbook with one worksheet                                               |
| 2             | New workbook with one chart based on the current selection                    |
| 3             | New workbook with one macro sheet                                             |
| 4             | New workbook with one international macro sheet                               |
| 5             | New workbook with 16 worksheets or based on the default workbook              |
| 6             | New workbook with one Visual Basic module                                     |
| 7             | New workbook with one dialog sheet                                            |
| Quoted text   | Template.                                                                     |

Xy\_series    is a number from 0 to 3 that specifies how data is
arranged in a chart.

|                |                                                                                         |
| -------------- | --------------------------------------------------------------------------------------- |
| **Xy\_series** | **Result**                                                                              |
| 0              | Displays a dialog box if the selection is ambiguous.                                    |
| 1 or omitted   | The first row/column is the first data series.                                          |
| 2              | The first row/column contains the category (x) axis labels.                             |
| 3              | The first row/column contains the x-values; the created chart is an xy (scatter) chart. |

Add\_logical    specifies whether or not to add the sheet type to the
open workbook. If add\_logical is TRUE, the sheet type is inserted
before the current sheet; if FALSE or omitted, it is not inserted. This
argument is for compatibility with Microsoft Excel version 4.0.

Add\_logical is ignored if type\_num is 5.

**Remarks**

You can also use NEW to create new sheets from templates that exist in
the startup directory or folder, using for type\_num the text that
appears in the File New list box. To create new sheets from any template
that is not in the start-up directory, use the OPEN function.

**Related Functions**

NEW.WINDOW   Creates a new window for an existing worksheet or macro
sheet

OPEN   Opens a workbook

Return to [top](#H)

# NEW.WINDOW

Equivalent to clicking the New Window command on the Window menu.
Creates a new window for the active workbook.

**Syntax**

**NEW.WINDOW**( )

After you use NEW.WINDOW, use the WINDOW.MOVE, WINDOW.SIZE, and
ARRANGE.ALL functions to size and position the new window.

**Related Functions**

ARRANGE.ALL   Arranges all displayed windows to fill the workspace and
synchronizes windows for scrolling

WINDOW.MOVE   Moves a window

WINDOW.SIZE   Changes the size of a window

Return to [top](#H)

# NEXT

Ends a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop and continues
carrying out the current macro with the formula that follows the NEXT
function.

**Syntax**

**NEXT**( )

**Related Functions**

FOR   Starts a FOR-NEXT loop

FOR.CELL   Starts a FOR.CELL-NEXT loop

WHILE   Starts a WHILE-NEXT loop

Return to [top](#H)

# NOTE

Equivalent to choosing the Comment command from the Insert menu. Creates
a comment or replaces text characters in a comment.

**Syntax**

**NOTE**(add\_text, cell\_ref, start\_char, num\_chars)

**NOTE**?( )

Add\_text    is text of up to 255 characters you want to add to a
comment. Add\_text must be enclosed in quotation marks.

  - > If add\_text is omitted, it is assumed to be "" (empty text).

>  

Cell\_ref    is the cell to which you want to add the comment text. If
cell\_ref is omitted, add\_text is added to the active cell's comment.

Start\_char    is the number of the character at which you want
add\_text to be added. If start\_char is omitted, it is assumed to be 1.
If there is no existing comment, start\_char is ignored.

Num\_chars    is the number of characters that you want to replace in
the comment. If num\_chars is omitted, it is assumed to be equal to the
length of the comment.

**Remarks**

  - > NOTE returns the number of the last character entered in the
    > comment. This is useful if you want to know how many characters
    > are in the text string.

  - > The dialog-box form of this function, NOTE?, takes no arguments.

  - > NOTE() deletes the comment attached to the active cell.

>  

To find out if a cell has a comment attached to it, use GET.CELL.

**Related Function**

GET.NOTE   Returns characters from a comment

Return to [top](#H)

# OBJECT.PROPERTIES

Determines how the selected object or objects are attached to the cells
beneath them and whether they are printed. The way an object is attached
to the cells beneath it affects how the object is moved or sized
whenever you move or size the cells.

**OBJECT.PROPERTIES**(placement\_type, print\_object)

**OBJECT.PROPERTIES**?(placement\_type, print\_object)

Placement\_type    is a number from 1 to 3 specifying how to attach the
selected object or objects. If placement\_type is omitted, the current
status is unchanged.

|                           |                                                              |
| ------------------------- | ------------------------------------------------------------ |
| **If placement\_type is** | **The selected object is**                                   |
| 1                         | Moved and sized with cells.                                  |
| 2                         | Moved but not sized with cells.                              |
| 3                         | Free-floating—it is not affected by moving and sizing cells. |

Print\_object    is a logical value specifying whether to print the
selected object or objects. If TRUE or omitted, the objects are printed;
if FALSE, they are not printed.

**Remarks**

If an object is not selected, OBJECT.PROPERTIES interrupts the macro and
returns the \#VALUE\! error value.

**Related Functions**

CREATE.OBJECT   Creates an object

FORMAT.MOVE   Moves the selected object

FORMAT.SIZE   Changes the size of the selected object

Return to [top](#H)

# OBJECT.PROTECTION

Changes the protection status of the selected object.

**Syntax**

**OBJECT.PROTECTION**(locked, lock\_text)

**OBJECT.PROTECTION**?(locked, lock\_text)

Locked    is a logical value that determines whether the selected object
is locked or unlocked. If locked is TRUE, Microsoft Excel locks the
object; if FALSE, Microsoft Excel unlocks the object.

Lock\_text    is a logical value that determines whether text in a text
box or button can be changed. Lock\_text applies only if the object is a
text box, button, or worksheet control. If lock\_text is TRUE or
omitted, text cannot be changed; if FALSE, text can be changed.

**Remarks**

  - > You cannot lock or unlock an individual object with
    > OBJECT.PROTECTION when protection is selected for objects in the
    > Protect Sheet dialog box.

  - > If an object is not selected, the function returns the \#VALUE\!
    > error value and halts the macro.

  - > In order for an object to be protected, you must use the
    > PROTECT.DOCUMENT(, , , TRUE) function after changing the object's
    > status with OBJECT.PROTECTION.

>  

**Related Functions**

PROTECT.DOCUMENT   Controls protection for the active worksheet

WORKBOOK.PROTECT   Controls protection for the active workbook

Return to [top](#H)

# ON.DATA

Runs a specified macro when another application sends data to a
particular workbook via dynamic data exchange (DDE) or via Publish and
Subscribe on the Macintosh. Workbook links to other applications are
called remote references.

**Syntax**

**ON.DATA**(document\_text, macro\_text)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for this function.

Document\_text    is the name of the sheet to which remote data will be
sent or the name of the source of the remote data.

  - > If document\_text is the name of the remote data source, it must
    > be in the form app|topic\!item. You can use the form app|topic to
    > include all items for a particular topic, or app| to specify an
    > app alone, but you must include the | to indicate that you are
    > specifying the name of a data source.

  - > If document\_text is omitted, the macro specified by macro\_text
    > is run whenever remote data is sent to any sheet not already
    > assigned to another ON.DATA function.

  - > In Microsoft Excel for the Macintosh, document\_text can also be
    > the name of a published edition file. Unless the file is in the
    > current folder, document\_text must include the complete path.

>  

Macro\_text    is the name of, or an R1C1-style reference to, a macro
that you want to run when data comes into the workbook or from the
source specified by document\_text. The name or reference must be in
text form.

  - > If macro\_text is omitted, the ON.DATA function is turned off for
    > the specified workbook or source.

>  

**Remarks**

  - > ON.DATA remains in effect until you either clear it or quit
    > Microsoft Excel. You can clear ON.DATA by specifying
    > document\_text and omitting the macro\_text argument.

  - > If the macro sheet containing macro\_text is closed when data is
    > sent to document\_text, an error is returned.

  - > If the incoming data causes recalculation, Microsoft Excel first
    > runs the macro macro\_text and then performs the recalculation.

>  

**Examples**

In Microsoft Excel for Windows, the following macro formula runs the
macro AddOrders when data is sent to the sheet New in the workbook
ORDERSDB.XLS:

ON.DATA("\[ORDERSDB.XLS\]New", "AddOrders")

In Microsoft Excel for the Macintosh, the following macro formula runs
the macro beginning at cell R2C3 when data is sent to the sheet North in
the workbook SALES DATABASE:

ON.DATA("\[SALES DATABASE\]North", "R2C3")

**Related Functions**

ERROR   Specifies what action to take if an error is encountered while a
macro is running

INITIATE   Opens a channel to another application

ON.ENTRY   Runs a macro when data is entered

ON.RECALC   Runs a macro when a workbook is recalculated

Return to [top](#H)

# ON.DOUBLECLICK

Runs a macro when you double-click any cell or object on the specified
sheet or macro sheet or double-click any item on the specified chart.

**Syntax**

**ON.DOUBLECLICK**(sheet\_text, macro\_text)

Sheet\_text    is a text value specifying the name of a sheet in a
workbook. If sheet\_text is omitted, the macro is run whenever you
double-click any sheet not specified by a previous ON.DOUBLECLICK
formula. Sheet\_text must be in the form "\[book1\]sheet1".

Macro\_text    is the name of, or an R1C1-style reference to, a macro
you want to run when you double-click the sheet specified by
sheet\_text. The name or reference must be in text form. If macro\_text
is omitted, double-clicking reverts to its normal behavior, and any
macros assigned by previous ON.DOUBLECLICK functions are turned off.

**Remarks**

  - > ON.DOUBLECLICK overrides Microsoft Excel's normal double-click
    > behavior, such as displaying cell notes, displaying the Patterns
    > dialog box, or allowing editing cell text directly in cells.

  - > To determine what cell, object, or chart item has been
    > double-clicked, use a CALLER function in the macro specified by
    > macro\_text.

  - > ON.DOUBLECLICK does not affect objects to which ASSIGN.TO.OBJECT
    > macros have already been assigned. Use ON.DOUBLECLICK (TRUE) to
    > make Microsoft Excel carry out the action that would normally
    > occur if you double-click on the current selection.

>  

**Related Functions**

ASSIGN.TO.OBJECT   Assigns a macro to an object

ON.WINDOW   Runs a macro when you switch to a window

Return to [top](#H)

# ON.ENTRY

Runs a macro when you enter data into any cell on the specified sheet.

**Syntax**

**ON.ENTRY**(sheet\_text, macro\_text)

Sheet\_text    is a text value specifying the name of a sheet in a
workbook. If sheet\_text is omitted, the macro is run whenever you enter
data into any sheet or macro sheet.

Macro\_text    is the name of, or an R1C1-style reference to, a macro
you want to run when you enter data into the sheet specified by
sheet\_text. The name or reference must be in text form. If macro\_text
is omitted, entering data reverts to its normal behavior, and any macros
assigned by previous ON.ENTRY functions are turned off.

**Remarks**

  - > The macro is run only when you enter data in a cell, not when you
    > use edit commands or macro functions.

  - > To determine what cell had data entered into it, use a CALLER
    > function in the macro specified by macro\_text.

>  

**Related Functions**

ENTER.DATA   Turns Data Entry mode on or off

ON.RECALC   Runs a macro when a workbook is recalculated

Return to [top](#H)

# ON.KEY

Runs a specified macro when a particular key or key combination is
pressed.

**Syntax**

**ON.KEY**(**key\_text**, macro\_text)

Key\_text    can specify any single key, or any key combined with ALT,
CTRL, or SHIFT, or any combination of those keys (in Microsoft Excel for
Windows) or COMMAND, CTRL, OPTION, or SHIFT or any combination of those
keys (in Microsoft Excel for the Macintosh). Each key is represented by
one or more characters, such as "a" for the character a, or "{ENTER}"
for the ENTER key.

To specify characters that aren't displayed when you press the key, such
as ENTER or TAB, use the codes shown in the following table. Each code
in the table represents one key on the keyboard.

|                        |                         |
| ---------------------- | ----------------------- |
| **Key**                | **Code**                |
| BACKSPACE              | "{BACKSPACE}" or "{BS}" |
| BREAK                  | "{BREAK}"               |
| CAPS LOCK              | "{CAPSLOCK}"            |
| CLEAR                  | "{CLEAR}"               |
| DELETE or DEL          | "{DELETE}" or "{DEL}"   |
| DOWN                   | "{DOWN}"                |
| END                    | "{END}"                 |
| ENTER (numeric keypad) | "{ENTER}"               |
| ENTER                  | "\~" (tilde)            |
| ESC                    | "{ESCAPE} or {ESC}"     |
| HELP                   | "{HELP}"                |
| HOME                   | "{HOME}"                |
| INS                    | "{INSERT}"              |
| LEFT                   | "{LEFT}"                |
| NUM LOCK               | "{NUMLOCK}"             |
| PAGE DOWN              | "{PGDN}"                |
| PAGE UP                | "{PGUP}"                |
| RETURN                 | "{RETURN}"              |
| RIGHT                  | "{RIGHT}"               |
| SCROLL LOCK            | "{SCROLLLOCK}"          |
| TAB                    | "{TAB}"                 |
| UP                     | "{UP}"                  |
| F1 through F15         | "{F1}" through "{F15}"  |

In Microsoft Excel for Windows, you can also specify keys combined with
SHIFT and/or CTRL and/or ALT. In Microsoft Excel for the Macintosh, you
can also specify keys combined with SHIFT and/or CTRL and/or OPTION
and/or COMMAND. To specify a key combined with another key or keys, use
the following table.

|                     |                             |
| ------------------- | --------------------------- |
| **To combine with** | **Precede the key code by** |
| SHIFT               | "+" (plus sign)             |
| CTRL                | "^" (caret)                 |
| ALT or OPTION       | "%" (percent sign)          |
| COMMAND             | "\*" (asterisk)             |

To assign a macro to one of the special characters (+, ^, %, and so on),
enclose the character in brackets. For example, ON.KEY("^{+}",
"InsertItem") assigns a macro named InsertItem to the key sequence
CTRL+PLUS SIGN.

Macro\_text    is the name of a macro that you want to run when
key\_text is pressed. The reference must be in text form.

  - > If macro\_text is "" (empty text), nothing happens when key\_text
    > is pressed. This form of ON.KEY disables the normal meaning of
    > keystrokes in Microsoft Excel.

  - > If macro\_text is omitted, key\_text reverts to its normal meaning
    > in Microsoft Excel, and any special key assignments made with
    > previous ON.KEY functions are cleared.

>  

**Remarks**

  - > ON.KEY remains in effect until you clear it or quit Microsoft
    > Excel. You can clear ON.KEY by specifying key\_text and omitting
    > the macro\_text argument.

  - > If the macro sheet containing macro\_text is closed when you press
    > key\_text, an error is returned.

  - > If another macro is running when you press key\_text, macro\_text
    > will not run.

>  

**Examples**

Suppose you wanted the key combination SHIFT+CTRL+RIGHT to run the macro
Print. You use the following macro formula:

ON.KEY("+^{RIGHT}", "Print")

To return SHIFT+CTRL+RIGHT to its normal meaning, you would use the
following macro formula:

ON.KEY("+^{RIGHT}")

To disable SHIFT+CTRL+RIGHT altogether, you would use the following
macro formula:

ON.KEY("+^{RIGHT}", "")

**Related Functions**

CANCEL.KEY   Disables macro interruption

ERROR   Specifies what action to take if an error is encountered while a
macro is running

SEND.KEYS   Sends a key sequence to an application

Return to [top](#H)

# ON.RECALC

Runs a macro when a specific sheet is recalculated. Use ON.RECALC to
perform an operation on a sheet each time the sheet is recalculated,
such as checking that a certain condition is still being met.

**Syntax**

**ON.RECALC**(sheet\_text, macro\_text)

Sheet\_text    is the name of a sheet, given as text. If sheet\_text is
omitted, the macro is run whenever any open sheet not specified by a
previous ON.RECALC formula is recalculated. Only one ON.RECALC formula
can be run for each recalculation.

Macro\_text    is the name of, or an R1C1-style reference to, a macro
you want to run when the sheet specified by sheet\_text is recalculated.
The name or reference must be in text form. The macro will be run each
time the sheet is recalculated until ON.RECALC is cleared. If
macro\_text is omitted, recalculating reverts to its normal behavior,
and any macros assigned by previous ON.RECALC functions are turned off.

**Remarks**

A macro specified to be run by ON.RECALC is not run by actions taken by
other macros. For example, a macro specified by ON.RECALC will not be
run after the CALCULATE.NOW function is carried out, but will be run if
you change data in a sheet set to calculate automatically or choose the
Calc Now button in the Calculation panel of the Options dialog box,
which appears when you choose the Options command from the Tools menu.

**Examples**

In Microsoft Excel for Windows, the following macro formula specifies
that the macro Printer on the macro sheet AUTOREPT.XLS be run when the
worksheet named REPORT.XLS is recalculated:

ON.RECALC("REPORT.XLS", "AUTOREPT.XLS\!Printer")

In Microsoft Excel for the Macintosh, the following macro formula turns
off ON.RECALC for the workbook named SALES:

ON.RECALC("SALES")

**Related Functions**

CALCULATE.DOCUMENT   Calculates the active sheet only

CALCULATE.NOW   Calculates all open workbooks immediately

CALCULATION   Controls calculation settings

ON.ENTRY   Runs a macro when data is entered

Return to [top](#H)

# ON.SHEET

Starts a macro whenever the specified sheet is activated from another
sheet.

**Syntax**

**ON.SHEET**(sheet\_text, macro\_text, activate\_logical)

Sheet\_text    is the name of the sheet that triggers a macro when it is
activated, in the form "\[Book1\]Sheet1". If omitted, then when any
sheet in any book is activated, macro\_text will run.

Macro\_text    is the name of the macro to run when the specified sheet
is activated. If omitted, then the triggering of a macro on the
specified sheet is cancelled.

Activate\_logical    is a logical value that specifies if the macro is
run when the sheet is activated or deactivated. If TRUE or omitted, the
macro is run when the sheet is activated. If FALSE, the macro is run
when the sheet is deactivated.

**Examples**

ON.SHEET("\[STORE.XLS\]Sheet1","WeeklyCalc") will run the macro
"WeeklyCalc" when "\[STORE.XLS\]Sheet1" is activated.

ON.SHEET(,"WeeklyCalc") runs "WeeklyCalc" when any sheet in the book is
activated.

ON.SHEET("\[STORE.XLS\]Sheet1") cancels the triggering of "WeeklyCalc"
when Sheet1 in the book STORE.XLS is activated.

ON.SHEET("\[STORE.XLS\]","WeeklyCalc") runs "WeeklyCalc" when any sheet
in the book STORE.XLS is activated

**Related Function**

ON.WINDOW   Runs a specified macro when you switch to a particular
window.

Return to [top](#H)

# ON.TIME

Runs a macro at a specified time. Use ON.TIME to run a macro at a
specific time of day or after a specified period has passed.

**Syntax**

**ON.TIME**(**time, macro\_text**, tolerance, insert\_logical)

Time    is the time and date, given as a serial number, at which the
macro is to be run. If time does not include a date (that is, if time is
a serial number less than 1), the macro is run the next time this time
occurs.

Macro\_text    is the name of, or an R1C1-style reference to, a macro to
run at the specified time and every subsequent day at that time.

Tolerance    is the time and date, given as a serial number, that you
are willing to wait until and still have the macro run. For example, if
Microsoft Excel is not in Ready, Copy, Cut, or Find mode at time,
because another macro is running, but is in Ready mode 15 seconds later,
and tolerance is set to time plus 30 seconds, the macro specified by
macro\_text will run. If Microsoft Excel was not in Ready mode within 30
seconds, the macro would not run. If tolerance is omitted, it is assumed
to be infinite.

Insert\_logical    is a logical value specifying whether you want every
day macro\_text to run at time. Use insert\_logical when you want to
clear a previously set ON.TIME formula. If insert\_logical is TRUE or
omitted, the macro specified by macro\_text is carried out at time. If
insert\_logical is FALSE and macro\_text is not set to run at time,
ON.TIME returns the \#VALUE error value.

**Examples**

The following macro formula runs a macro called Test at 5:00:00 P.M.
every day when Microsoft Excel is in Ready mode:

ON.TIME("5:00:00 PM", "Test")

The following macro formula runs a macro called Test 5 seconds after the
formula is evaluated:

ON.TIME(NOW()+"00:00:05", "Test")

The following macro formula runs a macro called Test 10 seconds after
the formula is evaluated. If Microsoft Excel is not in Ready mode at
that time (because it is in Edit mode, for example), the tolerance
argument specifies 5 seconds of additional time to wait to run the
macro. If Microsoft Excel is still not in Ready mode at that time,
macro\_text is not run.

ON.TIME(NOW()+"00:00:10", "Test", NOW()+"00:00:15")

Return to [top](#H)

# ON.WINDOW

Runs a specified macro when you switch to a particular window.

**Syntax**

**ON.WINDOW**(window\_text, macro\_text)

Window\_text    is the name of a window in the form of text. If
window\_text is omitted, ON.WINDOW starts the macro whenever you switch
to any window, except for windows that are named in other ON.WINDOW
statements.

Macro\_text    is the name of, or an R1C1-style reference to, a macro to
run when you switch to window\_text. If macro\_text is omitted,
switching to window\_text no longer runs the previously specified macro.

**Remarks**

  - > A macro specified to be run by ON.WINDOW is not run when other
    > macros switch to the window or when a command to switch to a
    > window is received through a DDE channel. Instead, ON.WINDOW
    > responds to a user's actions, such as clicking a window with the
    > mouse, clicking the Copy command on the Edit menu, and so on.

  - > If a sheet or macro sheet has an Auto\_Activate or
    > Auto\_Deactivate macro defined for it, those macros will be run
    > after the macro specified by ON.WINDOW.

>  

**Examples**

In Microsoft Excel for Windows, the following macro formula runs the
macro beginning at cell R1C2 when you switch to the window MAIN.XLS:

ON.WINDOW("MAIN.XLS", "R1C2")

The following macro formula stops the macro from running when you switch
to MAIN.XLS:

ON.WINDOW("MAIN.XLS")

In Microsoft Excel for the Macintosh, the following macro formula runs
the macro named ShowAlert when you switch to the window MAIN WINDOW:

ON.WINDOW("MAIN WINDOW", "ShowAlert")

The following macro formula stops the macro from running when you switch
to MAIN WINDOW:

ON.WINDOW("MAIN WINDOW")

**Related Functions**

GET.WINDOW   Returns information about a window

ON.KEY   Runs a macro when a specified key is pressed

ON.SHEET   Triggers a macro whenever the specified sheet is activated
from another sheet

WINDOWS   Returns the names of all open windows

Return to [top](#H)

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

File\_text    is the name, as text, of the workbook you want to open.
File\_text can include a drive and path, and can be a network pathname.
In the dialog-box form in Microsoft Excel for Windows, file\_text can
include an asterisk (\*) to represent any sequence of characters and a
question mark (?) to represent any single character.

Update\_links    specifies whether and how to update external and remote
references. If update\_links is omitted, Microsoft Excel displays a
message asking if you want to update links.

|                         |                                                |
| ----------------------- | ---------------------------------------------- |
| **If update\_links is** | **Then Microsoft Excel**                       |
| 0                       | Updates neither external nor remote references |
| 1                       | Updates external references only               |
| 2                       | Updates remote references only                 |
| 3                       | Updates external and remote references         |

**Note   **When you are opening a file in WKS, WK1, or WK3 format, the
update\_links argument specifies whether Microsoft Excel generates
charts from any graphs attached to the WKS, WK1, or WK3 file.

|                         |                |
| ----------------------- | -------------- |
| **If update\_links is** | **Charts are** |
| 0                       | Not created    |
| 2                       | Created        |

Read\_only    corresponds to the Read Only check box in the Open dialog
box. If read\_only is TRUE, the workbook can be modified but changes
cannot be saved; if FALSE or omitted, changes to the workbook can be
saved.

Format    specifies what character to use as a delimiter when opening
text files. If format is omitted, Microsoft Excel uses the current
delimiter setting.

|                  |                             |
| ---------------- | --------------------------- |
| **If format is** | **Values are separated by** |
| 1                | Tabs                        |
| 2                | Commas                      |
| 3                | Spaces                      |
| 4                | Semicolons                  |
| 5                | Nothing                     |
| 6                | Custom characters           |

Prot\_pwd    is the password, as text, required to unprotect a protected
file. If prot\_pwd is omitted and file\_text requires a password, the
Password dialog box is displayed. Passwords are case-sensitive.
Passwords are not recorded when you open a workbook and supply the
password with the macro recorder on.

Write\_res\_pwd    is the password, as text, required to open a
read-only file with full write privileges. If write\_res\_pwd is omitted
and file\_text requires a password, the Password dialog box is
displayed.

Ignore\_rorec    is a logical value that controls whether the read-only
recommended message is displayed. If ignore\_rorec is TRUE, Microsoft
Excel prevents display of the message; if FALSE or omitted, and if
read\_only is also FALSE or omitted, Microsoft Excel displays the alert
when opening a read-only recommended workbook.

File\_origin    is a number specifying whether a text file originated on
the Macintosh or in Windows.

|                  |                                    |
| ---------------- | ---------------------------------- |
| **File\_origin** | **Original operating environment** |
| 1                | Macintosh                          |
| 2                | Windows (ANSI)                     |
| 3                | MS-DOS (PC-8)                      |

OmittedCurrent operating environment

Custom\_delimit    is the character you want to use as a custom
delimiter when opening text files.

  - > Custom\_delimit is text or a reference or formula that returns
    > text, such as CHAR(124).

  - > Custom\_delimit is required if format is 6; it is ignored if
    > format is not 6.

  - > Only the first character in custom\_delimit is used.

Add\_logical    is a logical value that specifies whether or not to add
file\_text to the open workbook. If add\_logical is TRUE, the document
is added; if FALSE or omitted, it is not added. This argument is for
compatibility with workbooks from Microsoft Excel version 4.0.

Editable    is a logical value that corresponds to opening a file (such
as a template) while holding down SHIFT key. If TRUE, editable is the
equivalent to holding down the SHIFT key while clicking the OK button in
the Open dialog box. If FALSE or omitted, this argument is ignored.

File\_access     is a number specifying how the file is to be accessed.
If the file is being opened for the first time, this argument is
ignored. If the file is already opened, this argument determines how to
change the user's access permissions for the file.

|                 |                             |
| --------------- | --------------------------- |
| **File Access** | **How Accessed**            |
| 1               | Revert to saved copy        |
| 2               | Change to read/write access |
| 3               | Change to read only access  |

Notify\_logical    is a logical value that specifies whether the user
should be notified when the shared workbook is available to be opened
across a network. If TRUE, the user will be notified when the workbook
is available to be opened. If FALSE or omitted, the user will not be
notified when the file available to be opened.

Converter     is a number corresponding to the file converter to use to
open the file. Normally, Microsoft Excel automatically determines which
file converter to use; therefore, this argument can usually be excluded.
If you want to be certain, however, that a specific manually installed
converter be used, then include this argument. Use GET.WORKSPACE(62) to
determine which numbers corresponds to all of the installed converters.

**Related Functions**

CLOSE   Closes the active window

FCLOSE   Closes a text file

FOPEN   Opens a file with the type of permission specified

OPEN.LINKS   Opens specified supporting workbooks

Return to [top](#H)

# OPEN.DIALOG

Displays the standard Microsoft Excel File Open dialog box with the
specified file filters. When the user presses the button specified by
button\_text, the filename the user picks or types in will be returned.

**Syntax**

**OPEN.DIALOG**(file\_filter, button\_text, title, filter\_index)

File\_filter    is the file filtering criteria to use, as text. For
Microsoft Excel for Windows, file\_filter consists of two parts, a
descriptive phrase denoting the file type followed by a comma and then
the MS-DOS wildcard file filter specification, as in "Text Files
(\*.TXT), \*.TXT, Add-in Files (\*.XLA), \*.XLA". Groups of filter
specifications are also separated by commas. Each separate pair is
listed in the file type drop-down list box. File\_filter can include an
asterisk (\*) to represent any sequence of characters and a question
mark (?) to represent any single character. For Microsoft Excel for the
Macintosh, file\_filter consists of file type codes separated by commas,
as in "TEXT,XLA ,XLS4". Spaces are significant and should not be
inserted before or after the comma separators unless they are part of
the file type code.

Button\_text    is the text used for the Open button in the dialog. If
omitted, "Open" will be used. Button\_text is ignored on Microsoft Excel
for Windows.

Title    specifies the dialog window title. If omitted, "File Open" will
be used as the default.

Filter\_index    specifies the index number of the default file
filtering criteria from 1 to the number of filters specified in
file\_filter. If omitted or greater than the number of filters present,
1 will be used as the starting index number. The argument is ignored on
Microsoft Excel for the Macintosh.

**Remarks**

  - > To use multiple MS-DOS wildcard expressions within file\_filter
    > for a single file filter type, separate the wildcard expressions
    > with semicolons, as in "VB Files (\*.bas; \*.txt), \*.bas;\*.txt"

  - > If file\_filter is omitted, "ALL Files (\*.\*), \*.\*" will be
    > used as the default in Microsoft Excel for Windows. The default
    > for Microsoft Excel for the Macintosh is all file types.

  - > If the user cancels the dialog box, FALSE is returned.

**Examples**

OPEN.DIALOG("Text Files (\*.TXT), \*.TXT, Add-in Files (\*.XLA),
\*.XLA,ALL FILES (\*.\*), \*.\*","FILE OPENER") opens a file open dialog
box titled "FILE OPENER" with three file filter criteria in the
drop-down list box.

**Related Function**

SAVE.DIALOG   Displays the standard Microsoft Excel File Save As dialog
box and gets a file name from the user

Return to [top](#H)

# OPEN.LINKS

Equivalent to clicking the Links command on the Edit menu. Use
OPEN.LINKS with the LINKS function to open workbooks linked to a
particular sheet.

**Syntax**

**OPEN.LINKS**(**document\_text1**, document\_text2, ..., read\_only,
type\_of\_link)

**OPEN.LINKS**?(document\_text1, document\_text2, ..., read\_only,
type\_of\_link)

Document\_text1, document\_text2,    are 1 to 12 arguments that are the
names of supporting documents in the form of text, or arrays or
references that contain text.

Read\_only    is a logical value corresponding to the read/write status
of the linked worksheet. If read\_only is TRUE, the sheet can be
modified but changes cannot be saved; if FALSE or omitted, changes to
the sheet can be saved. Read\_only applies only to Microsoft Excel, WKS,
and SYLK documents.

Type\_of\_link    is a number from 1 to 6 that specifies what type of
link you want to get information about.

|                    |                        |
| ------------------ | ---------------------- |
| **Type\_of\_link** | **Link document type** |
| 1                  | Microsoft Excel link   |
| 2                  | DDE link               |
| 3                  | Reserved               |
| 4                  | Not applicable         |
| 5                  | Subscriber             |
| 6                  | Publisher              |

**Remarks**

You can generate an array of the names of linked workbooks with the
LINKS function.

**Related Functions**

CHANGE.LINK   Changes supporting workbook links

GET.LINK.INFO   Returns information about a link

LINKS   Returns the name of all linked workbooks

UPDATE.LINK   Updates a link to another document

Return to [top](#H)

# OPEN.MAIL

Equivalent to clicking the Open Mail command on the Mail submenu on File
menu.

**Note   **This function is available for only Microsoft Excel for the
Macintosh and Microsoft Mail version 2.0 or later.

**Syntax**

**OPEN.MAIL**(subject, comments)

**OPEN.MAIL**?(subject, comments)

Subject    is the subject of the message containing a file that
Microsoft Excel can open.

  - > For each message whose subject matches the subject argument and
    > contains a file that Microsoft Excel can open, the file is opened
    > in Microsoft Excel; if the message has no unread enclosures, it is
    > deleted from the list of pending mail.

  - > If subject is omitted, then for all messages containing files that
    > Microsoft Excel can open, the files are opened; each message that
    > has no unread enclosures is deleted from the list of pending mail.

>  

Comments    is a logical value that specifies whether comments
associated with the Microsoft Excel files are displayed. If comments is
TRUE, Microsoft Excel displays the comments; if FALSE, comments are not
displayed. If omitted, the current setting is not changed.

**Tips**

  - > If you use consistent subjects in your Microsoft Mail messages,
    > you can easily create a macro that always opens mail messages with
    > certain files attached. For example, an OPEN.MAIL formula with
    > subject specified as "Weekly Report" would open the Microsoft
    > Excel file attached to the message containing that subject each
    > week.

  - > In OPEN.MAIL?, the dialog-box form of the function, the currently
    > running macro pauses while the Microsoft Mail documents window is
    > displayed. The macro resumes after you close the Microsoft Mail
    > documents window.

**Related Function**

SEND.MAIL   Sends the active workbook

Return to [top](#H)

# OPEN.TEXT

Equivalent to using the Text Import Wizard to open a text file in
Microsoft Excel.

**Syntax**

**OPEN.TEXT**(file\_name, file\_origin, start\_row, file\_type,
text\_qualifier, consecutive\_delim, tab, semicolon, comma, space,
other, other\_char, {field\_info1; field\_info2;...})

File\_name    is the full pathname of the text file you want to open.

File\_origin    specifies the operating environment the text file was
created in.

|                  |                               |
| ---------------- | ----------------------------- |
| **File\_origin** | **Operating system**          |
| 1                | Macintosh                     |
| 2                | Windows (ANSI)                |
| 3                | MS DOS (PC-8)                 |
| Omitted          | Current operating environment |

Start\_row    is a number greater than or equal to one, specifying the
row in the text file where you want to start importing into Microsoft
Excel. This number should be less than the number of lines in the text
file.

File\_type    specifies the type of delimited text file to import:

|                |                  |
| -------------- | ---------------- |
| **File\_type** | **Type of file** |
| 1 or omitted   | Delimited        |
| 2              | Fixed width      |

Text\_qualifier    indicates the character-enclosing text fields in the
text file:

|                           |                           |
| ------------------------- | ------------------------- |
| **Text\_qualifier value** | **Qualifier**             |
| 1 or "                    | " (double quotation mark) |
| 2 or '                    | ' (single quotation mark) |
| 3 or {None}               | No text qualifier         |

Consecutive\_delim    is a logical value corresponding to the Treat
Consecutive Delimiters as One check box which, if TRUE, allows
consecutive delimiters (such as ",,,") to be treated as a single
delimiter. If FALSE, all consecutive delimiters are considered separate
column breaks.

Tab, semicolon, comma, space    are logical values corresponding to the
check boxes in the Column Delimiters group. If an argument is TRUE, the
check box is selected. These arguments apply only when the file\_type
argument is 1 or omitted (delimited file type).

Other    is a logical value indicating a custom delimiter to use in
opening the text file.

Other\_char    specifies the custom delimiter to use or FALSE if no
custom delimiter is used.

Field\_info    is an array which consists of the following elements:
"column number, data\_format", if file\_type is 1; or "start\_pos,
data\_format" if file\_type is 2.

**Related Function**

TEXT.TO.COLUMNS   Parses text, as in a text file, into columns of data

Return to [top](#H)

# OPTIONS.CALCULATION

Equivalent to clicking the Options command on the Tools menu, and
clicking the Calculation tab in the Options dialog box. Sets various
worksheet calculation settings.

**Syntax**

**OPTIONS.CALCULATION**(type\_num, iter, max\_num, max\_change, update,
precision, date\_1904, calc\_save, save\_values)

**OPTIONS.CALCULATION**?(type\_num, iter, max\_num, max\_change, update,
precision, date\_1904, calc\_save, save\_values)

Arguments correspond to check boxes and options in the Calculation tab
in the Options dialog box. Arguments that correspond to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Type\_num    is a number from 1 to 3 indicating the type of calculation.

|               |                         |
| ------------- | ----------------------- |
| **Type\_num** | **Type of calculation** |
| 1             | Automatic               |
| 2             | Automatic except tables |
| 3             | Manual                  |

Iter    corresponds to the Iteration check box. The default is FALSE.

Max\_num    is the maximum number of iterations. The default is 100.

Max\_change    is the maximum change of each iteration. The default is
0.001.

Update    corresponds to the Update Remote References check box. The
default is TRUE.

Precision    corresponds to the Precision As Displayed check box. The
default is FALSE.

Date\_1904    corresponds to the 1904 Date System check box. The default
is FALSE in Microsoft Excel for Windows and TRUE in Microsoft Excel for
the Macintosh.

Calc\_save    corresponds to the Recalculate Before Save check box. If
calc\_save is FALSE, the workbook is not recalculated before saving when
in manual calculation mode. The default is TRUE.

Save\_values    corresponds to the Save External Link Values check box.
The default is TRUE.

**Note   **Microsoft Excel for Windows and Microsoft Excel for the
Macintosh use different date systems as their default. For more
information, see NOW.

Return to [top](#H)

# OPTIONS.CHART

Equivalent to clicking the Options command on the Tools menu and then
clicking the Chart Tab in the Options dialog box when a chart is
activated for editing. Sets various chart settings.

**Syntax**

**OPTIONS.CHART**(Display\_Blanks, Plot\_Visible, Size\_With\_Window)

**OPTIONS.CHART**?( Display\_Blanks, Plot\_Visible, Size\_With\_Window)

Display\_Blanks    is a number indicating how blank cells are plotted.

|            |                              |
| ---------- | ---------------------------- |
| **Number** | **Blanks are displayed as**  |
| 1          | Not plotted (gaps are shown) |
| 2          | Zero values                  |
| 3          | Interpolated                 |

Plot\_Visible    is a logical value that if TRUE plots only visible
data. If FALSE, all cells in the selection are plotted.

Size\_With\_Window    is a logical value that if TRUE allows the chart
to resize with window. If FALSE, chart will not size with window.

**Remarks**

If any of the arguments are omitted, then that setting is unchanged
within the Options dialog box.

**Related Functions**

PREFERRED   Changes the format of the active chart

SET.PREFERRED   Changes the default chart format

Return to [top](#H)

# OPTIONS.EDIT

Equivalent to clicking the Options command on the tools menu and then
clicking the Edit tab in the Options dialog box. Sets various worksheet
editing options.

**Syntax**

**OPTIONS.EDIT**(incell\_edit, drag\_drop, alert, entermove, fixed,
decimals, copy\_objects, update\_links, move\_direction, autocomplete,
animations)

**OPTIONS.EDIT**?(incell\_edit, drag\_drop, alert, entermove, fixed,
decimals, copy\_objects, update\_links, move\_direction, autocomplete,
animations)

Incell\_edit    is a logical value corresponding to the Edit Directly In
Cell check box, which if TRUE allows In Cell Editing. If FALSE, editing
directly in cells is not allowed. If omitted, the dialog box setting is
not changed.

Drag\_drop    is a logical value corresponding to the Allow Cell Drag
And Drop check box, which if TRUE allows drag and dropping on sheets. If
FALSE, drag and drop is not allowed. If omitted, the dialog box setting
is not changed.

Alert    is a logical value corresponding to the Alert Before
Overwriting Cells check box, which if TRUE displays an alert message
warning you that cells containing values are about to be overwritten. If
FALSE, an alert will not be displayed if your cells are about to
overwritten. If omitted, the dialog box setting is not changed.

Entermove    is a logical value corresponding to the Move Selection
After Enter check box, which if TRUE moves the selection after the ENTER
key is pressed. If FALSE, the selection is not moved. If omitted, the
dialog box setting is not changed.

Fixed    is a logical value corresponding to the Fixed Decimal check
box, which if TRUE fixes the decimal place according to decimals. If
FALSE, the decimal places are not fixed. If omitted, the dialog box
setting is not changed.

Decimals    is a number specifying the number of decimal places.
Decimals is ignored if fixed is FALSE or omitted.

Copy\_objects     is a logical value corresponding to the Cut, Copy, And
Sort Objects With Cells check box. If TRUE allows objects to be cut,
copied and sorted with their cells. If FALSE, objects are not cut,
copied or sorted with cells. If omitted, the dialog box setting is not
changed.

Update\_links    is a logical value corresponding to the Ask To Update
Automatic Links check box, which if TRUE will prompt the user when the
workbook is opened that has links to other documents. If FALSE, the
prompt will not be displayed. If omitted, the dialog box setting is not
changed.

Move\_direction    is a number specifying the direction to move the
selection when the ENTER key is pressed and Entermove is TRUE. Setting
this number to 1 moves down one cell, 2 moves right one cell, 3 moves up
on cell, and 4 moves down one cell..

Autocomplete    is a logical value corresponding to the Enable
AutoComplete For Cell Values check box, which, if TRUE, will enable the
AutoComplete feature of Microsoft Excel 95 and later versions.

Animations    is a logical value corresponding to the Provide Feedback
With Animation check box, which if TRUE will enable the worksheet
Animations feature of Microsoft Excel 95 and later versions. Deleted
worksheet rows and columns will slowly disappear, and inserted worksheet
rows and columns will slowly appear.

**Related Function**

OPTIONS.GENERAL   Sets various general Microsoft Excel settings

Return to [top](#H)

# OPTIONS.GENERAL

Equivalent to clicking the Options command on the Tools menu and then
clicking the General tab from the Options dialog box. Sets various
general Microsoft Excel settings.

**Syntax**

**OPTIONS.GENERAL**(R1C1\_mode, dde\_on, sum\_info, tips, recent\_files,
old\_menus, user\_info, font\_name, font\_size, default\_location,
alternate\_location, sheet\_num, enable\_under)

**OPTIONS.GENERAL**?(R1C1\_mode, dde\_on, sum\_info, tips,
recent\_files, old\_menus, user\_info, font\_name, font\_size,
default\_location, alternate\_location, sheet\_num, enable\_under)

Arguments correspond to option buttons, check boxes and text boxes on
the General tab of the Options dialog box. Arguments corresponding to
check boxes are logical values. If an argument is TRUE, the check box is
selected; if FALSE, the check box is cleared; if omitted, the current
setting is not changed.

R1C1\_mode     is a number specifying the reference style. Use 1 for A1
style references; 2 for R1C1 style references.

Dde\_on    is a logical value corresponding to the Ignore Other
Applications check box, which if TRUE ignores DDE request from other
applications. If FALSE, DDE requests from other applications are allowed
to happen.

Sum\_info    is a logical value corresponding to the Prompt For Workbook
Properties check box, which if TRUE displays the Summary tab of the
Properties dialog box when a workbook is initially saved. If FALSE, the
dialog box is not displayed.

Tips    is a logical value corresponding to the Reset TipWizard check
box in Microsoft Excel 95 or earlier versions, which, if TRUE, resets
the TipWizard. In Microsoft Excel 97 or later, tips corresponds to the
Reset My Tips button in the Office Assistant dialog box. If FALSE, the
TipWizard will not be reset.

Recent\_files    is a logical value corresponding to the Recently Used
File List check box, which if TRUE displays the four last opened files
from the File menu. If FALSE, the file list will not be displayed.

Old\_menus    is a logical value corresponding to the Microsoft Excel
4.0 Menus check box in Microsoft Excel version 5.0, which if TRUE
replaces the current menu bar with the Microsoft Excel 4.0 menu bar. If
FALSE, the menu bar will not be replaced. This argument is for
compatibility with Microsoft Excel version 5.0 only and is ignored in
later versions. User\_info    corresponds to the Name text box, and is
the name of the user of this copy of Microsoft Excel. By default it is
the registered user, but can be changed to work on a network.

Font\_name     corresponds to the Standard Font text box, and is the
name of the default font.

Font\_size    corresponds to the Size drop-down edit box, and is the
size of the default font.

Default\_location     corresponds to the Default File Location text box,
and is the default location that the File Open command displays. The
Default is where Microsoft Excel is installed.

Alternate\_location    corresponds to the Alternate Startup File
Location text box, and is the alternate startup directory.

Sheet\_num    corresponds to the Sheets In New Workbook spin control,
and is the number of sheets in a new workbook. Default is 3. Can go up
to 255.

Enable\_under    enables underlining of the menus. Used for only
Microsoft Excel for the Macintosh. Ignored in Microsoft Excel for
Windows.

**Related Functions**

OPTIONS.LISTS.DELETE   Deletes a custom list

OPTIONS.LISTS.GET   Returns contents of custom AutoFill lists

OPTIONS.VIEW   Sets various view settings

Return to [top](#H)

# OPTIONS.LISTS.ADD

This is the equivalent to clicking the Options command on the Tools menu
and then clicking the Custom Lists tab in the Options diralog box. Used
to add a new custom list.

**Syntax**

**OPTIONS.LISTS.ADD**(**string\_array**)

**OPTIONS.LISTS.ADD**(**import\_ref, by\_row**)

**OPTIONS.LISTS.ADD**?(**import\_ref, list\_num**)

String\_array    is an array of strings or cell reference that contains
the custom items in the list, a named cell reference, or an external
reference containing the items of the custom list to add.

Import\_ref    is the reference to the cells that contain the members of
the custom list. If A1:A12 contains the twelve signs of the Zodiac
starting with Aquarius, then this function will add the contents of
these twelve cells as a custom list.

By\_row    is a logical value that if TRUE, and if importing from cells,
assumes that the list items are in sequential rows. If FALSE, assumes
that the list items are in columns. If omitted, Microsoft Excel will try
to determine the order of the custom lists according to the layout of
the sheet.

List\_num    is a number specifying which list to activate. If omitted,
then New List will be activated.

**Remarks**

  - > To replace an existing custom list, you must first delete it and
    > then add the new list to the end.

  - > If the list already exists, then this function will do nothing.
    > The list is not case sensitive, so "Scorpio" and "scorpio" are
    > treated the same in custom lists.

**Related Functions**

OPTIONS.VIEW   Sets various view settings

OPTIONS.LISTS.GET   Returns contents of custom AutoFill lists

OPTIONS.LISTS.DELETE   Deletes a custom list

Return to [top](#H)

# OPTIONS.LISTS.DELETE

Equivalent to clicking the Options command on the Tools menu and then
clicking the Delete button on the Custom Lists tab when a custom list is
selected.

**Syntax**

**OPTIONS.LISTS.DELETE**(**list\_num**)

List\_num    is the number of the custom list to delete. The first five
lists (numbered zero through 4) cannot be deleted. If list\_num doesn't
exist, then FALSE is returned.

**Related Functions**

OPTIONS.LISTS.GET   Returns contents of custom AutoFill lists

OPTIONS.LISTS.ADD   Used to add a new custom list

Return to [top](#H)

# OPTIONS.LISTS.GET

Returns contents of custom AutoFill lists as an array of text strings.

**Syntax**

**OPTIONS.LISTS.GET**(**list\_num**)

**OPTIONS.LISTS.GET**(**string\_array**)

List\_num   is a number specifying which list to return, as a horizontal
string array. If list\_num is zero, then FALSE is returned.

String\_array    is an array or cell reference to store the returned
list.

**Example**

OPTIONS.LIST.GET(3) returns the twelve months of the year in the form
{"Jan", "Feb", "Mar"}

**Remarks**

If list\_num is zero or omitted, then FALSE is returned.

**Related Functions**

OPTIONS.LISTS.ADD   Adds a new custom list

OPTIONS.LISTS.DELETE   Deletes a custom list

OPTIONS.VIEW   Sets various view settings

Return to [top](#H)

# OPTIONS.TRANSITION

Equivalent to clicking the Options command on the Tools menu and then
clicking the Transition tab in the Options dialog box. Sets options
relating to compatibility with other spreadsheets.

**Syntax**

**OPTIONS.TRANSITION**(menu\_key, menu\_key\_action, nav\_keys,
trans\_eval, trans\_entry)

**OPTIONS.TRANSITION**?(menu\_key, menu\_key\_action, nav\_keys,
trans\_eval, trans\_entry)

Menu\_key    is text specifying which alternate menu key to use.

Menu\_key\_action    is the number 1 or 2 specifying options for the
alternate menu or Help key. In Microsoft Excel for the Macintosh,
menu\_key\_action is ignored.

|                       |                                          |
| --------------------- | ---------------------------------------- |
| **Menu\_key\_action** | **Alternate menu or Help key activates** |
| 1 or omitted          | Microsoft Excel menus                    |
| 2                     | Lotus 1-2-3 Help                         |

Nav\_keys    is a logical value that corresponds to the Transition
Navigation Keys check box, which if TRUE uses alternate navigation keys
that correspond to the navigation keys for Lotus 1-2-3. In Microsoft
Excel for the Macintosh, nav\_keys is ignored.

Trans\_eval    is a logical value that corresponds to the Transition
Formula Evaluation check box.

  - > If trans\_eval is TRUE, Microsoft Excel uses a set of rules
    > compatible with that of Lotus 1-2-3 when calculating formulas.
    > Text is treated as 0. TRUE and FALSE are treated as 1 and 0.
    > Certain characters in database criteria ranges are interpreted the
    > same way Lotus 1-2-3 interprets them.

  - > If trans\_eval is FALSE or omitted, Microsoft Excel calculates
    > normally.

Trans\_entry    is a logical value that corresponds to the Transition
Formula Entry check box.

  - > This argument is available only in Microsoft Excel for Windows.

  - > If trans\_entry is TRUE, Microsoft Excel accepts formulas entered
    > in  
    > Lotus 1-2-3 style.

  - > If trans\_entry is FALSE or omitted, Microsoft Excel only accepts
    > formulas entered in Microsoft Excel style.

**Related Functions**

OPTIONS.LISTS.DELETE   Deletes a custom list

OPTIONS.LISTS.GET   Returns contents of custom AutoFill lists

OPTIONS.VIEW   Sets various view settings

Return to [top](#H)

# OPTIONS.VIEW

Equivalent to clicking the Options command on the Tools menu and then
clicking the View tab in the Options dialog box. Sets various view
settings.

**Syntax**

**OPTIONS.VIEW**(formula, status, notes, show\_info, object\_num,
page\_breaks, formulas, gridlines, color\_num, headers, outline, zeros,
hor\_scroll, vert\_scroll, sheet\_tabs)

**OPTIONS.VIEW**?(formula, status, notes, show\_info, object\_num,
page\_breaks, formulas, gridlines, color\_num, headers, outline, zeros,
hor\_scroll, vert\_scroll, sheet\_tabs)

Arguments correspond to check boxes and text boxes in the View tab on
the Options dialog box. Arguments corresponding to check boxes are
logical values. If an argument is TRUE, the check box is selected; if
FALSE, the check box is cleared; if omitted, the current setting is not
changed.

Formula    is a logical value corresponding to the Formula Bar check
box. If TRUE, displays the formula bar. If FALSE, the formula bar is not
displayed.

Status    is a logical value corresponding to the Status Bar check box.
If TRUE, the status bar is displayed. If FALSE, the status bar is not
displayed.

Notes    is a logical value corresponding to the Comment & Indicator
check box. If TRUE, comments and comment indicators will be displayed.
If FALSE, comments and indicators will not be displayed.

Show\_info    is a logical value corresponding to the Info Window check
box (only in Microsoft Excel 95 and earlier versions). If TRUE, displays
the Info Window. The Info Window is not available in Microsoft Excel 97
or later.

Object\_num    is a number from 1 to 3 corresponding to the display
options in the Objects box.

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

Page\_breaks    is a logical value corresponding to the Page Breaks
check box. If TRUE, page breaks will appear. If FALSE, page breaks will
not appear.

Formulas    is a logical value corresponding to the Formulas check box.
If TRUE, formulas will appear in the cells. If FALSE, formulas will not
appear in the cells. The default is FALSE on worksheets and TRUE on
macro sheets.

Gridlines    is a logical value corresponding to the Gridlines check
box. If TRUE, gridlines will be displayed. If FALSE, gridlines will not
appear. The default is TRUE.

Color\_num    is a number from 0 to 56 corresponding to gridline color.
Zero corresponds to automatic color and is the default value.

Headings    is a logical value corresponding to the Row & Column Headers
check box. If TRUE, row and column headers will be displayed. If FALSE,
they will not be displayed. The default is TRUE.

Outline    is a logical value corresponding to the Outline Symbols check
box. If TRUE, outline symbols will appear. If FALSE, they will not
appear. The default is TRUE.

Zeros    is a logical value corresponding to the Zero Values check box.
If TRUE, zero values will appear, If FALSE, zero values will not appear.
The default is TRUE.

Hor\_scroll    is a logical value corresponding to the Horizontal Scroll
Bar checkbox. If TRUE, the horizontal scroll bar will be displayed. If
FALSE, it will not be displayed. The default is TRUE.

Vert\_scroll    is a logical value corresponding to the Vertical Scroll
Bar checkbox. If TRUE, the vertical scroll bar will be displayed. If
FALSE, it will not be displayed. The default is TRUE.

Sheet\_tabs    is a logical value corresponding to the Sheet Tabs check
box. If TRUE, sheet tabs will be displayed. If FALSE, sheet tabs will
not be displayed. The default is TRUE.

**Related Functions**

OPTIONS.LISTS.GET   Returns contents of custom AutoFill lists

OPTIONS.LISTS.DELETE   Deletes a custom list

Return to [top](#H)

# OUTLINE

Creates an outline and defines settings for automatically creating
outlines.

The first three arguments are logical values corresponding to check
boxes in the Outline dialog box, which appears when you choose the
Settings command from the Group and Outline submenu on the Data menu. If
an argument is TRUE, Microsoft Excel selects the check box; if FALSE,
Microsoft Excel clears the check box. If an argument is omitted, the
check box is left in its current state.

**Syntax**

**OUTLINE**(auto\_styles, row\_dir, col\_dir, create\_apply)

**OUTLINE**?(auto\_styles, row\_dir, col\_dir, create\_apply)

Auto\_styles    corresponds to the Automatic Styles check box.

Row\_dir    corresponds to the Summary Rows Below Detail check box.

Col\_dir    corresponds to the Summary Columns To Right Of Detail check
box.

Create\_apply    is the number 1 or 2 and corresponds to the Create
button and the Apply Styles button.

|                   |                                                                         |
| ----------------- | ----------------------------------------------------------------------- |
| **Create\_apply** | **Result**                                                              |
| 1                 | Creates an outline with the current settings                            |
| 2                 | Applies outlining styles to the selection based on outline levels       |
| Omitted           | Corresponds to choosing the OK button to set the other outline settings |

**Related Functions**

DEMOTE   Demotes the selection in an outline

PROMOTE   Promotes the selection in an outline

Return to [top](#H)

# OVERLAY

Equivalent to choosing the Overlay command from the Format menu in
Microsoft Excel version 2.2 or earlier. This function is included only
for macro compatibility. To format chart types in Microsoft Excel
version 5.0 or later, use the FORMAT.CHART function.

**Syntax**

**OVERLAY**(**type\_num**, stack, 100, vary, overlap, drop, hilo,
overlap%, cluster, angle, series\_num, auto)

**Related Function**

FORMAT.CHART   Formats the chart according to the arguments you specify.

Return to [top](#H)

# PAGE.SETUP

Equivalent to clicking the Page Setup command on the File menu. Use
PAGE.SETUP to control the printed appearance of your sheets.

There are three syntax forms of PAGE.SETUP. Syntax 1 applies if a sheet
or macro sheet is active; syntax 2 applies if a chart is active; syntax
three applies to Visual Basic modules and the info Window.

Arguments correspond to check boxes and text boxes in the Page Setup
dialog box. Arguments that correspond to check boxes are logical values.
If an argument is TRUE, Microsoft Excel selects the check box; if FALSE,
Microsoft Excel clears the check box. Arguments for margins are always
in inches, regardless of your country setting.

**Syntax 1**

Worksheets and macro sheets

**PAGE.SETUP**(head, foot, left, right, top, bot, hdng, grid, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, pg\_order, bw\_cells,
quality, head\_margin, foot\_margin, notes, draft)

**PAGE.SETUP**?(head, foot, left, right, top, bot, hdng, grid, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, pg\_order, bw\_cells,
quality, head\_margin, foot\_margin, notes, draft)

**Syntax 2**

Charts

**PAGE.SETUP**(head, foot, left, right, top, bot, size, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, bw\_chart, quality,
head\_margin, foot\_margin, draft)

**PAGE.SETUP**?(head, foot, left, right, top, bot, size, h\_cntr,
v\_cntr, orient, paper\_size, scale, pg\_num, bw\_chart, quality,
head\_margin, foot\_margin, draft)

**Syntax 3**

Visual Basic Modules and the Info Window

**PAGE.SETUP**(head, foot, left, right, top, bot, orient, paper\_size,
scale, quality, head\_margin, foot\_margin, pg\_num)

**PAGE.SETUP**?(head, foot, left, right, top, bot, orient, paper\_size,
scale, quality, head\_margin, foot\_margin, pg\_num)

Head    specifies the text and formatting codes for the header for the
current sheet . For information about formatting codes, see "Remarks"
later in this topic.

Foot    specifies the text and formatting codes for the workbook footer.

Left    corresponds to the Left box and is a number specifying the left
margin.

Right    corresponds to the Right box and is a number specifying the
right margin.

Top    corresponds to the Top box and is a number specifying the top
margin.

Bot    corresponds to the Bottom box and is a number specifying the
bottom margin.

Hdng    corresponds to the Row & Column Headings check box. Hdng is
available only in the sheet and macro sheet form of the function.

Grid    corresponds to the Cell Gridlines check box. Grid is available
only in the sheet and macro sheet form of the function.

H\_cntr    corresponds to the Center Horizontally check box in the
Margins panel of the Page Setup dialog box.

V\_cntr    corresponds to the Center Vertically check box in the Margins
panel of the Page Setup dialog box.

Orient    determines the direction in which your workbook is printed.

|            |                  |
| ---------- | ---------------- |
| **Orient** | **Print format** |
| 1          | Portrait         |
| 2          | Landscape        |

Paper\_size    is a number from 1 to 26 that specifies the size of the
paper.

|                 |                |
| --------------- | -------------- |
| **Paper\_size** | **Paper type** |
| 1               | Letter         |
| 2               | Letter (small) |
| 3               | Tabloid        |
| 4               | Ledger         |
| 5               | Legal          |
| 6               | Statement      |
| 7               | Executive      |
| 8               | A3             |
| 9               | A4             |
| 10              | A4 (small)     |
| 11              | A5             |
| 12              | B4             |
| 13              | B5             |
| 14              | Folio          |
| 15              | Quarto         |
| 16              | 10x14          |
| 17              | 11x17          |
| 18              | Note           |
| 19              | ENV9           |
| 20              | ENV10          |
| 21              | ENV11          |
| 22              | ENV12          |
| 23              | ENV14          |
| 24              | C Sheet        |
| 25              | D Sheet        |
| 26              | E Sheet        |

Scale    is a number representing the percentage to increase or decrease
the size of the sheet. All scaling retains the aspect ratio of the
original.

  - > To specify a percentage of reduction or enlargement, set scale to
    > the percentage.

  - > For worksheets and macros, you can specify the number of pages
    > that the printout should be scaled to fit. Set scale to a two-item
    > horizontal array, with the first item equal to the width and the
    > second item equal to the height. If no constraint is necessary in
    > one direction, you can set the corresponding value to \#N/A.

  - > Scale can also be a logical value. To fit the print area on a
    > single page, set scale to TRUE.

>  

Pg\_num    specifies the number of the first page. If zero, sets first
page to zero. If "Auto" is used, then the page numbering is set to
automatic. If omitted, PAGE.SETUP retains the existing pg\_num.

Pg\_order    specifies whether pagination is left-to-right and then
down, or top-to-bottom and then right.

|               |                           |
| ------------- | ------------------------- |
| **Pg\_order** | **Pagination**            |
| 1             | Top-to-bottom, then right |
| 2             | Left-to-right, then down  |

Bw\_cells    is a logical value that specifies whether to print cells
and all graphic objects, such as text boxes and buttons, in color.

  - > If bw\_cells is TRUE, Microsoft Excel prints cell text and borders
    > in black and cell backgrounds in white.

  - > If bw\_cells is FALSE , Microsoft Excel prints cell text, borders,
    > and background patterns in color (or in gray scale).

>  

Bw\_chart    is a logical value that specifies whether to print chart in
color.

Size    is a number corresponding to the options in the Chart Size box,
and determines how you want the chart printed on the page within the
margins. Size is available only in the chart form of the function.

|          |                             |
| -------- | --------------------------- |
| **Size** | **Size to print the chart** |
| 1        | Screen size                 |
| 2        | Fit to page                 |
| 3        | Full page                   |

Quality    specifies the print quality in dots-per-inch. To specify both
horizontal and vertical print quality, use an array of two values.

Head\_margin    is the placement, in inches, of the running head margin
from the edge of the page.

Foot\_margin    is the placement, in inches, of the running foot margin
from the edge of the page.

Draft    corresponds to the Draft Quality checkbox in the Sheet tab and
in the Chart tab of the Page Setup dialog box. If FALSE or omitted,
graphics are printed with the sheet. If TRUE, no graphics are printed.

Notes    specifies whether to print cell notes with the sheet. If TRUE,
both the sheet and the cell notes are printed. If FALSE or omitted, just
the sheet is printed.

**Remarks**

Microsoft Excel no longer requires you to enter formatting codes to
format headers and footers, but the codes are still supported and
recorded by the macro recorder. You can include these codes as part of
the head and foot text strings to align portions of the header or footer
to the left, right, or center; to include the page number, date, time,
or workbook name; and to print the header or footer in bold or italic.

|                         |                                                                                                                                                                                             |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Formatting code**     | **Result**                                                                                                                                                                                  |
| \&L                     | Left-aligns the characters that follow.                                                                                                                                                     |
| \&C                     | Centers the characters that follow.                                                                                                                                                         |
| \&R                     | Right-aligns the characters that follow.                                                                                                                                                    |
| \&B                     | Turns bold printing on or off (now obsolete).                                                                                                                                               |
| \&I                     | Turns italic printing on or off.                                                                                                                                                            |
| \&U                     | Turns single underlining printing on or off.                                                                                                                                                |
| \&S                     | Turns strikethrough printing on or off.                                                                                                                                                     |
| \&O                     | Turns outline printing on or off (Macintosh only).                                                                                                                                          |
| \&H                     | Turns shadow printing on or off (Macintosh only).                                                                                                                                           |
| \&D                     | Prints the current date.                                                                                                                                                                    |
| \&T                     | Prints the current time.                                                                                                                                                                    |
| \&A                     | Prints the name of the sheet                                                                                                                                                                |
| \&F                     | Prints the name of the workbook.                                                                                                                                                            |
| \&P                     | Prints the page number.                                                                                                                                                                     |
| \&P+number              | Prints the page number plus number.                                                                                                                                                         |
| \&P-number              | Prints the page number minus number.                                                                                                                                                        |
| &&                      | Prints a single ampersand.                                                                                                                                                                  |
| & "fontname, fontstyle" | Prints the characters that follow in the specified font and style. Be sure to include a comma immediately following the fontname, and double quotation marks around fontname and fontstyle. |
| \&nn                    | Prints the characters that follow in the specified font size. Use a two-digit number to specify a size in points.                                                                           |
| \&N                     | Prints the total number of pages in the workbook.                                                                                                                                           |
| \&E                     | Prints a double underline                                                                                                                                                                   |
| \&X                     | Prints the character as superscript                                                                                                                                                         |
| \&Y                     | Prints the character as subscript                                                                                                                                                           |

**Related Functions**

DISPLAY   Controls screen and Info Window display

GET.DOCUMENT   Returns information about a workbook

PRINT   Prints the active workbook

WORKSPACE   Changes workspace settings

Return to [top](#H)

# PARSE

Distributes the contents of the current selection to fill several
adjacent columns; the selection can be no more than one column wide. Use
PARSE to reorganize data, especially data that you've read from files
created by another application, such as a database.

**Syntax**

**PARSE**(parse\_text, destination\_ref)

**PARSE**?(parse\_text, destination\_ref)

Parse\_text    is the parse line in the form of text. It is a copy of
the first nonblank cell in the selected column, with square brackets
indicating where to distribute (or parse) text. If parse\_text is
omitted, Microsoft Excel guesses where to place the brackets based on
the spacing and formatting of data.

Destination\_ref    is a reference to the upper-left corner of the range
of cells where you want to place the parsed data. If destination\_ref is
omitted, it is assumed to be the current selection, so the parsed data
will replace the original data.

**Remarks**

When you use the PARSE function, Microsoft Excel splits the first column
into as many columns as you specify with parse\_text and replaces any
information in those columns.

Return to [top](#H)

# PASTE

Equivalent to clicking the Paste command on the Edit menu. Pastes a
selection or object that you copied or cut using the COPY or CUT
function. Use PASTE when you want to paste all components of the
selection. To paste only specific components of the selection, use
PASTE.SPECIAL.

**Syntax**

**PASTE**(to\_reference)

To\_reference    is a reference to the cell or range of cells where you
want to paste what you have copied. If to\_reference is omitted,
Microsoft Excel pastes to the current selection. If there is nothing to
paste, the macro halts.

**Related Functions**

COPY   Copies and pastes data or objects

CUT   Cuts or moves data or objects

FORMULA   Enters values into a cell or range or onto a chart

INSERT   Inserts cells

PASTE.LINK   Pastes copied data and establishes a link to its source

PASTE.SPECIAL   Pastes specific components of copied data

Return to [top](#H)

# PASTE.LINK

Equivalent to clicking the Paste Special command on the Edit menu, and
then clicking the Paste Link button in the Paste Special dialog box.
Pastes copied data or objects and establishes a link to the source of
the data or object. The source can be either another Microsoft Excel
workbook or another application. Use PASTE.LINK when you want Microsoft
Excel to automatically update the paste area with any changes that occur
in the source.

**Syntax**

**PASTE.LINK**( )

**Note   **To work properly, the application you are linking to must
support dynamic data exchange (DDE) or object linking and embedding
(OLE).

**Related Functions**

COPY   Copies and pastes data or objects

CUT   Cuts or moves data or objects

PASTE   Pastes cut or copied data

PASTE.SPECIAL   Pastes specific components of copied data

Return to [top](#H)

# PASTE.PICTURE

Equivalent to clicking the Paste Picture command on the Edit menu while
holding down the SHIFT key in Microsoft Excel version 4.0. Pastes a
picture of the Clipboard contents onto the sheet. This picture is not
linked, so changes to the source data will not be reflected in the
picture. In Microsoft Excel for Windows version 5.0 or later, use
INSERT.PICTURE to import pictures.

**Syntax**

**PASTE.PICTURE**( )

**Related Functions**

COPY.PICTURE   Creates a picture of the current selection for use in
another program

INSERT.PICTURE   Inserts a picture from a file

PASTE.PICTURE.LINK   Pastes a linked picture of the currently copied
area

Return to [top](#H)

# PASTE.PICTURE.LINK

Equivalent to holding down the SHIFT key and clicking the Paste Picture
Link command on the Edit menu in Microsoft Excel version 4.0 or to using
the camera tool. Pastes a linked picture of the Clipboard contents. This
picture is linked, so changes to the source data will be reflected in
the picture.

**Syntax**

**PASTE.PICTURE.LINK**( )

**Related Functions**

COPY.PICTURE   Copies and pastes data or object

COPY.PICTURE   Creates a picture of the current selection for use in
another program

CREATE.OBJECT   Creates an object

PASTE   Pastes cut or copied data

PASTE.PICTURE   Pastes a picture of the currently copied area

Return to [top](#H)

# PASTE.SPECIAL

Equivalent to clicking the Paste Special command on the Edit menu.
Pastes the specified components from the copy area into the current
selection. The PASTE.SPECIAL function has four syntax forms.

Syntax 1   Pasting into a sheet or macro sheet

Syntax 2   Copying from a sheet and pasting into a chart.

Syntax 3   Copying and pasting between charts

Syntax 4   Pasting information from another application.

Return to [top](#H)

# PASTE.SPECIAL Syntax 1

Equivalent to clicking the Paste Special command on the Edit menu.
Pastes the specified components from the copy area into the current
selection. The PASTE.SPECIAL function has four syntax forms. Use syntax
1 if you are pasting into a sheet or macro sheet.

**Syntax**

**PASTE.SPECIAL**(paste\_num, operation\_num, skip\_blanks, transpose)

**PASTE.SPECIAL**?(paste\_num, operation\_num, skip\_blanks, transpose)

Paste\_num    is a number from 1 to 6 specifying what to paste.
Paste\_num can also be quoted text of the object you want to paste.

|                |                    |
| -------------- | ------------------ |
| **Paste\_num** | **Pastes**         |
| 1              | All                |
| 2              | Formulas           |
| 3              | Values             |
| 4              | Formats            |
| 5              | Comments           |
| 6              | All except borders |

Operation\_num    is a number from 1 to 5 specifying which operation to
perform when pasting.

|                    |            |
| ------------------ | ---------- |
| **Operation\_num** | **Action** |
| 1                  | None       |
| 2                  | Add        |
| 3                  | Subtract   |
| 4                  | Multiply   |
| 5                  | Divide     |

Skip\_blanks    is a logical value corresponding to the Skip Blanks
check box in the Paste Special dialog box.

  - > If skip\_blanks is TRUE, Microsoft Excel skips blanks in the copy
    > area when pasting.

  - > If skip\_blanks is FALSE, Microsoft Excel pastes normally.

>  

Transpose    is a logical value corresponding to the Transpose check box
in the Paste Special dialog box.

  - > If transpose is TRUE, Microsoft Excel transposes rows and columns
    > when pasting.

  - > If transpose is FALSE, Microsoft Excel pastes normally.

>  

**Related Functions**

FORMULA   Enters values into a cell or range or onto a chart

PASTE   Pastes cut or copied data

PASTE.LINK   Pastes copied data and establishes a link to its source

Syntax 2   Copying from a sheet and pasting into a chart.

Syntax 3   Copying and pasting between charts

Syntax 4   Pasting information from another application.

Return to [top](#H)

# PASTE.SPECIAL Syntax 2

Equivalent to clicking the Paste Special command on the Edit menu on the
Chart menu bar. Pastes the specified components from the copy area into
a chart. The PASTE.SPECIAL function has four syntax forms. Use syntax 2
if you have copied from a sheet and are pasting into a chart.

**Syntax**

**PASTE.SPECIAL**(rowcol, titles, categories, replace, series)

**PASTE.SPECIAL**?(rowcol, titles, categories, replace, series)

Rowcol    is the number 1 or 2 and specifies whether the values
corresponding to a particular data series are in rows or columns. Enter
1 for rows or 2 for columns.

Titles    is a logical value corresponding to the Series Names In First
Column check box (or First Row, depending on the value of rowcol) in the
Paste Special dialog box.

  - > If series is TRUE, Microsoft Excel selects the check box and uses
    > the contents of the cell in the first column of each row (or first
    > row of each column) as the name of the data series in that row (or
    > column).

  - > If series is FALSE, Microsoft Excel clears the check box and uses
    > the contents of the cell in the first column of each row (or first
    > row of each column) as the first data point of the data series.

>  

Categories    is a logical value corresponding to the Categories (X
Labels) In First Row (or First Column, depending on the value of rowcol)
check box in the Paste Special dialog box.

  - > If categories is TRUE, Microsoft Excel selects the check box and
    > uses the contents of the first row (or column) of the selection as
    > the categories for the chart.

  - > If categories is FALSE, Microsoft Excel clears the check box and
    > uses the contents of the first row (or column) as the first data
    > series in the chart.

Replace    is a logical value corresponding to the Replace Existing
Categories check box in the Paste Special dialog box.

  - > If replace is TRUE, Microsoft Excel selects the check box and
    > applies categories while replacing existing categories with
    > information from the copied cell range.

  - > If replace is FALSE, Microsoft Excel clears the check box and
    > applies new categories without replacing any old ones.

Series    is a number specifying how cells are added to a chart.

|            |              |
| ---------- | ------------ |
| **Series** | **Added as** |
| 1          | New series   |
| 2          | New point(s) |

**Related Functions**

Syntax 1   Pasting into a sheet or macro sheet

Syntax 3   Copying and pasting between charts

Syntax 4   Pasting information from another application

Return to [top](#H)

# PASTE.SPECIAL Syntax 3

Equivalent to clicking the Paste Special command on the Edit menu on the
Chart menu bar. Pastes the specified components from the copy area into
a chart. The PASTE.SPECIAL function has four syntax forms. Use syntax 3
if you have copied from a chart and are pasting into a chart.

**Syntax**

**PASTE.SPECIAL**(paste\_num)

**PASTE.SPECIAL**?(paste\_num)

Paste\_num    is a number from 1 to 3 specifying what to paste.

|                |                               |
| -------------- | ----------------------------- |
| **Paste\_num** | **Pastes**                    |
| 1              | All (formats and data series) |
| 2              | Formats only                  |
| 3              | Formulas (data series) only   |

**Related Functions**

Syntax 1   Pasting into a sheet or macro sheet

Syntax 2   Copying from a sheet and pasting into a chart

Syntax 4   Pasting information from another application

Return to [top](#H)

# PASTE.SPECIAL Syntax 4

Equivalent to clicking the Paste Special command on the Edit menu when
pasting data from another application into Microsoft Excel. Use syntax 4
when pasting information from another application.

**Syntax**

**PASTE.SPECIAL**(format\_text, pastelink\_logical,
display\_icon\_logical, icon\_file, icon\_number, icon\_label)

**PASTE.SPECIAL**?(format\_text, pastelink\_logical,
display\_icon\_logical, icon\_file, icon\_number, icon\_label)

Format\_text    is text specifying the type of data you want to paste
from the Clipboard.

  - > The valid data types vary depending on the application from which
    > the data was copied. For example, if you're copying data from
    > Microsoft Word, some of the data types are "Microsoft Document
    > Object", "Picture", and "Text".

  - > For more information about object classes, see your Microsoft
    > Windows or Apple system software documentation.

>  

Pastelink\_logical    is a logical value specifying whether to link the
pasted data to its source application.

  - > If pastelink\_logical is TRUE, Microsoft Excel updates the pasted
    > information whenever it changes in the source application.

  - > If pastelink\_logical is FALSE or omitted, the information is
    > pasted without a link.

  - > If Microsoft Excel or the source application does not support
    > linking for the specified format\_text, then pastelink\_logical is
    > ignored.

Display\_icon\_logical    is a logical value that specifies whether you
want an application's icon to be displayed on the worksheet instead of
the linked data. Equivalent to the Display as Icon check box on the
Paste Special dialog box. If TRUE, the application's icon will be
displayed. If FALSE or omitted, the application's icon will not be
displayed.

Icon\_file    is the name of the file (with an .EXE or .DLL extension)
that contains the icon. If display\_icon\_logical is FALSE, this
argument is ignored.

Icon\_number    is the number associated with the icon and corresponds
to the icon's relative position within the Icon Drop Down list box on
the Change Icon Dialog box, which appears when you click the Change Icon
button in the Paste Special dialog box. If display\_icon\_logical is
FALSE, this argument is ignored.

Icon\_label    is the caption that you want to appear below the icon,
and is equivalent to the Caption text box on the Change Icon dialog box,
which appears when you click the Change Icon button in the Paste Special
dialog box. If display\_icon\_logical is FALSE, this argument is
ignored.

**Related Functions**

Syntax 1   Pasting into a sheet or macro sheet

Syntax 2   Copying from a sheet and pasting into a chart

Syntax 3   Copying and pasting between charts

Return to [top](#H)

# PASTE.TOOL

Pastes a button face from the Clipboard to a specified position on a
toolbar.

**Syntax**

**PASTE.TOOL**(**bar\_id, position**)

Bar\_id    specifies the number or name of the toolbar into which you
want to paste the button face. For detailed information about bar\_id,
see ADD.TOOL.

Position    specifies the position within the toolbar of the button on
which you want to paste the button face. Position starts with 1 at the
left side (if horizontal) or at the top (if vertical).

**Related Function**

COPY.TOOL   Copies a button face

Return to [top](#H)

# PATTERNS

Equivalent to clicking the Patterns tab in the Format Cells dialog box,
which appears when you click the Cells command on the Format menu.
Changes the appearance of the selected cells or objects or the selected
chart item (you can select only one chart item at a time). The PATTERNS
function has eight syntax forms: syntax 1 is for cells on a sheet or
macro sheet. Syntax 2 is for lines or arrows on a worksheet, macro
sheet, or chart. Syntax 3 is for objects on a sheet or macro sheet.
Syntax 4 through syntax 8 are for chart items.

**Syntax 1**

Cells

**PATTERNS**(apattern, afore, aback, newui)

**PATTERNS**?(apattern, afore, aback, newui)

**Syntax 2**

Lines (arrows) on worksheets or charts

**PATTERNS**(lauto, lstyle, lcolor, lwt, hwidth, hlength, htype)

**PATTERNS**?(lauto, lstyle, lcolor, lwt, hwidth, hlength, htype)

**Syntax 3**

Text boxes, rectangles, ovals, arcs, and pictures on worksheets or macro
sheets

**PATTERNS**(bauto, bstyle, bcolor, bwt, shadow, aauto, apattern, afore,
aback, rounded, newui)

**PATTERNS**?(bauto, bstyle, bcolor, bwt, shadow, aauto, apattern,
afore, aback, rounded, newui)

**Syntax 4**

Chart plot areas, bars, columns, pie slices, and text labels

**PATTERNS**(bauto, bstyle, bcolor, bwt, shadow, aauto, apattern, afore,
aback, invert, apply, new\_fill)

**PATTERNS**?(bauto, bstyle, bcolor, bwt, shadow, aauto, apattern,
afore, aback, invert, apply, new\_fill)

**Syntax 5**

Chart axes

**PATTERNS**(lauto, lstyle, lcolor, lwt, tmajor, tminor, tlabel)

**PATTERNS**?(lauto, lstyle, lcolor, lwt, tmajor, tminor, tlabel)

**Syntax 6**

Chart gridlines, hi-lo lines, drop lines, lines on a picture line chart,
and picture charts of bar and column charts

**PATTERNS**(lauto, lstyle, lcolor, lwt, apply, smooth)

**PATTERNS**?(lauto, lstyle, lcolor, lwt, apply, smooth)

**Syntax 7**

Chart data lines

**PATTERNS**(lauto, lstyle, lcolor, lwt, mauto, mstyle, mfore, mback,
apply, smooth)

**PATTERNS**?(lauto, lstyle, lcolor, lwt, mauto, mstyle, mfore, mback,
apply, smooth)

**Syntax 8**

Picture chart markers

**PATTERNS**(type, picture\_units, apply)

**PATTERNS**?(type, picture\_units, apply)

The following argument descriptions are in alphabetic order. Arguments
correspond to check boxes, list boxes, and options in the Patterns tab
of the Format Cells dialog box for the selected item. The default for
each argument reflects the setting in the dialog box.

Aauto    is a number from 0 to 2 specifying area settings (that is, the
object's "surface area").

|                 |                                    |
| --------------- | ---------------------------------- |
| **If aauto is** | **Area settings are**              |
| 0               | Set by the user (custom)           |
| 1               | Automatic (set by Microsoft Excel) |
| 2               | None                               |

Aback    is a number from 1 to 56 corresponding to the 56 area
background colors in the Patterns tab of the Format Cells dialog box.

Afore    is a number from 1 to 56 corresponding to the 56 area
foreground colors in the Patterns tab of the Format Cells dialog box.

Apattern    is a number corresponding to the area patterns in the
Patterns tab of the Format Cells or Format Object dialog box. If an
object is selected, apattern can be from 1 to 18; if a cell is selected,
apattern can be from 0 to 18. If apattern is 0 and a cell is selected,
Microsoft Excel applies no pattern.

Apply    is a logical value corresponding to the Apply To All check box
in Microsoft Excel version 4.0. This argument is for compatibility with
previous versions of Microsoft Excel and applies only when a chart data
point or a data series is selected.

  - > If apply is TRUE, Microsoft Excel applies any formatting changes
    > to all items that are similar to the selected item on the chart.

  - > If apply is FALSE, Microsoft Excel applies formatting changes only
    > to the selected item on the chart.

>  

Bauto    is a number from 0 to 2 specifying border settings.

|                 |                                    |
| --------------- | ---------------------------------- |
| **If bauto is** | **Border settings are**            |
| 0               | Set by the user (custom)           |
| 1               | Automatic (set by Microsoft Excel) |
| 2               | None                               |

Bcolor    is a number from 1 to 56 corresponding to the 56 border colors
in the Border tab of the Format Object or Format (chart element) dialog
box.

Bstyle    is a number from 1 to 8 corresponding to the eight border
styles in the Border tab of the Format Object or Format (chart element)
dialog box.

Bwt    is a number from 1 to 4 corresponding to the four border weights
in the Border tab of the Format Object or Format (chart element) dialog
box.

|               |               |
| ------------- | ------------- |
| **If bwt is** | **Border is** |
| 1             | Hairline      |
| 2             | Thin          |
| 3             | Medium        |
| 4             | Thick         |

Hlength    is a number from 1 to 3 specifying the length of the
arrowhead.

|                   |                  |
| ----------------- | ---------------- |
| **If hlength is** | **Arrowhead is** |
| 1                 | Short            |
| 2                 | Medium           |
| 3                 | Long             |

Htype    is a number from 1 to 5 specifying the style of the arrowhead.

|                 |                           |
| --------------- | ------------------------- |
| **If htype is** | **Style of arrowhead is** |
| 1               | No head                   |
| 2               | Open head                 |
| 3               | Closed head               |
| 4               | Double open head          |
| 5               | Double closed head        |

Hwidth    is a number from 1 to 3 specifying the width of the arrowhead.

|                  |                  |
| ---------------- | ---------------- |
| **If hwidth is** | **Arrowhead is** |
| 1                | Narrow           |
| 2                | Medium           |
| 3                | Wide             |

Invert    is a logical value corresponding to the Invert If Negative
check box in the Patterns tab of the Format Data Series dialog box. This
argument applies only to data markers.

  - > If invert is TRUE, Microsoft Excel inverts the pattern in the
    > selected item if it corresponds to a negative number.

  - > If invert is FALSE, Microsoft Excel removes the inverted pattern,
    > if present, from the selected item corresponding to a negative
    > value.

>  

Lauto    is a number from 0 to 2 specifying line settings.

|                 |                                    |
| --------------- | ---------------------------------- |
| **If lauto is** | **Line settings are**              |
| 0               | Set by the user (custom)           |
| 1               | Automatic (set by Microsoft Excel) |
| 2               | None                               |

Lcolor    is a number from 1 to 56 corresponding to the 16 line colors
in the Patterns tab of the Format Object or Format (chart element)
dialog box.

Lstyle    is a number from 1 to 8 corresponding to the eight line styles
in the Patterns tab of the Format Object or Format (chart element)
dialog box.

Lwt    is a number from 1 to 4 corresponding to the four line weights in
the Patterns tab of the Format Object or Format (chart element) dialog
box.

|               |             |
| ------------- | ----------- |
| **If lwt is** | **Line is** |
| 1             | Hairline    |
| 2             | Thin        |
| 3             | Medium      |
| 4             | Thick       |

Mauto    is a number from 0 to 2 specifying marker settings.

|                 |                                    |
| --------------- | ---------------------------------- |
| **If mauto is** | **Marker settings are**            |
| 0               | Set by the user                    |
| 1               | Automatic (set by Microsoft Excel) |
| 2               | None                               |

Mback    is a number from 1 to 56corresponding to the 56 marker
background colors in the Patterns tab of the Format Data Series dialog
box.

Mfore    is a number from 1 to 56 corresponding to the 56 marker
foreground colors in the Patterns tab of the Format Data Series dialog
box.

Mstyle    is a number from 1 to 9 corresponding to the nine marker
styles in the Patterns tab of the Format Data Series dialog box.

Picture\_units    is the number of units you want each picture to
represent in a scaled, stacked picture chart. This argument applies only
to picture charts and only if type is 3.

Rounded    is a logical value corresponding to the Round Corners check
box and specifying whether to make the corners on text boxes and
rectangles rounded. If rounded is TRUE, the corners are rounded; if
FALSE, the corners are square. If the selection is an arc or an oval,
rounded is ignored.

Newui    is a logical value that specifies whether to use the
foreground, background, and patterns of Microsoft Excel version 5.0 or
later. If TRUE or omitted, the colors and patterns of Microsoft Excel
version 5.0 or later will be used. If FALSE, the colors and patterns of
Microsoft Excel version 4.0 will be used.

Newfill    is a logical value that specifies whether to use the chart
patterns of Microsoft Excel version 5.0 or later. If TRUE or omitted,
the chart patterns of Microsoft Excel version 5.0 or later will be used.
If FALSE, the chart patterns of Microsoft Excel version 4.0 will be
used.

Shadow    is a logical value corresponding to the Shadow check box.
Shadow does not apply to area charts or bars in bar charts. If shadow is
TRUE, Microsoft Excel adds a shadow to the selected item; if FALSE,
Microsoft Excel removes the shadow, if one is present, from the selected
item. If the selection is an arc, shadow is ignored.

Smooth    is a logical value that applies smoothing to picture markers
in line or xy (scatter) charts. The default is FALSE.

Tlabel    is a number from 1 to 4 specifying the position of tick
labels.

|                  |                            |
| ---------------- | -------------------------- |
| **If tlabel is** | **Tick label position is** |
| 1                | None                       |
| 2                | Low                        |
| 3                | High                       |
| 4                | Next to axis               |

Tmajor    is a number from 1 to 4 specifying the type of major tick
marks.

|                  |                                 |
| ---------------- | ------------------------------- |
| **If tmajor is** | **Type of major tick marks is** |
| 1                | None                            |
| 2                | Inside                          |
| 3                | Outside                         |
| 4                | Cross                           |

Tminor    is a number from 1 to 4 specifying the type of minor tick
marks.

|                  |                                 |
| ---------------- | ------------------------------- |
| **If tminor is** | **Type of minor tick marks is** |
| 1                | None                            |
| 2                | Inside                          |
| 3                | Outside                         |
| 4                | Cross                           |

Type    is a number from 1 to 3 specifying the type of pictures to use
in a picture chart.

|                |                                                                                           |
| -------------- | ----------------------------------------------------------------------------------------- |
| **If type is** | **Pictures should be**                                                                    |
| 1              | Stretched to reach a particular value                                                     |
| 2              | Stacked on top of each other to reach a particular value                                  |
| 3              | Stacked on top of each other, but you specify the number of units each picture represents |

**Remarks**

  - > You can select many graphic objects on a sheet or macro sheet and
    > apply formatting to them at the same time, but you can select only
    > one chart item at a time.

  - > If you select multiple objects and if one or more of the objects
    > requires a different form of the PATTERNS function, then choose
    > the syntax corresponding to the object with the most formatting
    > attributes—that is, the syntax with the most arguments. If you
    > specify an argument that does not apply to an item, the argument
    > has no effect on that item.

  - > To apply formatting to similar items on a chart, use the apply
    > argument described above.

>  

**Related Functions**

FONT.PROPERTIES   Applies a font to the selection

FORMAT.TEXT   Formats a text box or a chart text item

Return to [top](#H)

# PAUSE

Pauses a macro. Use the PAUSE function, instead of clicking the Pause
button in the Single Step dialog box, as a debugging tool when you do
not wish to step through a macro. You can also use PAUSE to enter and
edit data, to work directly with Microsoft Excel commands, or to perform
other actions that are not normally available when a macro is running.

**Syntax**

**PAUSE**(no\_tool)

No\_tool    is a logical value specifying whether to display the Resume
Macro button when the macro is paused. If no\_tool is TRUE, the toolbar
is not displayed; if FALSE, the toolbar is displayed; if omitted, the
toolbar is displayed unless you previously clicked the close box on the
toolbar.

**Remarks**

  - > All commands and tools that are available when no macro is running
    > are still available when a macro is paused.

  - > You can run other macros while a macro is paused, but you can
    > pause only one macro at a time. If a macro is paused when you run
    > a second macro containing a PAUSE function, Macro Resume resumes
    > only the second macro; you cannot resume or return to the first
    > macro automatically.

  - > PAUSE is ignored in custom worksheet functions, unless you
    > manually run them by clicking the Run button in the Macro dialog
    > box, which appears when you click the Macro command on the Tools
    > menu. PAUSE is also ignored if it's placed in a formula for which
    > the resume behavior would be unclear, such as:

  - > IF(Cost\<10, AND(PAUSE(),SUM(\!$A$1:$A$4)))

  - > If one macro runs a second macro that pauses, Microsoft Excel
    > locks the calling cell in the first macro. If you try to edit this
    > cell, Microsoft Excel displays an error message.

  - > To resume a paused macro, click the Resume Macro button on the
    > toolbar or run a macro containing a RESUME function.

  - > If one macro runs a second macro that pauses and you need to halt
    > only the paused macro, use RESUME(2) instead of HALT. HALT halts
    > all macros and prevents resuming or returning to any macro. For
    > more information, see RESUME.

**Tip   **Since the automatic Resume Macro button can be customized, you
can create a custom toolbar that will appear whenever a macro pauses.

**Example**

The following macro formula checks to see if a variable named TestValue
is greater than 9. If it is, the macro pauses; otherwise, the macro
continues normally.

IF(TestValue\>9, PAUSE())

**Related Functions**

HALT   Stops all macros from running

RESUME   Resumes a paused macro

STEP   Turns on macro single-stepping

Return to [top](#H)

# PIVOT.ADD.DATA

Adds a field to a PivotTable report.

**Syntax**

**PIVOT.ADD.DATA**(name, pivot\_field\_name, new\_name, position,
function, calculation, base\_field, base\_item, format\_text)

Name    is the name of the PivotTable report to which the user wants to
add as a data field. If name is omitted, Microsoft Excel will use the
PivotTable report containing the active cell.

Pivot\_field\_name    is the name of a field which the user would like
to add to the PivotTable report as data or as text.

New\_name    is the name you would like to give to the new field once it
is added to your PivotTable report. If this argument is omitted,
Microsoft Excel will pick a default name for you. This function returns
new\_name or the name Microsoft Excel gives the field.

Position    is the position within all the Data fields you would like to
place the new data field. If position is omitted, the field will be
added as the last data field.

Function    is a number from 2 to 2048 specifying how the new field is
to be calculated. To compute the value to place in this column use one
value from the following table. If function is omitted, SUM will be
used. If the field is a numeric field or text field, COUNTA will be
used.

|           |              |
| --------- | ------------ |
| **Value** | **Function** |
| 2         | SUM          |
| 4         | COUNTA       |
| 8         | COUNT        |
| 16        | AVERAGE      |
| 32        | MAX          |
| 64        | MIN          |
| 128       | PRODUCT      |
| 256       | STDEV        |
| 512       | STDEVP       |
| 1024      | VAR          |
| 2048      | VARP         |

Calculation    is a number between 1 and 9 representing which custom
calculation you would like to apply to this data field. This corresponds
to the Show Data As drop-down box on the PivotTable Field dialog box. If
this argument is omitted, no special calculation will be applied to the
data field.

|           |                   |
| --------- | ----------------- |
| **Value** | **Calculation**   |
| 1         | Normal            |
| 2         | Difference From   |
| 3         | % Of Item         |
| 4         | % Difference From |
| 5         | Running Total In  |
| 6         | % of Row          |
| 7         | % of Column       |
| 8         | % of Total        |
| 9         | Index             |

Base\_Field    is the field on which you want to base the calculation.

Base\_Item    is the item within base\_field on which you want to base
the calculation.

Format\_text    is the type of number format you want to apply to the
PivotTable data. Corresponds to the number button in the PivotTable
Field dialog box, which appears when you click the PivotTable Field
command on the Data menu when the selection is in a data field.

**Remarks**

  - > If name is not a valid PivotTable name, then the \#VALUE\! error
    > value is returned.

  - > If field\_name is not a valid field for the current PivotTable
    > report then the \#VALUE\! error value is returned.

>  

**Related Functions**

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

PIVOT.FIELD.GROUP   Creates groups within a PivotTable report

PIVOT.FIELD.PROPERTIES   Changes the properties of a field inside a
PivotTable report

PIVOT.FIELD.UNGROUP   Ungroups all selected groups within a PivotTable
report

PIVOT.ITEM   Moves an item within a PivotTable report

PIVOT.ITEM.PROPERTIES   Changes the properties of an item within a
header field

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report

Return to [top](#H)

# PIVOT.ADD.FIELDS

Add fields onto a PivotTable report.

**Syntax**

**PIVOT.ADD.FIELDS**(name, row\_array, column\_array, page\_array,
add\_to\_table)

Name    is the name of the PivotTable report to which the user wants to
add fields. If name is omitted, Microsoft Excel will use the PivotTable
report containing the active cell.

Row\_array    is an array of text constants consisting of the names of
the fields which the user would like to add to the PivotTable report as
row fields.

Column\_array    is an array of text constants consisting of the names
of the fields which the user would like to add to the PivotTable report
as column fields.

Page\_array    is an array of text constants consisting of the names of
the fields which the user would like to add to the PivotTable report as
page fields.

Add\_to\_table    is a logical value which if TRUE adds the fields
specified by row\_array, column\_array and page\_array to the existing
fields on the PivotTable report. If add\_to\_table is FALSE, Microsoft
Excel will replace the fields already along the rows, columns and pages
with the fields specified by row\_array, column\_array and page\_array.

**Remarks**

If name is not a valid PivotTable name, then the \#VALUE\! error value
is returned.

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.FIELD   Pivots fields within a PivotTable report

PIVOT.FIELD.GROUP   Creates groups within a PivotTable report

PIVOT.FIELD.PROPERTIES   Changes the properties of a field inside a
PivotTable report

PIVOT.FIELD.UNGROUP   Ungroups all selected groups within a PivotTable
report

PIVOT.ITEM   Moves an item within a PivotTable report

PIVOT.ITEM.PROPERTIES   Changes the properties of an item within a
header field

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report

Return to [top](#H)

# PIVOT.FIELD

Pivots a field within a PivotTable report.

**Syntax**

**PIVOT.FIELD**(name, pivot\_field\_name, orientation, position)

Name    is the name of the PivotTable report in which the user wants to
pivot fields. If name is omitted, Microsoft Excel will use the
PivotTable report containing the active cell.

Pivot\_field\_name    is the name of the field which the user wishes to
pivot to another part of the PivotTable report. This argument is given
as a text constant or a reference to a text constant. If field\_name is
omitted, Microsoft Excel uses the field containing the active cell.

Orientation    is an integer representing the destination of the field
which is being pivoted. If this argument is omitted, then the
orientation remains unchanged. The integers refer to orientations as
follows:

|           |                 |
| --------- | --------------- |
| **Value** | **Orientation** |
| 0         | Hidden          |
| 1         | Row             |
| 2         | Column          |
| 3         | Page            |
| 4         | Data            |

Position    is an integer representing where in the orientation the
fields will be positioned. Position 1 is the leftmost header position in
the row header and the topmost position in the column header. This
argument is ignored if orientation is set to 0. If the position argument
is omitted, it will default to the last position in the field.

**Remarks**

  - > The function returns TRUE if successful.

  - > If name is not a valid PivotTable name then the \#VALUE\! error
    > value is returned.

  - > If pivot\_field\_name is not a text constant or contains text
    > which is not a valid field name for the PivotTable report then the
    > \#VALUE\! error value is returned.

  - > If destination is not an integer between 0 and 4, then the
    > \#VALUE\! error value is returned.

>  

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD.GROUP   Creates groups within a PivotTable report

PIVOT.FIELD.PROPERTIES   Changes the properties of a field inside a
PivotTable report

PIVOT.FIELD.UNGROUP   Ungroups all selected groups within a PivotTable
report

PIVOT.ITEM   Moves an item within a PivotTable report

PIVOT.ITEM.PROPERTIES   Changes the properties of an item within a
header field

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report

Return to [top](#H)

# PIVOT.FIELD.GROUP

Creates groups within a PivotTable report.

**Syntax**

**PIVOT.FIELD.GROUP**(start, end, by, periods)

**PIVOT.FIELD.GROUP**?(start, end, by, periods)

Start    is the beginning date of the range to be grouped. If start is
TRUE or omitted, it is assumed to be the first value in the field.

End    is the ending date of the range to be grouped. If end is TRUE or
omitted it is assumed to be the last value in the field.

By    is the size of the groups to be created. If by is omitted,
Microsoft Excel uses a default group size. If grouping a date field and
if periods is not 8(days) then by is ignored.

Periods    is a number between 1 and 127. It is calculated by summing
the values in the following table corresponding to the periods into
which you want to group your dates. This argument is ignored if the
field is not a date field. This argument takes precedence over By if
they are both specified for a date field.

|           |             |
| --------- | ----------- |
| **Value** | **Periods** |
| 1         | Seconds     |
| 2         | Minutes     |
| 4         | Hours       |
| 8         | Days        |
| 16        | Months      |
| 32        | Quarters    |
| 64        | Years       |

**Remarks**

  - > This function returns TRUE if the grouping is successful. The
    > \#N/A error value is returned if the grouping failed.

  - > If no arguments are specified and multiple items within the header
    > field are selected then this function groups those selected items.

  - > If no arguments are specified and a single item within the header
    > field is selected then the \#VALUE\! error value is returned.

>  

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

PIVOT.FIELD.PROPERTIES   Changes the properties of a field inside a
PivotTable report

PIVOT.FIELD.UNGROUP   Ungroups all selected groups within a PivotTable
report

PIVOT.ITEM   Moves an item within a PivotTable report

PIVOT.ITEM.PROPERTIES   Changes the properties of an item within a
header field

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report

Return to [top](#H)

# PIVOT.FIELD.PROPERTIES

Changes the properties of a field inside a PivotTable report.

**Syntax**

**PIVOT.FIELD.PROPERTIES**(name, pivot\_field\_name, new\_name,
orientation, function, formats)

**PIVOT.FIELD.PROPERTIES**?(name, pivot\_field\_name, new\_name,
orientation, function, formats)

Name    is the name of the PivotTable report containing the field which
the user wants to edit. If name is omitted, Microsoft Excel will use the
PivotTable report containing the active cell.

Pivot\_field\_name    is the name of a field in the PivotTable report
which the user would like to edit, as text. If it is omitted, Microsoft
Excel uses the field containing the active cell.

New\_name    is the name which you would like to rename the current
field. If it is omitted, the name of the current field will not change.

Orientation    is a number between 0 and 4 specifying which area will
show the field containing the active cell. If zero, then the field is
deleted and all other arguments to this function are ignored. If this
argument is omitted, the orientation of the field will not change.

|           |                   |
| --------- | ----------------- |
| **Value** | **Orientation**   |
| 0         | Delete            |
| 1         | Display as row    |
| 2         | Display as column |
| 3         | Display as page   |
| 4         | Display as data   |

Function    is a number between 0 and 4094 specifying which calculation
or subtotals to apply to the field. If you will be showing the field in
the header (orientation 1, 2, or 3), add up the values from the table
corresponding to the subtotals you would like to show. If you will be
showing the field as a data field (orientation 4), use one value from
the table. If an entry in this column is left blank, Microsoft Excel
will not change the calculation or subtotal which are currently attached
to the field.

|           |              |
| --------- | ------------ |
| **Value** | **Function** |
| 0         | NO SUBTOTALS |
| 1         | AUTOMATIC    |
| 2         | SUM          |
| 4         | COUNTA       |
| 8         | COUNT        |
| 16        | AVERAGE      |
| 32        | MAX          |
| 64        | MIN          |
| 128       | PRODUCT      |
| 256       | STDEV        |
| 512       | STDEVP       |
| 1024      | VAR          |
| 2048      | VARP         |

Formats    is either a one- or a two- dimensional array, depending on
whether the field is a header field or a data field.

  - > If the active field is a header field (orientation argument is 1,
    > 2 or 3) then this is a two-dimensional array. Each row of the
    > array should consist of two entries. The first is a text string
    > corresponding to the item whose property is being changed. The
    > second element specifies whether the item will be hidden. If this
    > argument is TRUE, the item will be hidden and therefore will not
    > be displayed in the PivotTable report. If the argument is FALSE,
    > then the item will be displayed in the PivotTable report.

  - > If the active field is a data field, then the array is a
    > one-dimensional array with four elements. The first element is a
    > number between 1 and 9 specifying which calculation you wish to
    > apply to the current data field. This corresponds to the Show Data
    > As drop-down box on the PivotTable Field dialog box.

|           |                  |
| --------- | ---------------- |
| **Value** | **Format**       |
| 1         | Normal           |
| 2         | Difference From  |
| 3         | %Of Item         |
| 4         | %Difference From |
| 5         | Running Total In |
| 6         | %Of Row          |
| 7         | %Of Column       |
| 8         | %Of Subtotal     |
| 9         | Index            |

 

  - > The second element contains a text string representing the field
    > to which your data field is related. This argument is not
    > necessary for the Normal calculation. If omitted, Microsoft Excel
    > will use the first field that would appear in the Base Field list
    > box.

  - > The third element must contain a text string representing an item
    > in the base field on which to base your calculation. Note that
    > this argument is not necessary for calculations like Running Total
    > In which relies only on a Base Field. If omitted, Microsoft Excel
    > will use the first item that would appear in the Base Item list
    > box.

  - > The fourth element is a text string representing the number format
    > you wish to apply to the data field.

>  

**Remarks**

  - > If pivot\_field\_name is not a valid field name for the PivotTable
    > report then the \#VALUE\! error value is returned.

  - > If name is not a valid PivotTable name, then the \#VALUE\! error
    > value is returned.

  - > If the orientation and function arguments do not contain numbers
    > or if these arguments contain numbers which are out of range then
    > the \#VALUE\! error value is returned.

>  

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

PIVOT.FIELD.GROUP   Creates groups within a PivotTable report

PIVOT.FIELD.UNGROUP   Ungroups all selected groups within a PivotTable
report

PIVOT.ITEM   Moves an item within a PivotTable report

PIVOT.ITEM.PROPERTIES   Changes the properties of an item within a
header field

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report

Return to [top](#H)

# PIVOT.FIELD.UNGROUP

Ungroups all selected groups within a PivotTable report.

**Syntax**

**PIVOT.FIELD.UNGROUP**( )

**Remark**

If the active cell is on a field header, then all the groups in that
field are ungrouped and the field will be removed from the PivotTable
report. Similarly, if the last group in a Parent field is ungrouped, the
entire field will be removed from the PivotTable report.

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

PIVOT.FIELD.GROUP   Creates groups within a PivotTable report

PIVOT.FIELD.PROPERTIES   Changes the properties of a field inside a
PivotTable report

PIVOT.ITEM   Moves an item within a PivotTable report

PIVOT.ITEM.PROPERTIES   Changes the properties of an item within a
header field

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report

Return to [top](#H)

# PIVOT.ITEM

Moves an item within a PivotTable report.

**Syntax**

**PIVOT.ITEM**(name, pivot\_field\_name, pivot\_item\_name,position)

Name    is the name of the PivotTable report within which an item will
be repositioned. If name is omitted, Microsoft Excel will use the
PivotTable report containing the active cell.

Pivot\_field\_name    is the name of the field within which an item will
be repositioned, given as a text string. If pivot\_field\_name is
omitted, Microsoft Excel will use the field containing the active cell.
If the active cell is not within a field, then this argument is
required.

Pivot\_item\_name    is the name of the item to be repositioned in its
field (given as a text constant). If it is omitted, Microsoft Excel uses
the item containing the active cell. If the active cell is not contained
within an item, then this argument is required.

Position    is a number representing where in the field the items will
be moved. Position 1 is the topmost position within the row field and
the leftmost position within the column field and the highest position
within the page field. If the position argument is omitted, it will
default to the last position in the field.

**Remarks**

  - > If an item is set to be visible, but its display is suppressed
    > because there is no data, this item still occupies a valid
    > position.

  - > If name is not a valid PivotTable name then the \#VALUE\! error
    > value is returned.

  - > If pivot\_field\_name is not a text string, or if
    > pivot\_field\_name is not a text string within a valid field name,
    > then \#VALUE\! is returned.

  - > If pivot\_item\_name is an item which is not currently showing in
    > the PivotTable report because it does not exist in the field
    > pivot\_field\_name, the \#VALUE\! error value is returned.

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

PIVOT.FIELD.GROUP   Creates groups within a PivotTable report

PIVOT.FIELD.PROPERTIES   Changes the properties of a field inside a
PivotTable report

PIVOT.FIELD.UNGROUP   Ungroups all selected groups within a PivotTable
report

PIVOT.ITEM.PROPERTIES   Changes the properties of an item within a
header field

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report

Return to [top](#H)

# PIVOT.ITEM.PROPERTIES

Changes the properties of an item within a header field.

**Syntax**

**PIVOT.ITEM.PROPERTIES**(name, pivot\_field\_name, pivot\_item\_name,
new\_name, position, show, active\_page)

Name    is the name of the PivotTable report containing the item which
the user wants to edit.

Pivot\_field\_name    is the name of a field in the PivotTable report
containing the item which the user would like to edit. If it is omitted,
Microsoft Excel uses the field containing the active cell.

Pivot\_item\_name    is the name of the item which the user would like
to edit. If it is omitted, Microsoft Excel uses the item containing the
active cell.

New\_name    is the name which you would like to rename the current
item. If it is omitted, then the name of the current item will not
change.

Position    is a number representing where in the field the item will
appear. Position 1 is the topmost position within the row field, the
leftmost position within the column field, and the highest position
within the page field. If the position argument is omitted, it will
default to the last position in the field.

Show    is a logical value which if TRUE causes the item to appear in
the PivotTable report. If FALSE, the item will be hidden.

Active\_page    is a logical value specifying whether item\_name will
become the active item in the page field. If TRUE then item\_name will
become the active item in the page field. If FALSE or omitted, the
active item in the page field does not change. Applies to only page
fields.

**Remarks**

  - > If name is omitted, Microsoft Excel will use the PivotTable report
    > containing the active cell.

  - > If pivot\_field\_name is not a header field, then this function
    > will return the \#VALUE\! error value.

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

PIVOT.FIELD.GROUP   Creates groups within a PivotTable report

PIVOT.FIELD.PROPERTIES   Changes the properties of a field inside a
PivotTable report

PIVOT.FIELD.UNGROUP   Ungroups all selected groups within a PivotTable
report

PIVOT.ITEM   Moves an item within a PivotTable report

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report

Return to [top](#H)

# PIVOT.REFRESH

Refreshes a PivotTable report.

**Syntax**

**PIVOT.REFRESH**(name)

Name    is the name of the PivotTable report the user would like to
refresh with current data. If name is omitted, Microsoft Excel will use
the PivotTable report containing the active cell.

**Remarks**

  - > If the function is successful, it returns TRUE; otherwise, it
    > returns the \#VALUE\! error value.

  - > If name is not a valid PivotTable name, then the \#VALUE\! error
    > value is returned.

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

PIVOT.FIELD.GROUP   Creates groups within a PivotTable report

PIVOT.FIELD.PROPERTIES   Changes the properties of a field inside a
PivotTable report

PIVOT.FIELD.UNGROUP   Ungroups all selected groups within a PivotTable
report

PIVOT.ITEM   Moves an item within a PivotTable report

PIVOT.ITEM.PROPERTIES   Changes the properties of an item within a
header field

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report

Return to [top](#H)

# PIVOT.SHOW.PAGES

Creates new sheets in the workbook containing the active cell. The
function will iterate through each item in page\_field and create a new
PivotTable report on a new sheet with the page field set to that
particular item.

**Syntax**

**PIVOT.SHOW.PAGES**(name, page\_field)

Name    is the name of the target PivotTable report. If name is omitted,
Microsoft Excel will use the PivotTable report containing the active
cell.

Page\_field    is the name of a page field in the PivotTable report
specified by the name argument.

**Remarks**

  - > If the function is successful, it returns TRUE; otherwise, it
    > returns the \#VALUE\! error value.

  - > If name is not a valid PivotTable name then the \#VALUE\! error
    > value is returned.

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

PIVOT.FIELD.GROUP   Creates groups within a PivotTable report

PIVOT.FIELD.PROPERTIES   Changes the properties of a field inside a
PivotTable report

PIVOT.FIELD.UNGROUP   Ungroups all selected groups within a PivotTable
report

PIVOT.ITEM   Moves an item within a PivotTable report

PIVOT.ITEM.PROPERTIES   Changes the properties of an item within a
header field

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.TABLE.WIZARD   Creates an empty PivotTable report

Return to [top](#H)

# PIVOT.TABLE.WIZARD

Creates an empty PivotTable report.

**Syntax**

**PIVOT.TABLE.WIZARD**(type, source, destination, name, row\_grand,
col\_grand, save\_data, apply\_auto\_format, autopage)

**PIVOT.TABLE.WIZARD**?(type, source, destination, name, row\_grand,
col\_grand, save\_data, apply\_auto\_format, autopage)

Type    is a number specifying the type of source data used to create
the PivotTable report.

|           |                                  |
| --------- | -------------------------------- |
| **Value** | **Type of source data**          |
| 1         | Microsoft Excel list or database |
| 2         | External data source             |
| 3         | Multiple consolidation ranges    |
| 4         | Another PivotTable report        |

Source    can be one of four things. If type is 1, then source is a cell
reference or name to the range to be used as the PivotTable source. If
type is 2, then source is a one-dimensional array describing the
external database to be used as the PivotTable source. If type is 3,
then source is a multi-dimensional array listing the cell ranges and
associated page field items describing the consolidation PivotTable
source. If type is 4, then source is the name of another PivotTable
report with which to share its source.

Destination    is a cell reference or name. The upper-left corner of
this range will act as the upper-left corner of the PivotTable report
which will be created. If destination is omitted, Microsoft Excel will
create the PivotTable report on a new sheet.

Name    is the name of the PivotTable report to be created given as a
text. If name is omitted, Microsoft Excel uses a default name.

Row\_grand    is a logical value which if TRUE displays a Grand Total
for each row on the PivotTable report. If FALSE, a Grand Total for each
row is not displayed.

Col\_grand    is a logical value which if TRUE displays a Grand Total
for each column. If FALSE, a Grand Total for each column is not
displayed.

Save\_data    is a logical value which if TRUE causes the data for the
PivotTable report to be saved along with the PivotTable definition. If
FALSE, the data is not saved along with the PivotTable definition.

Apply\_auto\_format    is a logical value which if TRUE causes
autoformatting upon pivotting or refreshing. If FALSE, the PivotTable
report will not be formatted automatically upon pivoting or refreshing.

Autopage    Applies only to type 3. This argument is a logical value
which if TRUE or omitted causes Microsoft Excel to create a page field
automatically. If FALSE, the page field must be created manually.

**Remarks**

  - > The function will return TRUE if successful; otherwise, returns
    > the \#VALUE\! error value.

  - > If destination is not a valid Microsoft Excel reference, then
    > \#VALUE\! error value is returned.

  - > If name is not a valid PivotTable name, then the \#VALUE\! error
    > value is returned.

**Related Functions**

PIVOT.ADD.DATA   Adds a field to a PivotTable report as a data field

PIVOT.ADD.FIELDS   Adds fields to a PivotTable report

PIVOT.FIELD   Pivots fields within a PivotTable report

PIVOT.FIELD.GROUP   Creates groups within a PivotTable report

PIVOT.FIELD.PROPERTIES   Changes the properties of a field inside a
PivotTable report

PIVOT.FIELD.UNGROUP   Ungroups all selected groups within a PivotTable
report

PIVOT.ITEM   Moves an item within a PivotTable report

PIVOT.ITEM.PROPERTIES   Changes the properties of an item within a
header field

PIVOT.REFRESH   Refreshes a PivotTable report

PIVOT.SHOW.PAGES   Creates new sheets in the workbook containing the
active cell

Return to [top](#H)

# PLACEMENT

Equivalent to choosing the Object Placement command from the Format menu
in Microsoft Excel version 3.0. Determines how the selected object or
objects are attached to the cells beneath them. This function is
included only for macro compatibility and will be converted to
OBJECT.PROPERTIES when you load older macro sheets. For more
information, see OBJECT.PROPERTIES.

**Syntax**

**PLACEMENT**(placement\_type)

**PLACEMENT**?(placement\_type)

**Related Functions**

OBJECT.PROPERTIES   Determines an object's relationship to underlying
cells

Return to [top](#H)

# POKE

Sends data to another application. Use POKE to send data to documents in
other applications you are communicating with through dynamic data
exchange (DDE).

**Syntax**

**POKE**(**channel\_num, item\_text**, **data\_ref**)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for this function.

Channel\_num    is the channel number returned by a previously run
INITIATE function.

Item\_text    is text that identifies the item you want to send data to
in the application you are accessing through channel\_num. The form of
item\_text depends on the application connected to channel\_num.

Data\_ref    is a reference to the workbook containing the data to send.

If POKE is not successful, it returns the following values.

|                    |                                                                                                                 |
| ------------------ | --------------------------------------------------------------------------------------------------------------- |
| **Value returned** | **Meaning**                                                                                                     |
| \#VALUE\!          | Channel\_num is not a valid channel number.                                                                     |
| \#DIV/0\!          | The application you are accessing does not respond after a certain length of time, and you press ESC to cancel. |
| \#REF\!            | POKE is refused.                                                                                                |

**Examples**

In Microsoft Excel for Windows, the following macro inserts the text
from cell C3 into the Microsoft Word for Windows document SALES.DOC at
the start of the document.

\=POKE(SendChanl, "StartOfDoc", C3)

In Microsoft Excel for the Macintosh, the following macro inserts the
text from cell C3 into the Microsoft Word document named Report.

\=POKE(SendChanl, "TopicName", C3)

**Related Functions**

INITIATE   Opens a channel to another application

REQUEST   Returns data from another application

TERMINATE   Closes a channel to another application

Return to [top](#H)

# PRECISION

Equivalent to selecting or clearing the Precision As Displayed check box
in the Calculation tab of the Options dialog box, which appears when you
click the Options command on the Tools menu. Controls how values are
stored in cells. Use PRECISION when the results of formulas do not seem
to match the values used to calculate the formulas.

**Syntax**

**PRECISION**(logical)

Logical    is a logical value corresponding to the Precision As
Displayed check box in the Calculation tab.

  - > If logical is TRUE, Microsoft Excel stores future entries at full
    > precision (15 digits).

  - > If logical is FALSE or omitted, Microsoft Excel stores values
    > exactly as they are displayed.

>  

**Caution   **The PRECISION function may permanently alter your data.
PRECISION(FALSE) causes Microsoft Excel to change values on your
worksheet or macro sheet to match displayed values. PRECISION(TRUE)
causes Microsoft Excel to store future values at full precision, but it
does not restore previously entered numbers to their original values.

**Remarks**

  - > Precision As Displayed does not affect numbers in General format.
    > Numbers in General format are always calculated to full precision.

  - > Microsoft Excel calculates slightly faster when using full
    > precision because with Precision As Displayed selected, Microsoft
    > Excel has to round off numbers as it calculates.

>  

**Related Functions**

FORMAT.NUMBER   Applies a number format to the selection

WORKSPACE   Changes workspace settings

Return to [top](#H)

# PREFERRED

Equivalent to clicking the Preferred command on the Gallery menu in
Microsoft Excel version 4.0. Changes the format of the active chart to
the format currently defined by the Set As Default Chart option in the
Standard Types tab of the Chart Type dialog box or the SET.PREFERRED
macro function.

**Syntax**

**PREFERRED**( )

**Related Function**

SET.PREFERRED   Changes the default chart format

Return to [top](#H)

# PRESS.TOOL

Formats a button so that it appears either normal or depressed into the
screen.

**Syntax**

**PRESS.TOOL**(**bar\_id, position**, down)

Bar\_id    specifies the number or name of the toolbar in which you want
to change the button appearance. For detailed information about bar\_id,
see ADD.TOOL.

Position    specifies the position of the button within the toolbar.
Position starts with 1 at the left side (if horizontal) or at the top
(if vertical).

Down    is a logical value specifying the appearance of the button. If
down is TRUE, the button appears depressed into the screen; if FALSE or
omitted, it appears normal (up).

**Remarks**

This function applies only to custom buttons to which macros have
already been assigned. An error occurs if you try to process any other
type of button.

**Example**

The following macro formula sets the third button image on Toolbar4 to
normal (up).

PRESS.TOOL("Toolbar4", 3, FALSE)

**Related Functions**

ADD.TOOL   Adds one or more buttons to a toolbar

DELETE.TOOL   Deletes a button from a toolbar

Return to [top](#H)

# PRINT

Equivalent to clicking the Print command on the File menu. Prints the
active workbook.

Arguments correspond to options, check boxes, and edit boxes in the
Print dialog box. Arguments corresponding to check boxes are logical
values. If an argument is TRUE, Microsoft Excel selects the check box;
if FALSE, Microsoft Excel clears the check box.

**Syntax**

**PRINT**(range\_num, from, to, copies, draft, preview, print\_what,
color, feed, quality, y\_resolution, selection)

**PRINT**?(range\_num, from, to, copies, draft, preview, print\_what,
color, feed, quality, y\_resolution, selection)

Range\_num    is a number specifying which pages to print.

|                |                                                                                        |
| -------------- | -------------------------------------------------------------------------------------- |
| **Range\_num** | **Prints the following pages**                                                         |
| 1              | All the pages                                                                          |
| 2              | Prints a specified range. If range\_num is 2, then from and to are required arguments. |

From    specifies the first page to print. This argument is ignored
unless range\_num equals 2.

To    specifies the last page to print. This argument is ignored unless
range\_num equals 2.

Copies    specifies the number of copies to print. If omitted, the
default is 1.

Draft    This argument overrides the draft argument from the PAGE.SETUP
function. If omitted, the Draft Setting from the Page.Setup function is
used.

Preview    is a logical value corresponding to the Print Preview button
in the Print dialog box. If TRUE, the print preview window will be
displayed. If FALSE, the window will not be displayed

Print\_what    is a number from 1 to 3 that specifies what parts of the
sheet or macro sheet to print. If a chart is active, print\_what is
ignored. This argument will override the setting in the Page Setup
dialog box. If omitted, the note argument in the Page.Setup function is
used to determine whether to print notes or not.

|                 |                      |
| --------------- | -------------------- |
| **Print\_what** | **Prints**           |
| 1               | Sheet only           |
| 2               | Notes only           |
| 3               | Sheet and then notes |

Color    corresponds to the Print Using Color check box. Color is
available only in Microsoft Excel for the Macintosh. If omitted, the
setting is not changed.

Feed    is a number specifying the type of paper feed. Feed is available
only in Microsoft Excel for the Macintosh.

|              |                                   |
| ------------ | --------------------------------- |
| **Feed**     | **Type of paper feed**            |
| 1 or omitted | Continuous (paper cassette)       |
| 2            | Cut sheet or manual (manual feed) |

Quality    Specifies the DPI output quality you want. If omitted, the
corresponding settings in the Page Setup dialog box will be used. If
included, this argument overrides the quality argument in the PAGE SETUP
dialog box.

Y\_resolution    corresponds to the Print Quality box in the Page Setup
dialog box if you have specified a printer where the horizontal and
vertical resolution are not equal, such as a dot-matrix printer. If
omitted, the corresponding settings in the Page Setup dialog box will be
used. If included, this argument overrides the print quality setting in
the Page Setup dialog box.

Selection    specifies what portion of the sheet to print.

|               |                                                                                                                                                                         |
| ------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Selection** | **Portion printed**                                                                                                                                                     |
| 1             | Prints the current selection from all selected sheets. For example, if A1:F40 is selected on the active sheet, A1:F40 will be printed from each of the selected sheets. |
| 2             | Prints the print area or entire sheet from all selected sheets.                                                                                                         |
| 3             | Prints print area or entire sheet from all sheets in the workbook.                                                                                                      |

**Related Functions**

PAGE.SETUP   Sets page printing specifications

PRINT.PREVIEW   Previews pages and page breaks before printing

PRINTER.SETUP   Identifies the printer

SET.PRINT.AREA   Defines the print area

SET.PRINT.TITLES   Defines text to print as titles

DEFINE.NAME   Equivalent to clicking the Define command on the Name
submenu of the Insert menu

Return to [top](#H)

# PRINTER.SETUP

Use PRINTER.SETUP to change the printer you are using.

**Syntax**

**PRINTER.SETUP**(**printer\_text**)

**PRINTER.SETUP**?(printer\_text)

Printer\_text    is the name of the printer you want to switch to. Enter
printer\_text exactly as it appears in the Setup dialog box.

**Note   **This function is only available in Microsoft Excel for
Windows.

**Related Functions**

PAGE.SETUP   Sets page printing specifications

PRINT   Prints the active workbook

Return to [top](#H)

# PRINT.PREVIEW

Equivalent to clicking the Print Preview command on the File menu.
Previews the pages and page breaks of the active workbook on the screen
before printing.

**Syntax**

**PRINT.PREVIEW**( )

**Related Function**

PRINT   Prints the active workbook

Return to [top](#H)

# PROMOTE

Equivalent to clicking the Ungroup button. Promotes, or ungroups, the
currently selected rows or columns in an outline. Use PROMOTE to change
the configuration of an outline by promoting rows or columns of
information.

**Syntax**

**PROMOTE**(rowcol)

**PROMOTE**?(rowcol)

Rowcol    specifies whether to promote rows or columns.

|              |             |
| ------------ | ----------- |
| **Rowcol**   | **Demotes** |
| 1 or omitted | Rows        |
| 2            | Columns     |

**Remarks**

  - > If the selection consists of an entire row or rows, then rows are
    > promoted even if rowcol is 2. Similarly, selection of an entire
    > column overrides rowcol 1.

  - > Also, if the selection is unambiguous (an entire row or column),
    > then PROMOTE? will not display the dialog box.

>  

**Related Functions**

DEMOTE   Demotes the selection in an outline

SHOW.DETAIL   Expands or collapses a portion of an outline

SHOW.LEVELS   Displays a specific number of levels of an outline

Return to [top](#H)

# PROTECT.DOCUMENT

Adds protection to or removes protection from the active sheet, macro
sheet, chart, dialog sheet, module, or scenario. Use PROTECT.DOCUMENT to
prevent yourself or others from changing cell contents, or objects in a
workbook. To protect workbooks in Microsoft Excel version 5.0 or later,
see WORKBOOK.PROTECT.

**Syntax**

**PROTECT.DOCUMENT**(contents, windows, password, objects, scenarios)

**PROTECT.DOCUMENT**?( contents, windows, password, objects, scenarios)

Contents    is a logical value corresponding to the Contents check box
in the Protect Sheet dialog box.

  - > If contents is TRUE or omitted, Microsoft Excel selects the check
    > box and protects cells and chart elements on the sheet or macro
    > sheet.

  - > If contents is FALSE, Microsoft Excel clears the check box (and
    > removes protection if the correct password is supplied).

>  

Windows    is provided for compatibility with Microsoft Excel version
4.0. To protect the window placement and structure of workbooks in
Microsoft Excel version 5.0 or later, see WORKBOOK.PROTECT.

  - > If windows is TRUE, Microsoft Excel prevents a workbook's windows
    > from being moved or sized.

  - > If windows is FALSE or omitted, Microsoft Excel removes protection
    > if the correct password is supplied.

>  

Password    is the password you specify in the form of text to protect
or unprotect the file. Password is case-sensitive.

  - > If password is omitted when you protect a sheet, then you will be
    > able to remove protection without a password. This is useful if
    > you want only to protect the sheet from accidental changes.

  - > If password is omitted when you try to remove protection from a
    > sheet that was protected with a password, the normal password
    > dialog box is displayed.

  - > Passwords are not recorded into the PROTECT.DOCUMENT function when
    > you use the macro recorder.

>  

Objects    is a logical value. This argument applies only to charts,
worksheets and macro sheets. Objects corresponds to the Objects check
box in the Protect Sheet dialog box.

  - > If objects is TRUE or omitted, Microsoft Excel selects the check
    > box and protects all locked objects on the chart, worksheet or
    > macro sheet.

  - > If objects is FALSE, Microsoft Excel clears the check box.

Scenarios    is a logical value that corresponds to the Scenarios check
box on the Protect Sheet dialog box. If TRUE, Microsoft Excel protects
all the scenarios. If FALSE, the scenarios are not protected.

**Remarks**

  - > If contents and objects are FALSE, PROTECT.DOCUMENT carries out
    > the Unprotect Sheet command. If contents, or objects is TRUE, it
    > carries out the Protect Sheet command.

  - > Make sure that you hide macro sheets that protect or unprotect
    > worksheets. If you type a password directly into the function on
    > an unhidden macro sheet, then someone could see the password
    > needed to unprotect the worksheet. For example,
    > PROTECT.DOCUMENT(TRUE, TRUE, "XD1411C", TRUE).

>  

**Warning**   If you forget the password of a workbook that was
previously protected with a password, you cannot unprotect the workbook.

**Related Functions**

CELL.PROTECTION   Controls protection for the selected cells

ENTER.DATA   Turns Data Entry mode on and off

OBJECT.PROTECTION   Controls how an object is protected

SAVE.AS   Saves a workbook and allows you to specify the name, file
type, password, backup file, and location of the workbook

WORKBOOK.PROTECT   Protects a workbook

Return to [top](#H)

# PTTESTM

Performs a paired two-sample Student's t-Test for means.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**PTTESTM**(**inprng1, inprng2**, outrng, labels, alpha, difference)

**PTTESTM**?(inprng1, inprng2, outrng, labels, alpha, difference)

Inprng1    is the input range for the first data set.

Inprng2    is the input range for the second data set.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Labels    is a logical value.

  - > If labels is TRUE, then labels are in the first row or column of
    > the input ranges.

  - > If labels is FALSE or omitted, all cells in inprng1 and inprng2
    > are considered data. The output table will include default row or
    > column headings.

>  

Alpha    is the confidence level for the test. If omitted, alpha is
0.05.

Difference    is the hypothesized mean difference. If omitted,
difference is 0.

**Related Functions**

PTTESTV   Performs a two-sample Student's t-Test, assuming unequal
variances

TTESTM   Performs a two-sample Student's t-Test for means, assuming
equal variances

Return to [top](#H)

# PTTESTV

Performs a two-sample Student's t-Test, assuming unequal variances.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**PTTESTV**(**inprng1, inprng2**, outrng, labels, alpha)

**PTTESTV**?(inprng1, inprng2, outrng, labels, alpha)

Inprng1    is the input range for the first data set.

Inprng2    is the input range for the second data set.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Labels    is a logical value.

  - > If labels is TRUE, then labels are in the first row or column of
    > the input ranges.

  - > If labels is FALSE or omitted, all cells in inprng1 and inprng2
    > are considered data. The output table will include default row or
    > column headings.

>  

Alpha    is the confidence level for the test. If omitted, alpha is
0.05.

**Related Functions**

PTTESTM   Performs a paired two-sample Student's t-Test for means

TTESTM   Performs a two-sample Student's t-Test for means, assuming
equal variances

Return to [top](#H)

# PUSHBUTTON.PROPERTIES

Sets the properties of the push button control on a worksheet or dialog
sheet.

**Syntax**

**PUSHBUTTON.PROPERTIES**(default\_logical, cancel\_logical,
dismiss\_logical, help\_logical, accel\_text, accel\_text2)

**PUSHBUTTON.PROPERTIES**?(default\_logical, cancel\_logical,
dismiss\_logical, help\_logical, accel\_text, accel\_text2)

Default\_logical    is a logical value that determines whether the
button is the default button for the dialog. If TRUE, the button is the
default button. If FALSE, the button is not the default button for the
control.

Cancel\_logical    is a logical value that determines whether the button
is activated when the dialog is cancelled with the Close button or the
ESC key. If TRUE, the button is activated when the dialog box is
cancelled, and edit boxes are not checked to see if they contain valid
data types. If FALSE, the button is not activated when the dialog box is
cancelled.

Dismiss\_logical    is a logical value that determines whether the
button dismisses the dialog when pressed, as when the user presses the
box's OK button. If TRUE, the button dismisses the dialog box. If FALSE,
the button does not dismiss the dialog box.

Help\_logical    is a logical value that determines whether the button
is activated when the user presses the F1 key. If TRUE, the button is
activated when the user presses the F1 key. If FALSE, the button is not
activated when the user presses the F1 key.

Accel\_text    is a text string containing the character to use as the
dialog button's accelerator key. The character is matched against the
text of the control, and the first matching character is underlined.
When the user presses ALT+accel\_text in Microsoft Excel for Windows or
COMMAND+accel\_text in Microsoft Excel for the Macintosh, the control is
clicked. This argument is ignored for push button controls on
worksheets.

Accel\_text2    is a text string containing the second accelerator key
on a dialog sheet. This argument is for only Far East versions of
Microsoft Excel.

**Related Functions**

CHECKBOX.PROPERTIES   Sets various properties of check box and option
box controls

SCROLLBAR.PROPERTIES   Sets the properties of the scroll bar and spinner
controls

EDITBOX.PROPERTIES   Sets the properties of an edit box on a worksheet
or dialog sheet

Return to [top](#H)
