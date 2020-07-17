<span id="A" class="anchor"></span>This document contains reference
information on the following Excel macro functions:

# A

[A1.R1C1](#a1.r1c1), [ABSREF](#absref), [ACTIVATE](#activate),
ACTIVATE.NEXT, ACTIVATE.PREV, [ACTIVE.CELL](#active.cell),
[ACTIVE.CELL.FONT](#active.cell.font), [ADD.ARROW](#add.arrow),
[ADD.BAR](#add.bar), [ADD.CHART.AUTOFORMAT](#add.chart.autoformat),
[ADD.COMMAND](#add.command), [ADDIN.MANAGER](#addin.manager),
[ADD.LIST.ITEM](#add.list.item), [ADD.MENU](#add.menu),
[ADD.OVERLAY](#add.overlay), [ADD.TOOL](#add.tool),
[ADD.TOOLBAR](#add.toolbar), [ALERT](#alert), [ALIGNMENT](#alignment),
[ANOVA1](#anova1), [ANOVA2](#anova2), [ANOVA3](#anova3),
[APP.ACTIVATE](#app.activate),
[APP.ACTIVATE.MICROSOFT](#app.activate.microsoft),
[APPLY.NAMES](#apply.names), [APPLY.STYLE](#apply.style),
[APP.MAXIMIZE](#app.maximize), [APP.MINIMIZE](#app.minimize),
[APP.MOVE](#app.move), [APP.RESTORE](#app.restore),
[APP.SIZE](#app.size), [APP.TITLE](#app.title), [ARGUMENT](#argument),
[ARRANGE.ALL](#arrange.all), [ASSIGN.TO.OBJECT](#assign.to.object),
[ASSIGN.TO.TOOL](#assign.to.tool), [ATTACH.TEXT](#attach.text),
[ATTACH.TOOLBARS](#attach.toolbars), [AUTO.OUTLINE](#auto.outline),
[AXES](#axes)

# B

[BEEP](#beep), [BORDER](#border), [BREAK](#break),
[BRING.TO.FRONT](#bring.to.front)

# C

[CALCULATE.DOCUMENT](#calculate.document),
[CALCULATE.NOW](#calculate.now), [CALCULATION](#calculation),
[CALLER](#caller), [CANCEL.COPY](#cancel.copy),
[CANCEL.KEY](#cancel.key), [CELL.PROTECTION](#cell.protection),
[CHANGE.LINK](#change.link), [CHART.ADD.DATA](#chart.add.data),
[CHART.TREND](#chart.trend), [CHART.WIZARD](#chart.wizard),
[CHECKBOX.PROPERTIES](#checkbox.properties),
[CHECK.COMMAND](#check.command), [CLEAR](#clear),
[CLEAR.OUTLINE](#clear.outline),
[CLEAR.ROUTING.SLIP](#clear.routing.slip), [CLOSE](#close),
[CLOSE.ALL](#close.all), [COLOR.PALETTE](#color.palette),
[COLUMN.WIDTH](#column.width), [COMBINATION](#combination),
[CONSOLIDATE](#consolidate), [CONSTRAIN.NUMERIC](#constrain.numeric),
[COPY](#copy), [COPY.CHART](#copy.chart), [COPY.PICTURE](#copy.picture),
[COPY.TOOL](#copy.tool), [CREATE.NAMES](#create.names),
[CREATE.OBJECT](#create.object), [CREATE.PUBLISHER](#create.publisher),
[CUSTOMIZE.TOOLBAR](#customize.toolbar),
[CUSTOM.REPEAT](#custom.repeat), [CUSTOM.UNDO](#custom.undo),
[CUT](#cut)

# D

[DATA.DELETE](#data.delete), [DATA.FIND](#data.find), [DATA.FIND.NEXT,,
DATA.FIND.PREV](#data.find.next-data.find.prev),
[DATA.FORM](#data.form), [DATA.LABEL](#data.label),
[DATA.SERIES](#data.series), [DEFINE.NAME](#define.name),
[DEFINE.STYLE](#define.style), [DEFINE.STYLE, Syntax,
1](#define.style-syntax-1), [DEFINE.STYLE, Syntaxes, 2, -,
7](#define.style-syntaxes-2---7), [DELETE.ARROW](#delete.arrow),
[DELETE.BAR](#delete.bar),
[DELETE.CHART.AUTOFORMAT](#delete.chart.autoformat),
[DELETE.COMMAND](#delete.command), [DELETE.FORMAT](#delete.format),
[DELETE.MENU](#delete.menu), [DELETE.NAME](#delete.name),
[DELETE.OVERLAY](#delete.overlay), [DELETE.STYLE](#delete.style),
[DELETE.TOOL](#delete.tool), [DELETE.TOOLBAR](#delete.toolbar),
[DEMOTE](#demote), [DEREF](#deref), [DESCR](#descr),
[DIALOG.BOX](#dialog.box), [DIRECTORY](#directory),
[DISABLE.INPUT](#disable.input), [DISPLAY](#display), [DISPLAY, Syntax,
1](#display-syntax-1), [DISPLAY, Syntax, 2](#display-syntax-2),
[DOCUMENTS](#documents), [DUPLICATE](#duplicate)

# A1.R1C1

Displays row and column headings and cell references in either the R1C1
or A1 reference style. A1 is the Microsoft Excel default reference
style.

**Syntax**

**A1.R1C1**(**logical**)

Logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying which
reference style to use. If logical is TRUE, all worksheets and macro
sheets use A1 references; if FALSE, all worksheets and macro sheets use
R1C1 references.

**Example**

The following macro formula displays an alert box asking you to select
either A1 or R1C1 reference style. This is useful in an Auto\_Open macro
if several persons who prefer different reference styles must maintain
the same workbook.

A1.R1C1(ALERT("Click OK for A1 style; Cancel for R1C1", 1))

Return to [top](#A)

# ABSREF

Returns the absolute reference of the cells that are offset from a
reference by a specified amount. You should generally use OFFSET instead
of ABSREF. This function is provided for users who prefer to supply an
absolute reference in text form.

**Syntax**

**ABSREF**(**ref\_text**, **reference**)

Ref\_text&nbsp;&nbsp;&nbsp;&nbsp;specifies a position relative to
reference. Think of ref\_text as "directions" from one range of cells to
another.

  - > Ref\_text must be an R1C1-style relative reference in the form of
    > text, such as "R\[1\]C\[1\]".

  - > Ref\_text is considered relative to the cell in the upper-left
    > corner of reference.

> &nbsp;

Reference&nbsp;&nbsp;&nbsp;&nbsp;is a cell or range of cells specifying
a starting point that ref\_text uses to locate another range of cells.
Reference can be an external reference.

**Remarks**

  - > If you use ABSREF in a function or operation, you will usually get
    > the values contained in the reference instead of the reference
    > itself because the reference is automatically converted to the
    > contents of the reference.

  - > If you use ABSREF in a function that requires a reference
    > argument, then Microsoft Excel does not convert the reference to a
    > value.

  - > If you want to work with the actual reference, use the REFTEXT
    > function to convert the active-cell reference to text, which you
    > can then store or manipulate (or convert back to a reference with
    > TEXTREF). See the third example following.

> &nbsp;

**Examples**

ABSREF("R\[-2\]C\[-2\]", C3) equals $A$1

ABSREF(RELREF(A1, C3), D4) equals $B$2

REFTEXT(ABSREF("R\[-2\]C\[-2\]:R\[2\]C\[2\]", C3:G7), TRUE) is
equivalent to

REFTEXT(ABSREF("R\[-2\]C\[-2\]:R\[2\]C\[2\]", C3), TRUE), which equals
"$A$1:$E$5"

In Microsoft Excel for Windows ABSREF("R\[-2\]C\[-2\]",
\[FINANCE.XLS\]Sheet1\!C3) equals \[FINANCE.XLS\]Sheet1\!$A$1.

In Microsoft Excel for the Macintosh ABSREF("R\[-2\]C\[-2\]",
\[FINANCE\]Sheet1\!C3) equals \[FINANCE\]Sheet1\!$A$1

**Related Function**

RELREF&nbsp;&nbsp;&nbsp;Returns a relative reference

Return to [top](#A)

# ACTIVATE

Switches to a window if more than one window is open, or a pane of a
window if the window is split and its panes are not frozen. Switching to
a pane is useful with functions such as VSCROLL, HSCROLL, and GOTO,
which operate only on the active pane.

**Syntax**

**ACTIVATE**(window\_text, pane\_num)

**ACTIVATE**?(window\_text, pane\_num)

Window\_text&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the name of a
window to switch to: for example, "Book1" or "Book1:2".

  - > If a workbook is displayed in more than one window and
    > window\_text does not specify which window to switch to, the first
    > window containing that workbook is switched to.

  - > If window\_text is omitted, the active window is not changed.

> &nbsp;

Pane\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 specifying
which pane to switch to. If pane\_num is omitted and the window has more
than one pane, the active pane is not changed.

|               |                                                                                                                                                                                                                          |
| ------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Pane\_num** | **Activates**                                                                                                                                                                                                            |
| 1             | Upper-left pane of the active sheet. If the window is not split, this is the only pane. If the window is split only horizontally, this is the upper pane. If the window is split only vertically, this is the left pane. |
| 2             | Upper-right pane of the active sheet. If the window is split only vertically, this is the right pane. If the window is split only horizontally, an error occurs.                                                         |
| 3             | Lower-left pane of the active sheet. If the window is split only horizontally, this is the lower pane. If the window is split only vertically, an error occurs.                                                          |
| 4             | Lower-right pane of the active sheet. If the window is split into only two panes either vertically or horizontally, an error occurs.                                                                                     |

**Related Functions**

ACTIVATE.NEXT&nbsp;&nbsp;&nbsp;Switches to the next window, or switches
to the next sheet in a workbook

ACTIVATE.PREV&nbsp;&nbsp;&nbsp;Switches to the previous window, or
switches to the previous sheet in a workbook

DOCUMENTS&nbsp;&nbsp;&nbsp;Returns the names of the specified open
workbooks

FREEZE.PANES&nbsp;&nbsp;&nbsp;Freezes the panes of a window so that they
do not scroll

ON.WINDOW&nbsp;&nbsp;&nbsp;Runs a macro when you switch to a window

SPLIT&nbsp;&nbsp;&nbsp;Splits a window

WINDOWS&nbsp;&nbsp;&nbsp;Returns the names of all open windows

WORKBOOK.SELECT&nbsp;&nbsp;&nbsp;Select a sheet in a workbook

Return to [top](#A)

# ACTIVATE.NEXT, ACTIVATE.PREV

Switches to the next or previous window, respectively, or switches to
the next or previous sheet in a workbook.

**Syntax**

**ACTIVATE.NEXT**(workbook\_text)

**ACTIVATE.PREV**(workbook\_text)

Workbook\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of the workbook for
which you want to activate a window.

  - > If workbook\_text is specified, ACTIVATE.NEXT and ACTIVATE.PREV
    > are equivalent to pressing CTRL+PAGE DOWN and CTRL+PAGE UP (in
    > Microsoft Excel for Windows) or COMMAND+PAGE DOWN and COMMAND+PAGE
    > UP (in Microsoft Excel for the Macintosh). These functions switch
    > to the next and previous sheets, respectively.

  - > If workbook\_text is omitted, these functions are equivalent to
    > pressing CTRL+TAB or CTRL+SHIFT+TAB (in Microsoft Excel for
    > Windows) or COMMAND+M or COMMAND+SHIFT+M (in Microsoft Excel for
    > the Macintosh). These functions switch to the next and previous
    > windows, respectively.

> &nbsp;

**Related Functions**

ACTIVATE&nbsp;&nbsp;&nbsp;Switches to a window

ON.WINDOW&nbsp;&nbsp;&nbsp;Runs a macro when you switch to a window

WORKBOOK.NEXT&nbsp;&nbsp;&nbsp;Switches to the next sheet in a workbook

WORKBOOK.PREV&nbsp;&nbsp;&nbsp;Switches to the previous sheet in a
workbook

WORKBOOK.SELECT&nbsp;&nbsp;&nbsp;Select a sheet in a workbook

Return to [top](#A)

# ACTIVE.CELL

Returns the reference of the active cell in the selection as an external
reference.

**Syntax**

**ACTIVE.CELL**( )

**Remarks**

  - > If an object is selected, ACTIVE.CELL returns the \#N/A error
    > value.

  - > If a chart sheet is active, ACTIVE.CELL returns a zero value.

  - > If you use ACTIVE.CELL in a function or operation, you will
    > usually get the value contained in the active cell instead of its
    > reference, because the reference is automatically converted to the
    > contents of the reference. See the third example following.

  - > If you use ACTIVE.CELL in a function that requires a reference
    > argument, then Microsoft Excel does not convert the reference to a
    > value.

  - > If you want to work with the actual reference, use the REFTEXT XLM
    > function to convert the active-cell reference to text, which you
    > can then store or manipulate (or convert back to a reference with
    > TEXTREF). See the second example following.

**Tip&nbsp;&nbsp;&nbsp;**Use the following macro formula to verify that
the current selection is a cell or range of cells:

\=ISREF(ACTIVE.CELL( ))

**Examples**

The following macro formula assigns the name Sales to the active cell:

SET.NAME("Sales", ACTIVE.CELL())

In this example, note that "Sales" refers to a cell on the active
worksheet, but the name itself exists only in the macro sheet's list of
names. In other words, the preceding formula does not define a name on
the worksheet or in the workbook's global list of names.

The following macro formula puts the reference of the active cell into
the cell named Temp:

FORMULA("="\&REFTEXT(ACTIVE.CELL()), Temp)

The following macro formula checks the contents of the active cell. If
the cell contains only the letter "c" or "s", the macro branches to an
area named FinishRefresh:

IF(OR(ACTIVE.CELL()="c", ACTIVE.CELL()="s"), GOTO(FinishRefresh))

In Microsoft Excel for Windows, if the sheet in the active window is
named SALES and A1 is the active cell, then:

ACTIVE.CELL() equals SALES\!$A$1

In Microsoft Excel for the Macintosh, if the sheet in the active window
is named SALES 1 and A1 is the active cell, then:

ACTIVE.CELL() equals 'SALES&nbsp;1'\!$A$1

**Related Function**

SELECT&nbsp;&nbsp;&nbsp;Selects a cell, worksheet object, or chart item

Return to [top](#A)

# ACTIVE.CELL.FONT

Equivalent to formatting individual characters in a cell.

**Syntax**

**ACTIVE.CELL.FONT**(font, font\_style, size, strikethrough,
superscript, subscript, outline, shadow, underline, color, normal,
background, start\_char, char\_count)

The arguments for this function are the same as those for
FONT.PROPERTIES.

**Related Function**

FONT.PROPERTIES&nbsp;&nbsp;&nbsp;Applies a font and other attributes to
the selection

Return to [top](#A)

# ADD.ARROW

Equivalent to clicking the Arrow button on the Drawing toolbar. Adds an
arrow to the active chart. If a chart is not the active window, displays
an error value.

**Syntax**

**ADD.ARROW**( )

**Remarks**

After you create an arrow with ADD.ARROW, the arrow remains selected, so
you can use the arrow form of the PATTERNS function to format the arrow
and the FORMAT.MOVE and FORMAT.SIZE functions to change the position and
size of the arrow.

**Related Functions**

CREATE.OBJECT&nbsp;&nbsp;&nbsp;Creates an object

DELETE.ARROW&nbsp;&nbsp;&nbsp;Deletes the selected arrow

FORMAT.MOVE&nbsp;&nbsp;&nbsp;Moves the selected object

FORMAT.SIZE&nbsp;&nbsp;&nbsp;Changes the size of the selected object

PATTERNS&nbsp;&nbsp;&nbsp;Changes the appearance of the selected object

Return to [top](#A)

# ADD.BAR

Creates a new menu bar and returns the bar ID number. Use the bar ID
number to identify the menu in functions that display and add menus and
commands to the menu bar. You can also use ADD.BAR to restore a built-in
menu bar with its original menus and commands.

**Syntax**

**ADD.BAR**(bar\_num)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number of a built-in menu bar
that you want to restore. Use ADD.BAR(bar\_num) to restore an unaltered
version of a built-in menu bar after you have made changes to the menu
bar's menus and commands. See ADD.COMMAND for a list of ID numbers for
built-in menu bars.

**Important&nbsp;&nbsp;&nbsp;**Restoring a built-in menu bar will remove
menus and commands added by other macros. Use ADD.COMMAND and ADD.MENU
to restore individual commands and menus.

**Remarks**

  - > ADD.BAR just creates a new menu bar; it does not display it. Use
    > SHOW.BAR to display a menu bar. The argument to the SHOW.BAR
    > function should be the number returned by ADD.BAR or a reference
    > to the cell containing ADD.BAR.

  - > You can define up to 15 custom menu bars at one time. If you carry
    > out an ADD.BAR function when more than 15 custom menu bars are
    > already defined, Microsoft Excel returns the \#VALUE\! error
    > value.

**Example**

The following formula creates a new menu bar and returns a bar ID
number:

ADD.BAR()

**Related Functions**

ADD.COMMAND&nbsp;&nbsp;&nbsp;Adds a command to a menu

ADD.MENU&nbsp;&nbsp;&nbsp;Adds a menu to a menu bar

DELETE.BAR&nbsp;&nbsp;&nbsp;Deletes a menu bar

SHOW.BAR&nbsp;&nbsp;&nbsp;Displays a menu bar

Return to [top](#A)

# ADD.CHART.AUTOFORMAT

Adds the format of the currently active chart in the current window to
the list of custom formats in the Custom Types tab in the Chart Type
dialog box.

**Syntax**

**ADD.CHART.AUTOFORMAT**(**name\_text**, desc\_text)

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name you want to appear in the
list of custom formats.

Desc\_text&nbsp;&nbsp;&nbsp;&nbsp;is the description you want to appear
when the custom format is selected.

**Related Function**

DELETE.CHART.AUTOFORMAT&nbsp;&nbsp;&nbsp;Deletes a custom template

Return to [top](#A)

# ADD.COMMAND

Adds a command to a menu. ADD.COMMAND returns the position number on the
menu of the added command. Use ADD.COMMAND to add one or more custom
menu commands to a menu on a built-in or custom menu bar. You can also
use ADD.COMMAND to restore a deleted built-in command to its original
menu.

**Syntax**

**ADD.COMMAND**(**bar\_num, menu, command\_ref**, position1, position2)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number corresponding to a menu
bar or a type of shortcut menu to which you want to add a command.

  - > Bar\_num can be the ID number of a built-in or custom menu bar.
    > The ID number of a custom menu bar is the number returned by the
    > ADD.BAR function.

  - > Bar\_num can also refer to a type of shortcut menu; use menu to
    > identify the specific shortcut menu.

> &nbsp;

The ID numbers of the built-in menu bars and the types of shortcut menus
are listed in the following tables. Short menus are abbreviated versions
of the normal Microsoft Excel menus. To turn on short menus, use the
SHORT.MENUS function.

|              |                                                                          |
| ------------ | ------------------------------------------------------------------------ |
| **Bar\_num** | **Built-in menu bar**                                                    |
| 1            | Worksheet and macro sheet (Microsoft Excel 4.0 or later)                 |
| 2            | Chart (Microsoft Excel 4.0 or later)                                     |
| 3            | Null (the menu displayed when no workbooks are open)                     |
| 4            | Info                                                                     |
| 5            | Worksheet and macro sheet (short menus, Microsoft Excel 3.0 and earlier) |
| 6            | Chart (short menus, Microsoft Excel 3.0 and earlier)                     |
| 7            | Cell, toolbar, and workbook (shortcut menus)                             |
| 8            | Object (shortcut menus)                                                  |
| 9            | Chart (Microsoft Excel 4.0 or later shortcut menus)                      |
| 10           | Worksheet and macro sheet                                                |
| 11           | Chart                                                                    |
| 12           | Visual Basic                                                             |

Menu&nbsp;&nbsp;&nbsp;&nbsp;is the menu to which you want the new
command added.

  - > Menu can be either the name of a menu as text or the number of a
    > menu.

  - > If bar\_num is 1 through 6, menus are numbered starting with 1
    > from the left of the menu bar.

  - > If bar\_num is 7, 8, or 9, menu refers to a built-in shortcut
    > menu. The combination of bar\_num and menu determines which
    > shortcut menu to modify, as shown in the following table.

> &nbsp;

|              |          |                                                                    |
| ------------ | -------- | ------------------------------------------------------------------ |
| **Bar\_num** | **Menu** | **Shortcut menu**                                                  |
| 7            | 1        | Toolbars                                                           |
| 7            | 2        | Toolbar buttons                                                    |
| 7            | 3        | Workbook paging icons in Microsoft Excel 4.0                       |
| 7            | 4        | Cells (worksheet)                                                  |
| 7            | 5        | Column selections                                                  |
| 7            | 6        | Row selections                                                     |
| 7            | 7        | Workbook tabs                                                      |
| 7            | 8        | Cells (macro sheet)                                                |
| 7            | 9        | Workbook title bar                                                 |
| 7            | 10       | Desktop (Microsoft Excel for Windows only)                         |
| 7            | 11       | Module                                                             |
| 7            | 12       | Watch pane                                                         |
| 7            | 13       | Immediate pane                                                     |
| 7            | 14       | Debug code pane                                                    |
| 8            | 1        | Drawn or imported objects on worksheets, dialog sheets, and charts |
| 8            | 2        | Buttons on sheets                                                  |
| 8            | 3        | Text boxes                                                         |
| 8            | 4        | Dialog sheet                                                       |
| 9            | 1        | Chart series                                                       |
| 9            | 2        | Chart and axis titles                                              |
| 9            | 3        | Chart plot area and walls                                          |
| 9            | 4        | Entire chart                                                       |
| 9            | 5        | Chart axes                                                         |
| 9            | 6        | Chart gridlines                                                    |
| 9            | 7        | Chart floor and arrows                                             |
| 9            | 8        | Chart legend                                                       |

**Note**&nbsp;&nbsp;&nbsp;Any commands that you add to the toolbar
buttons, watch pane, immediate pane or debug code pane shortcut menus
will be dimmed.

Command\_ref&nbsp;&nbsp;&nbsp;&nbsp;is an array or a reference to an
area on the macro sheet that describes the new command or commands.

  - > Command\_ref must be at least two columns wide. The first column
    > specifies command names; the second specifies macro names.
    > Optional columns can be specified for shortcut keys (in Microsoft
    > Excel for the Macintosh), status bar messages, and custom Help
    > topics, in that order.

  - > Command\_ref is similar to menu\_ref in ADD.MENU. For more
    > information about command\_ref, see the description of menu\_ref
    > in ADD.MENU.

  - > Command\_ref can be the name, as text, of a previously deleted
    > built-in command that you want to restore. You can also use the
    > value returned by the DELETE.COMMAND formula that deleted the
    > command.

> &nbsp;

Position1&nbsp;&nbsp;&nbsp;&nbsp;specifies the placement of the new
command.

  - > Use a hyphen (-) to represent a line separating commands on a
    > menu. If you want to place a command before the second separator
    > on a menu, use two hyphens (--), three hyphens for the third
    > separator, and so on.

  - > Position1 can be a number indicating the position of the command
    > on the menu. Commands are numbered from the top of the menu
    > starting with 1.

  - > Position1 can be the name of an existing command, as text, above
    > which you want to add the new command.

  - > If position1 is omitted, the command is added to the bottom of the
    > menu.

  - > For the toolbar shortcut menu (bar\_num 7, menu 1) and the
    > shortcut menu for workbook paging icons in Microsoft Excel version
    > 4.0 (bar\_num 7, menu 3), you cannot add commands to the middle of
    > the toolbar name list or the middle of the workbook contents list.

> &nbsp;

Position2&nbsp;&nbsp;&nbsp;&nbsp;specifies the placement of the new
command on a submenu.

  - > Position2 can be a number indicating the position of the command
    > on the submenu. Commands are numbered from the top of the menu
    > starting with 1.

  - > Position2 can be the name of an existing command, as text, above
    > which you want to add the new command.

  - > If position2 is omitted, the command is added to the main menu,
    > not the submenu.

  - > To add a command to the bottom of a submenu, use 0 for position2.

**Tip**&nbsp;&nbsp;&nbsp;In general, use menu and command names rather
than numbers for arguments. The numbers assigned to menus and commands
change as you add and delete menus and commands. Using names ensures
that your menu and command macro functions always refer to the correct
items.

**Example**

The following macro formula adds the command described in cells G16:J16
to the bottom of the worksheet cells shortcut menu:

ADD.COMMAND(7, 4, G16:J16)

**Related Functions**

ADD.BAR&nbsp;&nbsp;&nbsp;Adds a menu bar

ADD.MENU&nbsp;&nbsp;&nbsp;Adds a menu to a menu bar

ADD.TOOL&nbsp;&nbsp;&nbsp;Adds one or more buttons to a toolbar

ADD.TOOLBAR&nbsp;&nbsp;&nbsp;Creates a toolbar with the specified tools

DELETE.COMMAND&nbsp;&nbsp;&nbsp;Deletes a command from a menu

ENABLE.COMMAND&nbsp;&nbsp;&nbsp;Enables or disables a menu or custom
command

GET.TOOLBAR&nbsp;&nbsp;&nbsp;Retrieves information about a toolbar

RENAME.COMMAND&nbsp;&nbsp;&nbsp;Changes the name of a command or menu

Return to [top](#A)

# ADDIN.MANAGER

Equivalent to clicking the Add-Ins command on the Tools menu. Adds or
removes an installed add-in from the working set. The add-in file must
already be installed.

**Syntax**

**ADDIN.MANAGER**(operation\_num, addinname\_text, copy\_logical)

**ADDIN.MANAGER**?(operation\_num, addinname\_text, copy\_logical)

Operation\_num&nbsp;&nbsp;&nbsp;&nbsp;determines the operation that the
add-in manager will perform.

|                    |                                                                                                                                                                       |
| ------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Operation\_num** | **Operation**                                                                                                                                                         |
| 1                  | Adds an add-in to the working set, using the descriptive name in the Add-Ins dialog box.                                                                              |
| 2                  | Removes an add-in from the working set, using the descriptive name in the Add-Ins dialog box.                                                                         |
| 3                  | Adds a new add-in to the list of add-ins that Microsoft Excel knows about. Equivalent to clicking on the Browse button in the Add-Ins dialog box and clicking a file. |

Addinname\_text&nbsp;&nbsp;&nbsp;&nbsp;specifies the name of the add-in.
If operation\_num is 1 or 2, use the descriptive name of the add-in,
such as "SOLVER". If operation\_num is 3, this contains the filename of
the add-in.

Copy\_logical&nbsp;&nbsp;&nbsp;&nbsp;specifies whether the add-in should
be copied to the library directory. This argument is only used if
operation\_num is 3. If omitted, and the file is on removable media, the
user will be asked if they want to copy it to removable media.

Return to [top](#A)

# ADD.LIST.ITEM

Adds an item in a list box or drop-down control on a worksheet or dialog
sheet control.

**Syntax**

**ADD.LIST.ITEM**(**text**, index\_num)

Text&nbsp;&nbsp;&nbsp;&nbsp;specifies the text of the item to be added.
Instead of text, an empty string may be inserted.

Index\_num&nbsp;&nbsp;&nbsp;&nbsp;is the list index to be used for the
new item. Blank entries are created from the end of the current list to
the new item index. If index\_num is omitted the new item is appended to
the list.

**Remarks**

If the list box or drop-down box was already filled using the
LISTBOX.PROPERTIES function, then adding an item with ADD.LIST.ITEM
causes the fillrange contents to be discarded in favor of the new list.

**Related Functions**

REMOVE.LIST.ITEM&nbsp;&nbsp;&nbsp;Removes an item in a list box or
drop-down box

SELECT.LIST.ITEM&nbsp;&nbsp;&nbsp;Selects an item in a list box or in a
group box

Return to [top](#A)

# ADD.MENU

Adds a menu to a menu bar. Use ADD.MENU to add a custom menu to a
built-in or custom menu bar. You can also use ADD.MENU to restore
built-in menus you have deleted with DELETE.MENU. ADD.MENU returns the
position number in the menu bar of the new menu.

**Syntax**

**ADD.MENU**(**bar\_num, menu\_ref**, position1, position2)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the menu bar to which you want a menu
added. Bar\_num can be the ID number of a built-in or custom menu bar.
See ADD.COMMAND for a list of ID numbers for built-in menu bars.

Menu\_ref&nbsp;&nbsp;&nbsp;&nbsp;is an array or a reference to an area
on the macro sheet that describes the new menu or the name of a deleted
built-in menu you want to restore.

  - > Menu\_ref must be made up of at least two rows and two columns of
    > cells. The upper-left cell of menu\_ref specifies the menu title,
    > which is displayed in the menu bar. In the following example, the
    > range A3:E10 is a valid menu\_ref.

![](./media/image1.png)

> The rest of the first column indicates the names of the commands. The
> corresponding rows in the second column give the names of the macros
> that run when the commands are chosen.

  - > You can also specify status-bar text and Help topics in the fourth
    > and fifth columns of menu\_ref. In Microsoft Excel for the
    > Macintosh, you can specify shortcut keys in the third column of
    > menu\_ref.

> &nbsp;

Position1&nbsp;&nbsp;&nbsp;&nbsp;specifies the placement of the new
menu. Position can be the name of a menu, as text, or the number of a
menu. Menus are numbered from left to right starting with 1. Menus are
added to the left of the position specified.

  - > Use a hyphen (-) to represent a line separating commands on a
    > menu. If you want to place a command before the second separator
    > on a menu, use two hyphens (--), three hyphens for the third
    > separator, and so on.

  - > If position1 is omitted, the menu is added to the end of the menu
    > bar.

  - > If there is already a menu at position1, that menu is shifted to
    > the right and the new menu is added in its place.

  - > If you are using ADD.MENU to restore a deleted built-in menu, you
    > can use the position argument to put it back in its original place
    > on the menu bar. For example, to restore the Data menu on the
    > worksheet and macro sheet menu bar, use position 7. If position1
    > is omitted, the menu is added to the right of the last menu
    > restored.

> &nbsp;

Position2&nbsp;&nbsp;&nbsp;&nbsp;specifies the placement of a submenu.

  - > Use a hyphen (-) to represent a line separating commands on a
    > menu. If you want to place a command before the second separator
    > on a menu, use two hyphens (--), three hyphens for the third
    > separator, and so on.

  - > Position2 can be a number indicating the position of the submenu
    > on the menu. Commands are numbered from the top of the menu
    > starting with 1 and include separators.

  - > Position2 can also be the name, as text, of an existing command
    > above which you want to add the new command.

  - > If position2 is omitted, the command is added to the main menu,
    > not the submenu.

> &nbsp;

**Example**

The following macro formula adds a new menu to the end of the worksheet
menu bar, where A10:B15 is the menu\_ref describing the menu:

ADD.MENU(1, A10:B15)

**Related Functions**

ADD.BAR&nbsp;&nbsp;&nbsp;Adds a menu bar

ADD.COMMAND&nbsp;&nbsp;&nbsp;Adds a command to a menu

DELETE.MENU&nbsp;&nbsp;&nbsp;Deletes a menu

ENABLE.COMMAND&nbsp;&nbsp;&nbsp;Enables or disables a menu or custom
command

Return to [top](#A)

# ADD.OVERLAY

Equivalent to clicking the Add Overlay command on the Chart menu in
Microsoft Excel version 4.0. Adds an overlay to a 2-D chart. If the
active chart already has an overlay, ADD.OVERLAY takes no action and
returns TRUE. In Microsoft Excel version 5.0 or later, ADD.OVERLAY works
with charts that have only one chart type.

**Syntax**

**ADD.OVERLAY**( )

**Related Functions**

ADD.ARROW&nbsp;&nbsp;&nbsp;Adds an arrow to a chart

LEGEND&nbsp;&nbsp;&nbsp;Adds a legend to a chart

Return to [top](#A)

# ADD.TOOL

Adds one or more buttons to a toolbar.

**Syntax**

**ADD.TOOL**(**bar\_id, position, tool\_ref**)

Bar\_id&nbsp;&nbsp;&nbsp;&nbsp;is either a number specifying one of the
built-in toolbars or the name of a custom toolbar.

|             |                      |
| ----------- | -------------------- |
| **Bar\_id** | **Built-in toolbar** |
| 1           | Standard             |
| 2           | Formatting           |
| 3           | Query and Pivot      |
| 4           | Chart                |
| 5           | Drawing              |
| 6           | TipWizard            |
| 7           | Forms                |
| 8           | Stop Recording       |
| 9           | Visual Basic         |
| 10          | Auditing             |
| 11          | WorkGroup            |
| 12          | Microsoft            |
| 13          | Full Screen          |

Position&nbsp;&nbsp;&nbsp;&nbsp;specifies the position of the button
within the toolbar. Position starts with 1 at the left side (if
horizontal) or at the top (if vertical).

Tool\_ref&nbsp;&nbsp;&nbsp;&nbsp;is either a number specifying a
built-in button or a reference to an area on the macro sheet that
defines a custom button or set of buttons (or an array containing this
information).

For customized buttons, the following example shows the components of a
button reference area on a macro sheet and defines custom tools. The
range A1:I5 is a valid tool\_ref. Row 1 refers to a built-in tool. Row 2
defines a gap. For this illustration, values are displayed instead of
formulas so that text can wrap in cells.

![](./media/image2.png)

  - > Tool\_id is a number associated with the tool. A zero specifies a
    > gap on the toolbar. To specify a custom button, use a name, or a
    > number between 201 and 231.

  - > Macro is the name of, or a quoted R1C1-style reference to, the
    > macro you want to run when the button is clicked.

  - > Down is a logical value specifying the default image of the tool.
    > If down is TRUE, the button appears depressed into the screen; if
    > FALSE or omitted, it appears normal (up).

  - > Enabled is a logical value specifying whether the button can be
    > used. If enabled is TRUE, the button is enabled; if FALSE, it is
    > disabled.

  - > Face specifies a face associated with the tool. Face must be a
    > reference to a picture-type object, for example "Picture 1". If
    > face is omitted, Microsoft Excel uses the default face for the
    > tool.

  - > Status\_text is the text, if any, that you want displayed in the
    > status bar when the button is selected.

  - > Balloon\_text is the balloon help text, if any, associated with
    > the tool. Balloon\_text is available only in Microsoft Excel for
    > the Macintosh using system software version 7.0 or later.

  - > Help\_topics is a reference to a topic in a Help file, in the form
    > "filename\!topic\_number". Help\_topics must be text. If
    > help\_topics is omitted, HELP displays the Contents topic for
    > Microsoft Excel Help.

  - > Tip\_text is the text, if any, that you want displayed as a
    > ToolTip when the mouse pointer moves over a tool button.

> &nbsp;

To indicate that a particular component of tool\_ref is not used, clear
the contents of the corresponding cell.

**Remarks**

  - > If you do not want to reserve a section of your macro sheet to
    > define the buttons, you can use an array as the tool\_ref argument
    > as shown in the following syntax:

> **ADD.TOOL**(**bar\_id**, position, {**tool\_id1**, macro1, down1,
> enabled1, face1,  
> status\_text1, balloon\_text1, help\_topics1;tool\_id2, macro2, down2,
> enabled2,  
> face2, status\_text2, balloon\_text2, help\_topics2;...})

  - > Picture objects can be created with the camera button or pasted in
    > from another application. In Microsoft Excel for Windows, the
    > graphic object must be either a Windows bitmap or picture object.
    > In Microsoft Excel for the Macintosh, the object must be a picture
    > object.

> &nbsp;

**Examples**

The following macro formula adds a button to Toolbar5. The cell range
B6:I6 contains tool\_ref.

ADD.TOOL("Toolbar5", 6, B6:I6)

The following macro formula adds the New Macro Sheet button to the fifth
position on the Standard toolbar:

ADD.TOOL(1, 5, 6)

**Related Functions**

ADD.COMMAND&nbsp;&nbsp;&nbsp;Adds a command to a menu

ADD.TOOLBAR&nbsp;&nbsp;&nbsp;Creates a toolbar with the specified tools

DELETE.TOOL&nbsp;&nbsp;&nbsp;Deletes a button from a toolbar

DELETE.TOOLBAR&nbsp;&nbsp;&nbsp;Deletes custom toolbars

RESET.TOOLBAR&nbsp;&nbsp;&nbsp;Resets a built-in toolbar to its default
initial setting

Return to [top](#A)

# ADD.TOOLBAR

Creates a new toolbar with the specified buttons.

**Syntax**

**ADD.TOOLBAR**(**bar\_name**, tool\_ref)

Bar\_name&nbsp;&nbsp;&nbsp;&nbsp;is a text string identifying the
toolbar you want to create.

Tool\_ref&nbsp;&nbsp;&nbsp;&nbsp;is either a number specifying a
built-in button or a reference to an area on the macro sheet that
defines a custom button or set of buttons (or an array containing this
information).

For a complete description of tool\_ref, see ADD.TOOL.

**Remarks**

If you create a toolbar without buttons, use ADD.TOOL to add them. Use
SHOW.TOOLBAR to display the toolbar.

**Example**

The following macro formula creates Toolbar9 with one button in it. The
cell range B7:I7 contains tool\_ref.

ADD.TOOLBAR("Toolbar9", B7:I7)

**Related Functions**

ADD.TOOL&nbsp;&nbsp;&nbsp;Adds a button to a toolbar

DELETE.TOOL&nbsp;&nbsp;&nbsp;Deletes a button from a toolbar

DELETE.TOOLBAR&nbsp;&nbsp;&nbsp;Deletes custom toolbars

RESET.TOOLBAR&nbsp;&nbsp;&nbsp;Resets a built-in toolbar to its default
initial setting

SHOW.TOOLBAR&nbsp;&nbsp;&nbsp;Hides or displays a toolbar

Return to [top](#A)

# ALERT

Displays a dialog box and message and waits for you to click a button.
Use ALERT instead of MESSAGE if you want to interrupt the flow of a
macro and force the user to make a choice or to notice an important
message.

**Syntax**

**ALERT**(message\_text, type\_num, help\_ref)

Message\_text&nbsp;&nbsp;&nbsp;&nbsp;is the message displayed in the
dialog box.

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 specifying
which type of dialog box to display. If you omit type\_num, it is
assumed to be 2.

  - > If type\_num is 1, ALERT displays a dialog box containing the OK
    > and Cancel buttons. Click a button to continue or cancel an
    > action. ALERT returns TRUE if you click the OK button and FALSE if
    > you click the Cancel button. See the last example below.

  - > If type\_num is 2 or 3, ALERT displays a dialog box containing an
    > OK button. Click the button to continue, and ALERT returns TRUE.
    > The only difference between specifying 2 or 3 is that ALERT
    > displays a different icon on the left side of the dialog box as
    > shown in the examples below. So, for example, you could use 2 for
    > notes or to present general information, and 3 for errors or
    > warnings.

> &nbsp;

Help\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to a custom online Help
topic, in the form "filename\! topic\_number".

  - > If help\_ref is present, a Help button appears in the lower-right
    > corner of the alert message. Clicking the Help button starts Help
    > and displays the specified topic.

  - > If help\_ref is omitted, no Help button appears.

  - > Help\_ref must be given in text form.

> &nbsp;

**Note&nbsp;&nbsp;&nbsp;**In Microsoft Excel for the Macintosh, the
ALERT dialog box is not a movable window.

**Examples**

The following dialog boxes show the results of using ALERT with
type\_num 1, 2, and 3. The first and fourth examples include a Help
button.

In Microsoft Excel for Windows, the following macro formulas display
these three dialog boxes.

ALERT("Are you sure you want to delete this item?", 1,
"CUSTHELP.HLP\!101")

![](./media/image3.png)

ALERT("The number should be between 1 and 100", 2)

![](./media/image4.png)

ALERT("Your debits and credits are not equal; do not end this
transaction.", 3)

![](./media/image5.png)

In Microsoft Excel for the Macintosh, the following macro formulas
display these three dialog boxes.

ALERT("Are you sure you want to delete this item?", 1, "'Custom
Help'\!101")

![](./media/image6.png)

ALERT("The number should be between 1 and 100", 2)

![](./media/image7.png)

ALERT("Your debits and credits are not equal; do not end this
transaction.", 3)

![](./media/image8.png)

A common use of the ALERT function is to give the user a choice of two
actions. The following macro formula in an Auto\_Open macro asks which
reference style to use when the workbook is opened.

A1.R1C1(ALERT("Click OK for A1 style; Cancel for R1C1", 1))

**Related Functions**

INPUT&nbsp;&nbsp;&nbsp;Displays a dialog box for user input

MESSAGE&nbsp;&nbsp;&nbsp;Displays a message in the status bar

Return to [top](#A)

# ALIGNMENT

Equivalent to clicking the Alignment tab in the Format Cells dialog box,
which is displayed when you click the Cells command on the Format menu.
Aligns the contents of the selected cells.

**Syntax**

**ALIGNMENT**(horiz\_align, wrap, vert\_align, orientation, add\_indent,
shrink\_to\_fit, merge\_cells)

**ALIGNMENT**?(horiz\_align, wrap, vert\_align, orientation,
add\_indent, shrink\_to\_fit, merge\_cells)

Horiz\_align&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 7 specifying
the type of horizontal alignment, as shown in the following table. If
horiz\_align is omitted, horizontal alignment does not change.

|                  |                          |
| ---------------- | ------------------------ |
| **Horiz\_align** | **Horizontal alignment** |
| 1                | General                  |
| 2                | Left                     |
| 3                | Center                   |
| 4                | Right                    |
| 5                | Fill                     |
| 6                | Justify                  |
| 7                | Center across selection  |

Wrap&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the Wrap
Text check box in the Alignment tab. If wrap is TRUE, Microsoft Excel
selects the check box and wraps text in cells; if FALSE, Microsoft Excel
clears the check box and does not wrap text. If wrap is omitted,
wrapping does not change.

Vert\_align&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 specifying
the vertical alignment of the text. If vert\_align is omitted, vertical
alignment does not change.

|                 |                        |
| --------------- | ---------------------- |
| **Vert\_align** | **Vertical alignment** |
| 1               | Top                    |
| 2               | Center                 |
| 3               | Bottom                 |
| 4               | Justify                |

Orientation&nbsp;&nbsp;&nbsp;&nbsp;is a number from 0 to 4 specifying
the orientation of the text. If orientation is omitted, text orientation
does not change.

|                 |                                               |
| --------------- | --------------------------------------------- |
| **Orientation** | **Text orientation**                          |
| 0               | Horizontal                                    |
| 1               | Vertical                                      |
| 2               | Upward                                        |
| 3               | Downward                                      |
| 4               | Automatic (applies to only chart tick labels) |

Add\_indent&nbsp;&nbsp;&nbsp;&nbsp; This argument is for only Far East
versions of Microsoft Excel.

Shrink\_to\_fit&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding
to the Shrink To Fit check box in the Alignment tab.

Merge\_cells&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Merge Cells check box in the Alignment tab. If merge\_cells is TRUE,
Microsoft Excel selects the check box and merges the selected cells; the
merged cell contains the value of the left-most cell that was merged. If
FALSE, Microsoft Excel clears the check box and unmerges the selected
cells; the left-most cell takes the formula or value of the cell that
was unmerged. If merge\_cells is omitted, cell mergers do not change.

**Related Function**

FORMAT.TEXT&nbsp;&nbsp;&nbsp;Formats a worksheet text box or a chart
text item

Return to [top](#A)

# ANOVA1

Performs single-factor analysis of variance, which tests the hypothesis
that means from several samples are equal.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**ANOVA1**(**inprng**, outrng, grouped, labels, alpha)

**ANOVA1**?(inprng, outrng, grouped, labels, alpha)

Inprng&nbsp;&nbsp;&nbsp;&nbsp;is the input range.

Outrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell) in
the output table or the name, as text, of a new sheet to contain the
output table. If FALSE, blank, or omitted, places the output table in a
new workbook.

Grouped&nbsp;&nbsp;&nbsp;&nbsp;is a text character that indicates
whether the data in the input range is organized by row or column.

  - > If grouped is "C" or omitted, then the data is organized by
    > column.

  - > If grouped is "R", then the data is organized by row.

Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that describes where
the labels are located in the input range, as shown in the following
table:

|                  |             |                                                   |
| ---------------- | ----------- | ------------------------------------------------- |
| **Labels**       | **Grouped** | **Labels are in**                                 |
| TRUE             | "C"         | First row of the input range.                     |
| TRUE             | "R"         | First column of the input range.                  |
| FALSE or omitted | (ignored)   | No labels. All cells in the input range are data. |

Alpha&nbsp;&nbsp;&nbsp;&nbsp;is the significance level at which to
evaluate critical values for the F statistic. If omitted, alpha is 0.05.

**Related Functions**

ANOVA2&nbsp;&nbsp;&nbsp;Performs two-factor analysis of variance with
replication

ANOVA3&nbsp;&nbsp;&nbsp;Performs two-factor analysis of variance without
replication

Return to [top](#A)

# ANOVA2

Performs two-factor analysis of variance with replication.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**ANOVA2**(**inprng**, outrng, **sample\_rows**, alpha)

**ANOVA2**?(inprng, outrng, sample\_rows, alpha)

Inprng&nbsp;&nbsp;&nbsp;&nbsp;is the input range. The input range should
contain labels in the first row and column.

Outrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell) in
the output table or the name, as text, of a new sheet to contain the
output table. If FALSE, blank, or omitted, places the output table in a
new workbook.

Sample\_rows&nbsp;&nbsp;&nbsp;&nbsp;is the number of rows in each
sample.

Alpha&nbsp;&nbsp;&nbsp;&nbsp;is the significance level at which to
evaluate critical values for the F statistic. If omitted, alpha is 0.05.

**Related Functions**

ANOVA1&nbsp;&nbsp;&nbsp;Performs single-factor analysis of variance

ANOVA3&nbsp;&nbsp;&nbsp;Performs two-factor analysis of variance without
replication

Return to [top](#A)

# ANOVA3

Performs two-factor analysis of variance without replication.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**ANOVA3**(**inprng**, outrng, labels, alpha)

**ANOVA3**?(inprng, outrng, labels, alpha)

Inprng&nbsp;&nbsp;&nbsp;&nbsp;is the input range.

Outrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell) in
the output table or the name, as text, of a new sheet to contain the
output table. If FALSE, blank, or omitted, places the output table in a
new workbook.

Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value.

  - > If labels is TRUE, then the first row and column of the input
    > range contain labels.

  - > If labels is FALSE or omitted, all cells in inprng are considered
    > data. Microsoft Excel will then generate the appropriate data
    > labels for the output table.

> &nbsp;

Alpha&nbsp;&nbsp;&nbsp;&nbsp;is the significance level at which to
evaluate critical values for the F statistic. If omitted, alpha is 0.05.

**Related Functions**

ANOVA1&nbsp;&nbsp;&nbsp;Performs single-factor analysis of variance

ANOVA2&nbsp;&nbsp;&nbsp;Performs two-factor analysis of variance with
replication

Return to [top](#A)

# APP.ACTIVATE

Switches to an application. Use APP.ACTIVATE to switch to another
application that is already running or that you have started by using
EXEC.

**Syntax**

**APP.ACTIVATE**(title\_text, wait\_logical)

**Important&nbsp;&nbsp;&nbsp;**Microsoft Excel for the Macintosh
requires system software version 7.0 or later for this function.

Title\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of an application as
displayed in its title bar.

  - > If title\_text is omitted, APP.ACTIVATE switches to Microsoft
    > Excel.

  - > If title\_text is not a currently running application,
    > APP.ACTIVATE returns the \#VALUE\! error value and interrupts the
    > macro.

  - > Title\_text is not necessarily the name of the application file.
    > Use the text that appears in the title bar of the application,
    > which might include the name of the open workbook and path
    > information.

  - > In Microsoft Excel for the Macintosh, title\_text can also refer
    > to the Process Serial Number (PSN) that is returned by an EXEC
    > function.

> &nbsp;

Wait\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value determining when
to switch to the application specified by title\_text.

  - > If wait\_logical is TRUE, Microsoft Excel waits to be switched to
    > before switching to the application specified by title\_text.

  - > If wait\_logical is FALSE or omitted, Microsoft Excel immediately
    > switches to the application specified by title\_text.

> &nbsp;

**Remarks**

If you are running an application using Microsoft Excel macros, and you
want to switch to a third application without switching to Microsoft
Excel first, use FALSE as the wait\_logical argument. With FALSE, you
can use the application title\_text without having to switch to
Microsoft Excel first.

**Examples**

The following macro formula switches to Microsoft Word, which is
currently displaying the workbook MONTHRPT.DOC in full screen mode:

APP.ACTIVATE("MICROSOFT WORD - MONTHRPT.DOC")

In Microsoft Excel for the Macintosh, the following macro formula
switches to Microsoft Word:

APP.ACTIVATE("MICROSOFT WORD")

**Tip**&nbsp;&nbsp;&nbsp;Use an IF statement with APP.ACTIVATE to run an
EXEC function if the application you want to switch to is not yet
running.

**Related Functions**

The first five functions following are only for Microsoft Excel for
Windows.

APP.MAXIMIZE&nbsp;&nbsp;&nbsp;Maximizes the Microsoft Excel application
window

APP.MINIMIZE&nbsp;&nbsp;&nbsp;Minimizes the Microsoft Excel application
window

APP.MOVE&nbsp;&nbsp;&nbsp;Moves the Microsoft Excel application window

APP.RESTORE&nbsp;&nbsp;&nbsp;Restores the Microsoft Excel application
window

APP.SIZE&nbsp;&nbsp;&nbsp;Changes the size of the Microsoft Excel
application window

EXEC&nbsp;&nbsp;&nbsp;Starts another application

Return to [top](#A)

# APP.ACTIVATE.MICROSOFT

Activates a Microsoft application. If the application is not already
activated, this function will load the application into memory.

**Syntax**

**APP.ACTIVATE.MICROSOFT**(**app\_id**)

App\_id&nbsp;&nbsp;&nbsp;&nbsp;is the ID number associated with the
Microsoft Application.

|             |                                               |
| ----------- | --------------------------------------------- |
| **App\_id** | **Application**                               |
| 1           | Microsoft Word                                |
| 2           | Microsoft PowerPoint                          |
| 3           | Microsoft Mail                                |
| 4           | Microsoft Access (for Microsoft Windows only) |
| 5           | Microsoft Fox Pro                             |
| 6           | Microsoft Project                             |
| 7           | Microsoft Schedule +                          |

**Remarks**

Returns TRUE if the application is activated successfully. Returns FALSE
if the application is not activated successfully.

**Related Function**

APP.ACTIVATE&nbsp;&nbsp;&nbsp;Switches to an application.

Return to [top](#A)

# APPLY.NAMES

Equivalent to clicking the Apply command on the Name submenu on the
Insert menu. Replaces definitions with their respective names. If no
names are defined in the current selection, APPLY.NAMES returns the
\#VALUE\! error value. Use APPLY.NAMES to replace references and values
in formulas with names.

**Syntax**

**APPLY.NAMES**(**name\_array**, ignore, use\_rowcol, omit\_col,
omit\_row, order\_num, append\_last)

**APPLY.NAMES**?(name\_array, ignore, use\_rowcol, omit\_col, omit\_row,
order\_num, append\_last)

Name\_array&nbsp;&nbsp;&nbsp;&nbsp;is the name or names to apply as text
elements in an array.

  - > To give more than one name as the argument, you must use an array.
    > For example:

  - > APPLY.NAMES({"DataRange", "CriteriaRange"})

  - > If the names indicated by the argument name\_array have already
    > replaced all of the appropriate references or values, the
    > \#VALUE\! error value is returned.

> &nbsp;

The next four arguments correspond to check boxes and options in the
Apply Names dialog box. Arguments that correspond to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Ignore&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Ignore
Relative/Absolute check box.

Use\_rowcol&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Use Row And Column
Names check box. If use\_rowcol is FALSE, the next three arguments are
ignored.

Omit\_col&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Omit Column Name If
Same Column check box.

Omit\_row&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Omit Row Name If
Same Row check box.

Order\_num&nbsp;&nbsp;&nbsp;&nbsp;determines which range name is listed
first when a cell reference is replaced by a row-oriented and a
column-oriented range name, as shown in the following table.

|                |                          |
| -------------- | ------------------------ |
| **Order\_num** | **Order of range names** |
| 1              | Row Column               |
| 2              | Column Row               |

Append\_last&nbsp;&nbsp;&nbsp;&nbsp;determines whether the names most
recently defined are also replaced.

  - > If append\_last is TRUE, Microsoft Excel replaces the definitions
    > of the names in name\_array and also replaces the definitions of
    > the last names defined.

  - > If append\_last is FALSE or omitted, Microsoft Excel replaces the
    > definitions of the names in name\_array only.

> &nbsp;

**Related Functions**

CREATE.NAMES&nbsp;&nbsp;&nbsp;Creates names automatically from text
labels on a sheet

DEFINE.NAME&nbsp;&nbsp;&nbsp;Defines a name in the active workbook

LIST.NAMES&nbsp;&nbsp;&nbsp;Lists names and their associated information

Return to [top](#A)

# APPLY.STYLE

Equivalent to clicking the Style command on the Format menu, selecting a
style, and clicking the OK button. Applies a previously defined style to
the current selection.

**Syntax**

**APPLY.STYLE**(style\_text)

**APPLY.STYLE**?(style\_text)

Style\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name, as text, of a previously
defined style. If style\_text is not defined, APPLY.STYLE returns the
\#VALUE\! error value and interrupts the macro. If style\_text is
omitted, the Normal style is applied to the selection.

**Related Functions**

DEFINE.STYLE&nbsp;&nbsp;&nbsp;Defines a cell style

DELETE.STYLE&nbsp;&nbsp;&nbsp;Deletes a cell style

MERGE.STYLES&nbsp;&nbsp;&nbsp;Imports styles from another workbook into
the active workbook

Return to [top](#A)

# APP.MAXIMIZE

Equivalent to clicking the Maximize command on the Control menu for the
application window. Maximizes the Microsoft Excel window.

**Syntax**

**APP.MAXIMIZE**( )

**Note**&nbsp;&nbsp;&nbsp;This function is only for Microsoft Excel for
Windows. You can use this function in macros created with Microsoft
Excel for the Macintosh, but it will return the \#N/A error value.

**Related Functions**

APP.ACTIVATE&nbsp;&nbsp;&nbsp;Switches to an application

APP.MINIMIZE&nbsp;&nbsp;&nbsp;Minimizes the Microsoft Excel application
window

APP.MOVE&nbsp;&nbsp;&nbsp;Moves the Microsoft Excel application window

APP.RESTORE&nbsp;&nbsp;&nbsp;Restores the Microsoft Excel application
window

APP.SIZE&nbsp;&nbsp;&nbsp;Changes the size of the Microsoft Excel
application window

FULL.SCREEN&nbsp;&nbsp;&nbsp;Controls full screen display

Return to [top](#A)

# APP.MINIMIZE

Equivalent to clicking the Minimize command on the Control menu for the
application window. Minimizes the Microsoft Excel window.

**Syntax**

**APP.MINIMIZE**( )

**Note**&nbsp;&nbsp;&nbsp;This function is only for Microsoft Excel for
Windows. You can use this function in macros created with Microsoft
Excel for the Macintosh, but it will return the \#N/A error value.

**Related Functions**

APP.ACTIVATE&nbsp;&nbsp;&nbsp;Switches to an application

APP.MAXIMIZE&nbsp;&nbsp;&nbsp;Maximizes the Microsoft Excel application
window

APP.MOVE&nbsp;&nbsp;&nbsp;Moves the Microsoft Excel application window

APP.RESTORE&nbsp;&nbsp;&nbsp;Restores the Microsoft Excel application
window

APP.SIZE&nbsp;&nbsp;&nbsp;Changes the size of the Microsoft Excel
application window

Return to [top](#A)

# APP.MOVE

Equivalent to clicking the Move command on the Control menu for the
application window. Moves the Microsoft Excel window. In Microsoft Excel
for Windows, if the application window is already maximized APP.MOVE
returns the \#VALUE\! error value and interrupts the macro.

**Syntax**

**APP.MOVE**(x\_num, y\_num)

**APP.MOVE**?(x\_num, y\_num)

**Note**&nbsp;&nbsp;&nbsp;This function is only for Microsoft Excel for
Windows. You can use this function in macros created with Microsoft
Excel for the Macintosh, but it will return the \#N/A error value.

X\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the horizontal position of the
Microsoft Excel window measured in points from the left edge of your
screen to the left side of the Microsoft Excel window.

Y\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the vertical position of the
Microsoft Excel window measured in points from the top edge of your
screen to the top of the Microsoft Excel window.

**Remarks**

  - > APP.MOVE?, the dialog-box form of the function, doesn't display a
    > dialog box. Instead, it is equivalent to pressing ALT + SPACEBAR,
    > M or to dragging the title bar with the mouse. With APP.MOVE?, you
    > can move the window with the keyboard or mouse.

  - > If you specify x\_num and/or y\_num in the dialog-box form of the
    > function, the window is moved according to the specified
    > coordinates, and you are left in move mode.

**Related Functions**

APP.ACTIVATE&nbsp;&nbsp;&nbsp;Switches to an application

APP.MAXIMIZE&nbsp;&nbsp;&nbsp;Maximizes the Microsoft Excel application
window

APP.MINIMIZE&nbsp;&nbsp;&nbsp;Minimizes the Microsoft Excel application
window

APP.RESTORE&nbsp;&nbsp;&nbsp;Restores the Microsoft Excel application
window

APP.SIZE&nbsp;&nbsp;&nbsp;Changes the size of the Microsoft Excel
application window

Return to [top](#A)

# APP.RESTORE

Equivalent to clicking the Restore command on the Control menu for the
application window. Restores the Microsoft Excel window to its previous
size and location.

**Syntax**

**APP.RESTORE**( )

**Note&nbsp;**&nbsp;&nbsp;This function is only for Microsoft Excel for
Windows. You can use this function in macros created with Microsoft
Excel for the Macintosh, but it will return the \#N/A error value.

**Related Functions**

APP.ACTIVATE&nbsp;&nbsp;&nbsp;Switches to an application

APP.MAXIMIZE&nbsp;&nbsp;&nbsp;Maximizes the Microsoft Excel application
window

APP.MINIMIZE&nbsp;&nbsp;&nbsp;Minimizes the Microsoft Excel application
window

APP.MOVE&nbsp;&nbsp;&nbsp;Moves the Microsoft Excel application window

APP.SIZE&nbsp;&nbsp;&nbsp;Changes the size of the Microsoft Excel
application window

Return to [top](#A)

# APP.SIZE

Equivalent to choosing the Size command from the Control menu for the
application window. Changes the size of the Microsoft Excel window.

**Syntax**

**APP.SIZE**(x\_num, y\_num)

**APP.SIZE**?(x\_num, y\_num)

**Note&nbsp;&nbsp;&nbsp;**This function is only for Microsoft Excel for
Windows. You can use this function in macros created with Microsoft
Excel for the Macintosh, but it will return the \#N/A error value.

X\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the width of the Microsoft Excel
window in points.

Y\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the height of the Microsoft
Excel window in points.

APP.SIZE?, the dialog-box form of the function, doesn't display a dialog
box. Instead, it is equivalent to pressing ALT, SPACEBAR, S or to
dragging a window border with the mouse. Using APP.SIZE?, you can size
the window with the keyboard or mouse. If you specify x\_num and/or
y\_num in the dialog-box form of the function, the window is sized
according to the specified coordinates, and you are left in size mode.

**Related Functions**

APP.ACTIVATE&nbsp;&nbsp;&nbsp;Switches to an application

APP.MAXIMIZE&nbsp;&nbsp;&nbsp;Maximizes the Microsoft Excel application
window

APP.MINIMIZE&nbsp;&nbsp;&nbsp;Minimizes the Microsoft Excel application
window

APP.MOVE&nbsp;&nbsp;&nbsp;Moves the Microsoft Excel application window

APP.RESTORE&nbsp;&nbsp;&nbsp;Restores the Microsoft Excel application
window

Return to [top](#A)

# APP.TITLE

Changes the title of the Microsoft Excel application workspace to the
title you specify. The title appears at the top of the application
window. Use APP.TITLE to control the application title when you're using
Microsoft Excel to create a custom application. This function does not
apply to Microsoft Excel for the Macintosh.

**Syntax**

**APP.TITLE**(text)

Text&nbsp;&nbsp;&nbsp;&nbsp;is the title you want to assign to the
Microsoft Excel application workspace. If text is omitted, it is
restored to Microsoft Excel.

**Remarks**

  - > The custom application title, followed by the individual workbook
    > title, will appear in the application title bar if the workbook is
    > maximized.

  - > APP.TITLE does not affect DDE communications. You will still refer
    > to the application as "Excel".

> &nbsp;

**Related Function**

WINDOW.TITLE&nbsp;&nbsp;&nbsp;Changes the title of the active window

Return to [top](#A)

# ARGUMENT

Describes the arguments used in a custom function, which is a type of
macro, or in a subroutine. A custom function or subroutine must contain
one ARGUMENT function for each argument in the macro itself. There are
two forms of the ARGUMENT function. In the first form, only name\_text
is required; in the second form, only reference is required. Use the
first form if you want to store the argument as a name. Use the second
form if you want to store the argument in a specific cell or cells.

**Syntax 1**

For name storage

**ARGUMENT**(**name\_text**, data\_type\_num)

**Syntax 2**

For cell storage

**ARGUMENT**(name\_text, data\_type\_num, **reference**)

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of the argument or of the
cells containing the argument. Name\_text is required if you omit
reference.

Data\_type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number that determines what
type of values Microsoft Excel accepts for the argument. The following
table lists the possible data types.

|                     |                   |
| ------------------- | ----------------- |
| **Data\_type\_num** | **Type of value** |
| 1                   | Number            |
| 2                   | Text              |
| 4                   | Logical           |
| 8                   | Reference         |
| 16                  | Error             |
| 64                  | Array             |

&nbsp;

  - > Data\_type\_num can be a sum of the preceding different numbers to
    > allow for more than one possible type of data. For example, if
    > data\_type\_num is 7, which is the sum of 1, 2, and 4, then the
    > value can be a number, text, or logical value.

  - > Data\_type\_num is an optional argument. If you omit
    > data\_type\_num, it is assumed to be 7.

  - > If the value that is passed to the function macro is not of the
    > type specified by data\_type\_num, Microsoft Excel first attempts
    > to convert it to the specified type. If the value cannot be
    > converted, the macro returns the \#VALUE\! error value.

> &nbsp;

Reference&nbsp;&nbsp;&nbsp;&nbsp;is the cell or cells in which you want
to store the argument's value.

  - > If you specify reference, the value that is passed to ARGUMENT is
    > entered as a constant in the specified cell, and name\_text
    > becomes an optional argument because you can refer to the cell
    > with either reference or name\_text.

  - > If you omit reference, name\_text is defined on the macro sheet
    > and refers to the value that is passed to ARGUMENT. Once
    > name\_text is defined, you can use it in formulas.

> &nbsp;

**Remarks**

  - > Custom functions and subroutines can accept from 1 to 29
    > arguments.

  - > If a macro contains an ARGUMENT function and you omit the
    > corresponding argument in the function that starts the macro, the
    > macro uses the \#N/A error value as the value of the argument.

> &nbsp;

**Examples**

To create a custom function that calculates profit, use the following
functions to specify arguments for cost, sales, and sales volume:

ARGUMENT("UnitsSold", 1)

ARGUMENT("UnitCost", 1)

ARGUMENT("UnitPrice", 1)

**Related Functions**

RESULT&nbsp;&nbsp;&nbsp;Specifies the data type a custom function
returns

VOLATILE&nbsp;&nbsp;&nbsp;Makes custom functions recalculate
automatically

Return to [top](#A)

# ARRANGE.ALL

Equivalent to clicking the Arrange command on the Window menu.
Rearranges open windows and icons and resizes open windows. Also can be
used to synchronize scrolling of windows of the active sheet.

**Syntax**

**ARRANGE.ALL**(arrange\_num, active\_doc, sync\_horiz, sync\_vert)

**ARRANGE.ALL**?(arrange\_num, active\_doc, sync\_horiz, sync\_vert)

Arrange\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 7 specifying
how to arrange the windows.

|                  |                                                                                                             |
| ---------------- | ----------------------------------------------------------------------------------------------------------- |
| **Arrange\_num** | **Result**                                                                                                  |
| 1 or omitted     | Tiled (also used to arrange icons in Microsoft Excel for Windows)                                           |
| 2                | Horizontal                                                                                                  |
| 3                | Vertical                                                                                                    |
| 4                | None                                                                                                        |
| 5                | Horizontally arranges and sizes the windows based on the position of the active cell.                       |
| 6                | Vertically arranges and sizes the windows based on the position of the active cell.                         |
| 7                | Arranges windows so that they cascade from the upper left to the bottom right of the application workspace. |

If you want to change whether the windows are synchronized for scrolling
but not how they are arranged, make sure arrange\_num is 4.

Active\_doc&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying which
windows to arrange. If active\_doc is TRUE, Microsoft Excel arranges
only windows on the active workbook; if FALSE or omitted, all open
windows are arranged.

Sync\_horiz&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Sync Horizontal check box in Microsoft Excel version 4.0.

  - > If sync\_horiz is TRUE, Microsoft Excel selects the check box and
    > synchronizes horizontal scrolling.

  - > If sync\_horiz is FALSE or omitted, Microsoft Excel clears the
    > check box, and windows will not be synchronized when you scroll
    > horizontally.

  - > This argument is used only when active\_doc is TRUE.

> &nbsp;

Sync\_vert&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Sync Vertical check box in Microsoft Excel version 4.0.

  - > If sync\_vert is TRUE, Microsoft Excel selects the check box and
    > synchronizes vertical scrolling.

  - > If sync\_vert is FALSE or omitted, Microsoft Excel clears the
    > check box, and windows will not be synchronized when you scroll
    > vertically.

  - > This argument is used only when active\_doc is TRUE.

> &nbsp;

**Note**&nbsp;&nbsp;&nbsp;If arguments are omitted in the dialog box
form of this function, the default values are the previous settings, if
any; otherwise the default values are as described above.

**Remarks**

  - > After arranging windows, the top or leftmost window is active.

  - > In Microsoft Excel for Windows, if all windows are minimized,
    > ARRANGE.ALL ignores its arguments, if any, and arranges the
    > corresponding icons horizontally along the bottom of the
    > workspace.

> &nbsp;

**Tip**&nbsp;&nbsp;&nbsp;You can use synchronized horizontal or vertical
scrolling when you need to scroll while viewing macro formulas in one
window and corresponding macro values in another window of the same
macro sheet.

**Related Function**

ACTIVATE&nbsp;&nbsp;&nbsp;Switches to a window

Return to [top](#A)

# ASSIGN.TO.OBJECT

Assigns a macro to the currently select object.

**Syntax**

**ASSIGN.TO.OBJECT**(**macro\_ref**)

**ASSIGN.TO.OBJECT**?(macro\_ref)

Macro\_ref&nbsp;&nbsp;&nbsp;&nbsp;is the name of, or a reference to, the
macro you want to run when the object is clicked. If macro\_ref is
omitted, Microsoft Excel no longer runs the previously specified macro
(ASSIGN.TO.OBJECT is turned off).

**Remarks**

  - > If an object is not selected, ASSIGN.TO.OBJECT returns the
    > \#VALUE\! error value and interrupts the macro.

  - > To change the macro assigned to an object, select the object and
    > use ASSIGN.TO.OBJECT again, using the reference to the new macro
    > as macro\_ref. The previous macro is replaced with the new macro.

> &nbsp;

**Related Functions**

CREATE.OBJECT&nbsp;&nbsp;&nbsp;Creates an object

RUN&nbsp;&nbsp;&nbsp;Runs a macro

Return to [top](#A)

# ASSIGN.TO.TOOL

Assigns a macro to be run when a button is clicked with the mouse.

**Syntax**

**ASSIGN.TO.TOOL**(**bar\_id, position**, macro\_ref)

Bar\_id&nbsp;&nbsp;&nbsp;&nbsp;specifies the number or name of a toolbar
to which you want to assign a macro. For more information about bar\_id,
see ADD.TOOL.

Position&nbsp;&nbsp;&nbsp;&nbsp;specifies the position of the button
within the toolbar. Position starts with 1 at the left side (if
horizontal) or at the top (if vertical).

Macro\_ref&nbsp;&nbsp;&nbsp;&nbsp;is the name of, or a reference to, the
macro you want to run when the button is clicked. If macro\_ref is
omitted, Microsoft Excel no longer runs the previously specified macro.
After canceling the macro, if the button is a built-in button, Microsoft
Excel performs the normal default action when the button is clicked. If
the button is a custom button, Microsoft Excel displays the Assign Macro
dialog box when the button is clicked.

**Related Functions**

ADD.TOOL&nbsp;&nbsp;&nbsp;Adds one or more buttons to a toolbar

GET.TOOL&nbsp;&nbsp;&nbsp;Returns information about a button or buttons
on a toolbar

Return to [top](#A)

# ATTACH.TEXT

Attaches text to certain parts of the selected chart. Use ATTACH.TEXT to
attach text as a title or as a label for an axis or data point.

**Syntax**

**ATTACH.TEXT**(**attach\_to\_num**, series\_num, point\_num)

**ATTACH.TEXT**?(attach\_to\_num, series\_num, point\_num)

Attach\_to\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies which item on a chart
to attach text to. Attach\_to\_num is different for 2-D and 3-D charts.
Attach\_to\_num values for 2-D charts are shown in the following table.

|                     |                             |
| ------------------- | --------------------------- |
| **Attach\_to\_num** | **Attaches text to**        |
| 1                   | Chart title                 |
| 2                   | Value (y) axis              |
| 3                   | Category (x) axis           |
| 4                   | Series and data point       |
| 5                   | Secondary value (y) axis    |
| 6                   | Secondary category (x) axis |

Attach\_to\_num values for 3-D charts are shown in the following table.

|                     |                       |
| ------------------- | --------------------- |
| **Attach\_to\_num** | **Attaches text to**  |
| 1                   | Chart title           |
| 2                   | Value (z) axis        |
| 3                   | Series (y) axis       |
| 4                   | Category (x) axis     |
| 5                   | Series and data point |

Series\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the series number if
attach\_to\_num specifies a series or data point. If attach\_to\_num
specifies a series or data point and series\_num is omitted, the macro
is interrupted.

Point\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of the data
point, but only if you specify a series number. Point\_num is required
if series\_num is specified, unless the chart is an area chart.

**Remarks**

When you record adding an axis title or a chart title, Microsoft Excel
records both an ATTACH.TEXT function to attach the text and a
FONT.PROPERTIES function to make the text bold.

**Example**

The following macro functions attach the text "Quarterly Sales" to the x
(category) axis of the selected chart:

ATTACH.TEXT(3)

FORMULA("Quarterly Sales")

**Related Functions**

DATA.LABEL&nbsp;&nbsp;&nbsp;Assigns text labels to point on a chart

FORMULA&nbsp;&nbsp;&nbsp;Enters values into a cell or range or onto a
chart

Return to [top](#A)

# ATTACH.TOOLBARS

Displays the Attach Toolbars dialog box, which you use to attach or
associate toolbars with documents. The Attach Toolbars dialog box is
available when you click the Customize command (View menu, Toolbars
submenu), click the Toolbars tab, and then click the Attach button.

**Syntax**

**ATTACH.TOOLBARS**?( )

Return to [top](#A)

# AUTO.OUTLINE

Equivalent to clicking the Auto Outline command on the Group And Outline
submenu of the Data menu. Creates an outline within the selection. If a
single cell is selected, creates an outline for the entire sheet.

**Syntax**

**AUTO.OUTLINE**( )

**Related Functions**

CLEAR.OUTLINE&nbsp;&nbsp;&nbsp;Removes outlining from the current sheet

OUTLINE&nbsp;&nbsp;&nbsp;Creates an outline and defines settings for
automatically creating outlines

Return to [top](#A)

# AXES

Controls whether the axes on a chart are visible. There are two syntax
forms of this function. Syntax 1 is for 2-D charts; syntax 2 is for 3-D
charts.

**Syntax 1**

For 2-D charts

**AXES**(x\_primary, y\_primary, x\_secondary, y\_secondary)

**AXES**?(x\_primary, y\_primary, x\_secondary, y\_secondary)

**Syntax 2**

For 3-D charts

**AXES**(x\_primary, y\_primary, z\_primary)

**AXES**?(x\_primary, y\_primary, z\_primary)

Arguments are logical values corresponding to the check boxes in the
Axes dialog box.

  - > If an argument is TRUE, Microsoft Excel selects the check box and
    > displays the corresponding axis.

  - > If an argument is FALSE, Microsoft Excel clears the check box and
    > hides the corresponding axis.

  - > If an argument is omitted, the display of that axis is unchanged.

> &nbsp;

X\_primary&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the primary category
(x) axis.

Y\_primary&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the primary value (y)
axis.

Z\_primary&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the value (z) axis on
the primary 3-D chart.

X\_secondary&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the secondary
category (x) axis on a 2-D chart only.

Y\_secondary&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the secondary value
(y) axis on a 2-D chart only.

If a 2-D chart has no secondary axis, only the first two arguments are
used.

**Related Function**

GRIDLINES&nbsp;&nbsp;&nbsp;Controls whether chart gridlines are visible

Return to [top](#A)

# BEEP

Sounds a tone. Use BEEP to signal a message, a dialog box, or the end of
a macro, or whenever you need to get the user's attention.

**Syntax**

**BEEP**(tone\_num)

Tone\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 specifying the
tone to be played.

  - > On most computers, all numbers produce the same sound, the sound
    > that you hear when an error occurs or when you click outside some
    > dialog boxes.

  - > If tone\_num is omitted, it is assumed to be 1.

> &nbsp;

**Remarks**

  - > With a Macintosh, you can control the volume of the tone by using
    > the Control Panel desk accessory.

  - > With Microsoft Windows version 3.0 or later, you can turn off the
    > tone by using the Control Panel.

> &nbsp;

**Related Functions**

ALERT&nbsp;&nbsp;&nbsp;Displays a dialog box and a message

MESSAGE&nbsp;&nbsp;&nbsp;Displays a message in the status bar

Return to [top](#A)

# BORDER

Equivalent to clicking the Border tab in the Format Cells dialog box,
which appears when you click the Cells command on the Format menu. Adds
a border to the selected cell or range of cells.

**Syntax**

**BORDER**(outline, left, right, top, bottom, shade, outline\_color,
left\_color, right\_color, top\_color, bottom\_color)

**BORDER**?(outline, left, right, top, bottom, shade, outline\_color,
left\_color, right\_color, top\_color, bottom\_color)

Outline, left, right, top, and bottom are numbers from 0 to 7
corresponding to the line styles in the Border dialog box, as shown in
the following table.

|              |               |
| ------------ | ------------- |
| **Argument** | **Line type** |
| 0            | No border     |
| 1            | Thin line     |
| 2            | Medium line   |
| 3            | Dashed line   |
| 4            | Dotted line   |
| 5            | Thick line    |
| 6            | Double line   |
| 7            | Hairline      |

**Note&nbsp;&nbsp;&nbsp;**For compatibility with earlier versions of
Microsoft Excel, TRUE and FALSE values for the above arguments create a
thin border or no border, respectively.

Shade&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Shade check box in the
Border dialog box of Microsoft Excel version 4.0. This argument is
included for compatibility only.

Outline\_color, left\_color, right\_color, top\_color, and bottom\_color
are numbers from 1 to 56 corresponding to the Color box in the Border
dialog box. Zero corresponds to automatic color.

Return to [top](#A)

# BREAK

Interrupts a FOR-NEXT, a FOR.CELL-NEXT, or a WHILE-NEXT loop. If BREAK
is encountered within a loop, that loop is terminated and the macro
proceeds to the statement following the NEXT statement at the end of the
current loop.

**Syntax**

**BREAK**( )

**Example**

Use BREAK to test for conditions not anticipated by the FOR or WHILE
statement. For example, use the BREAK nested in an IF statement to exit
a WHILE-NEXT loop when a certain value is encountered:

\=IF(Counter=8, BREAK())

**Related Functions**

FOR&nbsp;&nbsp;&nbsp;Starts a FOR-NEXT loop

FOR.CELL&nbsp;&nbsp;&nbsp;Starts a FOR.CELL-NEXT loop

NEXT&nbsp;&nbsp;&nbsp;Ends a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

WHILE&nbsp;&nbsp;&nbsp;Starts a WHILE-NEXT loop

Return to [top](#A)

# BRING.TO.FRONT

Puts the selected object or objects on top of all other objects. For
example, if some worksheet objects are covering part of an embedded
chart, you can select the chart and use BRING.TO.FRONT to display the
chart on top of the worksheet objects.

**Syntax**

**BRING.TO.FRONT**( )

If the selection is not an object or a group of objects, BRING.TO.FRONT
returns the \#VALUE\! error value.

**Related Function**

SEND.TO.BACK&nbsp;&nbsp;&nbsp;Sends selected objects behind other
objects

Return to [top](#A)

# CALCULATE.DOCUMENT

Equivalent to choosing the Calc Sheet button in the Calculation tab on
the Options dialog, which appears when you choose the Options command
from the Tools menu. Calculates only the active worksheet.

**Syntax**

**CALCULATE.DOCUMENT**( )

**Remarks**

If a chart is active, CALCULATE.DOCUMENT returns the \#VALUE\! error
value.

**Related Functions**

CALCULATE.NOW&nbsp;&nbsp;&nbsp;Calculates all open workbooks immediately

CALCULATION&nbsp;&nbsp;&nbsp;Controls calculation settings

Return to [top](#A)

# CALCULATE.NOW

Equivalent to clicking the Calculation tab in the Options dialog box and
then clicking the Calc Now button. Calculates all sheets in all open
workbooks. Use CALCULATE.NOW to calculate all open workbooks when
calculation is set to manual.

**Syntax**

**CALCULATE.NOW**( )

**Related Functions**

CALCULATE.DOCUMENT&nbsp;&nbsp;&nbsp;Calculates the active sheet only

CALCULATION&nbsp;&nbsp;&nbsp;Controls calculation settings

Return to [top](#A)

# CALCULATION

Controls when and how formulas in open workbooks are calculated. This
function is included for compatibility with Microsoft Excel version 4.0.
For controlling calculation in Microsoft Excel version 5.0 or later, see
OPTIONS.CALCULATION.

**Syntax**

**CALCULATION**(**type\_num**, iter, max\_num, max\_change, update,
precision, date\_1904, calc\_save, save\_values, alt\_exp, alt\_form)  
**CALCULATION**?(type\_num, iter, max\_num, max\_change, update,
precision, date\_1904, calc\_save, save\_values, alt\_exp, alt\_form)

Arguments correspond to check boxes and options in the Calculation
dialog box. Arguments that correspond to check boxes are logical values.
If an argument is TRUE, Microsoft Excel selects the check box; if FALSE,
Microsoft Excel clears the check box.

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 indicating the
type of calculation.

|               |                         |
| ------------- | ----------------------- |
| **Type\_num** | **Type of calculation** |
| 1             | Automatic               |
| 2             | Automatic except tables |
| 3             | Manual                  |

Iter&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Iteration check box. The
default is FALSE.

Max\_num&nbsp;&nbsp;&nbsp;&nbsp;is the maximum number of iterations. The
default is 100.

Max\_change&nbsp;&nbsp;&nbsp;&nbsp;is the maximum change of each
iteration. The default is 0.001.

Update&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Update Remote
References check box. The default is TRUE.

Precision&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Precision As
Displayed check box. The default is FALSE.

Date\_1904&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the 1904 Date System
check box. The default is FALSE in Microsoft Excel for Windows and TRUE
in Microsoft Excel for the Macintosh.

Calc\_save&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Recalculate Before
Save check box. If calc\_save is FALSE, the workbook is not recalculated
before saving when in manual calculation mode. The default is TRUE.

Save\_values&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Save External
Link Values check box. The default is TRUE.

Alt\_exp&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Transition Formula
Evaluation check box in the Transition tab of the Options dialog box.

  - > If alt\_exp is TRUE, Microsoft Excel uses a set of rules
    > compatible with that of Lotus 1-2-3 when calculating formulas.
    > Text is treated as 0; TRUE and FALSE are treated as 1 and 0; and
    > certain characters in database criteria ranges are interpreted the
    > same way Lotus 1-2-3 interprets them.

  - > If alt\_exp is FALSE or omitted, Microsoft Excel calculates
    > normally.

> &nbsp;

Alt\_form&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Transition Formula
Entry check box in the Transition tab of the Options dialog box.

  - > This argument is available only in Microsoft Excel for Windows.

  - > If alt\_form is TRUE, Microsoft Excel accepts formulas entered in
    > Lotus 1-2-3 style.

  - > If alt\_form is FALSE or omitted, Microsoft Excel only accepts
    > formulas entered in Microsoft Excel style.

> &nbsp;

**Note&nbsp;&nbsp;&nbsp;**Microsoft Excel for Windows and Microsoft
Excel for the Macintosh use different date systems as their default.
Excel for Windows uses the 1900 date system, in which serial numbers
correspond to the dates January 1, 1900, through December 31, 9999.
Excel for the Macintosh uses the 1904 date system, in which serial
numbers correspond to the dates January 1, 1904, through December 31,
9999.

**Remarks**

Use GET.DOCUMENT to return the current calculation settings for your
workbook. For more information, see GET.DOCUMENT.

**Related Functions**

CALCULATE.DOCUMENT&nbsp;&nbsp;&nbsp;Calculates the active sheet only

CALCULATE.NOW&nbsp;&nbsp;&nbsp;Calculates all open workbooks immediately

GET.DOCUMENT&nbsp;&nbsp;&nbsp;Returns information about a workbook

OPTIONS.CALCULATION&nbsp;&nbsp;&nbsp;Controls calculation

OPTIONS.TRANSITION&nbsp;&nbsp;&nbsp;Controls transition options

Return to [top](#A)

# CALLER

Returns information about the cell, range of cells, command on a menu,
tool on a toolbar, or object that called the macro that is currently
running. Use CALLER in a subroutine or custom function whose behavior
depends on the location, size, name, or other attribute of the caller.

**Syntax**

**CALLER**( )

**Remarks**

  - > If the custom function is entered in a single cell, CALLER returns
    > the reference of that cell.

  - > If the custom function was part of an array formula entered in a
    > range of cells, CALLER returns the reference of the range.

  - > If CALLER appears in a macro called by an Auto\_Open, Auto\_Close,
    > Auto\_Activate, or Auto\_Deactivate macro, it returns the name of
    > the calling sheet.

  - > If CALLER appears in a macro called by a command on a menu, it
    > returns a horizontal array of three elements including the
    > command's position number, the menu number, and the menu bar
    > number.

  - > If CALLER appears in a macro called by an assigned-to-object
    > macro, it returns the object identifier.

  - > If CALLER appears in a macro called by a tool on a toolbar, it
    > returns a horizontal array containing the position number and the
    > toolbar name.

  - > If CALLER appears in a macro called by an ON.DOUBLECLICK or
    > ON.ENTRY function, CALLER returns the name of the chart object
    > identifier or cell reference, if applicable, to which the
    > ON.DOUBLECLICK or ON.ENTRY macro applies.

  - > If CALLER appears in a macro that was run manually, or for any
    > reason not described above, it returns the \#REF\! error value.

> &nbsp;

**Examples**

If the custom function MACROS\!VALUEONE is entered in cell B3 on a sheet
named SALES, the nested CALLER function returns the following values.

|                       |             |
| --------------------- | ----------- |
| **Nested function**   | **Returns** |
| COLUMN(CALLER())      | 2           |
| COLUMNS(CALLER())     | 1           |
| GET.CELL(1, CALLER()) | SALES\!$B$3 |
| ROW(CALLER())         | 3           |
| ROWS(CALLER())        | 1           |

If the same custom function was entered into an array in cells B2:C3,
the following values would be returned.

|                     |             |
| ------------------- | ----------- |
| **Nested function** | **Returns** |
| COLUMN(CALLER())    | 2           |
| COLUMNS(CALLER())   | 2           |
| ROW(CALLER())       | 2           |
| ROWS(CALLER())      | 2           |

**Related Functions**

GET.BAR&nbsp;&nbsp;&nbsp;Returns the name or position number of menu
bars, menus, and commands

GET.CELL&nbsp;&nbsp;&nbsp;Returns information about the specified cell

Return to [top](#A)

# CANCEL.COPY

Equivalent to pressing ESC in Microsoft Excel for Windows or ESC or
COMMAND+PERIOD in Microsoft Excel for the Macintosh to cancel the
marquee after you copy or cut a selection.

**Syntax**

**CANCEL.COPY**(render\_logical)

Render\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that, if TRUE,
places the contents of the Microsoft Excel Clipboard on the Clipboard
or, if FALSE or omitted, does not place them on the Clipboard.
Render\_logical is available only in Microsoft Excel for the Macintosh.

Return to [top](#A)

# CANCEL.KEY

Disables macro interruption, or specifies a macro to run when a macro is
interrupted. Use CANCEL.KEY to control what happens when a macro is
interrupted.

**Syntax**

**CANCEL.KEY**(**enable**, macro\_ref)

Enable&nbsp;&nbsp;&nbsp;&nbsp;specifies whether the macro can be
interrupted by pressing ESC in Microsoft Excel for Windows or ESC or
COMMAND+PERIOD in Microsoft Excel for the Macintosh.

|                                  |                                                           |
| -------------------------------- | --------------------------------------------------------- |
| **If enable is**                 | **Then**                                                  |
| FALSE                            | Pressing ESC or COMMAND+PERIOD does not interrupt a macro |
| TRUE and macro\_ref is omitted   | Pressing ESC or COMMAND+PERIOD interrupts a macro         |
| TRUE and macro\_ref is specified | Macro\_ref runs when ESC or COMMAND+PERIOD is pressed     |

Macro\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to a macro, as a cell
reference or a name, that runs when enable is TRUE and ESC or
COMMAND+PERIOD is pressed.

**Remarks**

  - > CANCEL.KEY affects only the macro that is currently running. Once
    > the macro is stopped by a RETURN or HALT function, ESC or
    > COMMAND+PERIOD is reactivated.

  - > When CANCEL.KEY is in effect, users can still cancel a dialog box
    > displayed while the macro is running.

> &nbsp;

**Examples**

The following macro formula prevents the macro from being interrupted by
pressing ESC or COMMAND+PERIOD:

CANCEL.KEY(FALSE)

The following macro formula reactivates ESC or COMMAND+PERIOD to cancel
macro execution:

CANCEL.KEY(TRUE)

The following line in a macro runs CheckCancel when ESC or
COMMAND+PERIOD is pressed:

CANCEL.KEY(TRUE, CheckCancel)

**Related Functions**

ERROR&nbsp;&nbsp;&nbsp;Specifies an action to take if an error occurs
while a macro is running

ON.KEY&nbsp;&nbsp;&nbsp;Runs a macro when a specified key is pressed

ON.TIME&nbsp;&nbsp;&nbsp;Runs a macro at a specified time

Return to [top](#A)

# CELL.PROTECTION

Equivalent to choosing the Protection tab in the Format Cells dialog
box, which appears when you choose the Cells command from the Format
menu. Allows you to control cell protection and display.

Arguments are logical values corresponding to check boxes in the
Protection tab. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box. If an
argument is omitted and the setting has been previously changed from the
defaults, the setting is not changed.

**Syntax**

**CELL.PROTECTION**(locked, hidden)

**CELL.PROTECTION**?(locked, hidden)

Locked&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Locked check box. The
default is TRUE.

Hidden&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Hidden check box. The
default is FALSE.

**Remarks**

Options selected in the Protection tab of the Format Cells dialog box or
with the CELL.PROTECTION function are activated only when the Protect
Sheet command is chosen from the Protection submenu on the Tools menu or
the PROTECT.DOCUMENT function is used to select protection.

**Related Functions**

PROTECT.DOCUMENT&nbsp;&nbsp;&nbsp;Controls protection for the active
sheet

SAVE.AS&nbsp;&nbsp;&nbsp;Saves a workbook and allows you to specify the
name, file type, password, backup file, and location of the workbook

Return to [top](#A)

# CHANGE.LINK

Equivalent to clicking the Change Source button in the Links dialog box,
which appears when you click the Links command on the Edit menu. Changes
a link from one supporting workbook to another.

**Syntax**

**CHANGE.LINK**(**old\_text, new\_text**, type\_of\_link)

**CHANGE.LINK**?(old\_text, new\_text, type\_of\_link)

Old\_text&nbsp;&nbsp;&nbsp;&nbsp;is the path of the link from the active
dependent workbook you want to change.

New\_text&nbsp;&nbsp;&nbsp;&nbsp;is the path of the link you want to
change to.

Type\_of\_link&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 specifying
what type of link you want to change.

|                    |                        |
| ------------------ | ---------------------- |
| **Type\_of\_link** | **Link document type** |
| 1 or omitted       | Microsoft Excel link   |
| 2                  | DDE/OLE link           |

**Remarks**

The workbook whose links you want to change must be active when this
function is calculated.

**Related Functions**

GET.LINK.INFO&nbsp;&nbsp;&nbsp;Returns information about a link

LINKS&nbsp;&nbsp;&nbsp;Returns the name of all linked workbooks

OPEN.LINKS&nbsp;&nbsp;&nbsp;Opens specified supporting workbooks

SET.UPDATE.STATUS&nbsp;&nbsp;&nbsp;Controls the update status of a link

UPDATE.LINK&nbsp;&nbsp;&nbsp;Updates a link to another workbook

Return to [top](#A)

# CHART.ADD.DATA

Equivalent to dragging data from a worksheet onto a chart. Adds data to
an existing chart.

**Syntax**

**CHART.ADD.DATA**(**ref**, rowcol, titles, categories, replace, series)

Ref&nbsp;&nbsp;&nbsp;&nbsp;is the cell reference for the data that is
being dragged onto the chart

Rowcol&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies whether
the values corresponding to a particular data series are in rows or
columns. Enter 1 for rows or 2 for columns.

Titles&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
Series Names In First Column check box (or First Row, depending on the
value of rowcol) in the Paste Special dialog box.

  - > If titles is TRUE, Microsoft Excel selects the check box and uses
    > the contents of the cell in the first column of each row (or first
    > row of each column) as the name of the data series in that row (or
    > column).

  - > If titles is FALSE, Microsoft Excel clears the check box and uses
    > the contents of the cell in the first column of each row (or first
    > row of each column) as the first data point of the data series.

Categories&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Categories (X Labels) In First Row (or First Column, depending on
the value of rowcol) check box in the Paste Special dialog box.

  - > If categories is TRUE, Microsoft Excel selects the check box and
    > uses the contents of the first row (or column) of the selection as
    > the categories for the chart.

  - > If categories is FALSE, Microsoft Excel clears the check box and
    > uses the contents of the first row (or column) as the first data
    > series in the chart.

> &nbsp;

Replace&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
Replace Existing Categories check box in the Paste Special dialog box.

  - > If replace is TRUE, Microsoft Excel selects the check box and
    > applies categories while replacing existing categories with
    > information from the copied cell range.

  - > If replace is FALSE, Microsoft Excel clears the check box and
    > applies new categories without replacing any old ones.

Series&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying how cells are added
to a chart.

|            |              |
| ---------- | ------------ |
| **Series** | **Added as** |
| 1          | New series   |
| 2          | New point(s) |

Return to [top](#A)

# CHART.TREND

A trendline can be added to only to the these chart types: bar, column,
stacked column, scatter, line, and area.

**Syntax**

**CHART.TREND**(**type**, ord\_per, forecast, backcast, intercept,
equation, r\_squared, name)

Type&nbsp;&nbsp;&nbsp;&nbsp;is the type of trend or regression.

|            |                |
| ---------- | -------------- |
| **Number** | **Type used**  |
| 1          | Linear         |
| 2          | Logarithmic    |
| 3          | Polynomial     |
| 4          | Power          |
| 5          | Exponential    |
| 6          | Moving Average |

Ord\_per&nbsp;&nbsp;&nbsp;&nbsp;depends on type. If type is 3, then
ord\_per is the order of the polynomial. If type is 6, ord\_per is the
number of periods for the moving average. If type is neither 3 nor 6,
then ord\_per is ignored.

Forecast&nbsp;&nbsp;&nbsp;&nbsp;is the number of periods or units to
extrapolate the trendline in the positive or forward direction. This
argument is ignored for moving averages (type 6). The default is zero.

Backcast&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the number of
periods or units to extrapolate the trendline in the negative or
backward direction. This argument is ignored for moving averages (type
6). The default is zero.

Intercept&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the value of the
y-intercept of the trendline, if it is already known. If FALSE or
omitted, Microsoft Excel will calculate the y-intercept . This argument
is ignored for moving averages.

Equation&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether
the trend equation should be displayed on the chart. If TRUE, the
equation will be displayed on the chart. If FALSE or omitted, the
equation will not be displayed on the chart.

R\_squared&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether
the r-squared equation should be displayed on the chart. If TRUE, the
value will be displayed on the chart. If FALSE or omitted, the equation
will not be displayed on the chart.

Name&nbsp;&nbsp;&nbsp;&nbsp;is a text string specifying the custom name
of the trendline. Can also be a logical value. If TRUE or omitted, the
automatic name will be used instead.

**Remarks**

  - > A trendline can not be added to a 3-D chart, a stacked chart, or
    > an 100% chart.

  - > The linear model calculates the least squares fit for a line
    > represented by the equation y&nbsp;=&nbsp;mx&nbsp;+&nbsp;b, where
    > m is the slope and b is the intercept.

  - > The logarithmic model calculates the least squares fit through
    > points using the equation y&nbsp;=&nbsp;c\*ln(x)&nbsp;+&nbsp;b,
    > where c and b are constants.

  - > The exponential model calculates the least squares fit through
    > points using the following equation:

> ![](./media/image9.png)
> 
> where c and b are constants.

  - > The polynomial model calculates the least squares fit through
    > points using the following equation:

> ![](./media/image10.png)
> 
> where b, c1, c2, c3, etc. are constants.

  - > The power model calculates the least squares fit through points
    > using the following equation:

> ![](./media/image11.png)
> 
> where b and c are constants.

**Related Function**

CHART.WIZARD&nbsp;&nbsp;&nbsp;Equivalent to clicking the ChartWizard
button on the Standard toolbar

Return to [top](#A)

# CHART.WIZARD

Equivalent to clicking the ChartWizard button on the standard or chart
toolbar. Creates a chart. It is generally easier to use the macro
recorder to enter this function on your macro sheet.

**Syntax**

**CHART.WIZARD**(long, **ref**, gallery\_num, type\_num, plot\_by,
categories, ser\_titles, legend, title, x\_title, y\_title, z\_title,
number\_cats, number\_titles)

**CHART.WIZARD**?(long, ref, gallery\_num, type\_num, plot\_by,
categories, ser\_titles, legend, title, x\_title, y\_title, z\_title,
number\_cats, number\_titles)

Long&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that determines which
type of ChartWizard button CHART.WIZARD is equivalent to.

  - > If long is TRUE, CHART.WIZARD is equivalent to using the five-step
    > ChartWizard button.

  - > If long is FALSE or omitted, CHART.WIZARD is equivalent to using
    > the two-step ChartWizard button, and gallery\_num, type\_num,
    > legend, and the title arguments are ignored.

> &nbsp;

Ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the range of cells on the
active worksheet that contains the source data for the chart, or the
object identifier of the chart if it has already been created.

Gallery\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 15 specifying
the type of chart you want to create.

|                  |              |
| ---------------- | ------------ |
| **Gallery\_num** | **Chart**    |
| 1                | Area         |
| 2                | Bar          |
| 3                | Column       |
| 4                | Line         |
| 5                | Pie          |
| 6                | Radar        |
| 7                | XY (scatter) |
| 8                | Combination  |
| 9                | 3-D area     |
| 10               | 3-D bar      |
| 11               | 3-D column   |
| 12               | 3-D line     |
| 13               | 3-D pie      |
| 14               | 3-D surface  |
| 15               | Doughnut     |

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number identifying a formatting
option. The first formatting option in any gallery is 1.

Plot\_by&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
whether the data for each data series is in rows or columns. 1 specifies
rows; 2 specifies columns. If plot\_by is omitted, Microsoft Excel uses
the appropriate value for the chart you're creating.

Categories&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
whether the first row or column contains a list of x-axis labels, or
data for the first data series. 1 specifies x-axis labels; 2 specifies
the first data series. If categories is omitted, Microsoft Excel uses
the appropriate value for the chart you're creating. If number\_cats is
specified, this argument is ignored.

Ser\_titles&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
whether the first column or row contains series titles, or data for the
first data point in each series. 1 specifies series titles; 2 specifies
the first data point. If ser\_titles is omitted, Microsoft Excel uses
the appropriate value for the chart you're creating. If number\_titles
is specified, this argument is ignored.

Legend&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies whether
to include a legend. 1 specifies a legend; 2 specifies no legend. If
legend is omitted, Microsoft Excel does not include a legend.

For the following arguments, if an argument is omitted or is empty text
(""), no title is specified.

Title&nbsp;&nbsp;&nbsp;&nbsp;is text that you want to use as a chart
title.

X\_title&nbsp;&nbsp;&nbsp;&nbsp;is text that you want to use as an
x-axis title.

Y\_title&nbsp;&nbsp;&nbsp;&nbsp;is text that you want to use as a y-axis
title.

Z\_title&nbsp;&nbsp;&nbsp;&nbsp;is text that you want to use as a z-axis
title.

Number\_cats&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of rows or
columns (depending on the value of plot\_by) to use for the category
labels in the chart. This argument overrides the categories argument.

Number\_titles&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of rows or
columns (depending on the value of plot\_by) to use for the series
labels in the chart. This argument overrides the ser\_title argument.

**Remarks**

If you are using the macro recorder, Microsoft Excel records a
CREATE.OBJECT and a COPY function when the chart is created and a
CHART.WIZARD function when the chart is formatted. You must precede this
function with a CREATE.OBJECT function if you are not using the macro
recorder.

**Related Function**

CREATE.OBJECT&nbsp;&nbsp;&nbsp;Creates an object

Return to [top](#A)

# CHECKBOX.PROPERTIES

Sets various properties of check box and option box controls on a
worksheet or dialog sheet.

**Syntax**

**CHECKBOX.PROPERTIES**(value, link, accel\_text, 3d\_shading,
accel\_text2)

**CHECKBOX.PROPERTIES**?(value, link, accel\_text, 3d\_shading,
accel\_text2,)

Value&nbsp;&nbsp;&nbsp;&nbsp;is the value of the check box or option
button setting that determines whether it is selected or not.

|            |                           |
| ---------- | ------------------------- |
| **Value**  | **Box or Button Setting** |
| 0 or FALSE | Off                       |
| 1 or TRUE  | On                        |
| 2          | Mixed                     |

Link&nbsp;&nbsp;&nbsp;&nbsp;is the cell on the sheet to which the check
box or option button value is linked. Whenever one of these two controls
is changed, the value of the control is entered into the cell.
Similarly, whenever the value in the cell is changed, the setting for
the corresponding check box or option button is also changed. To clear
the link, set this value to an empty string. For example, entering
"TRUE" into a cell linked to a check box will select that check box.

3d\_shading&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether the check box appears as 3-D. If TRUE, the check box will appear
as 3-D. If FALSE or omitted, the check box will not be 3-D. This
argument is available for only worksheets.

Accel\_text&nbsp;&nbsp;&nbsp;&nbsp;is a text string containing the
character to use as the control's accelerator key on a dialog sheet. The
character is matched against the text of the control, and the first
matching character is underlined. When the user presses ALT+accel\_text
in Microsoft Excel for Windows or COMMAND+accel\_text in Microsoft Excel
for the Macintosh, the control is clicked.

Accel\_text2&nbsp;&nbsp;&nbsp;&nbsp;is a text string containing the
second accelerator key on a dialog sheet. This argument is for only Far
East versions of Microsoft Excel.

**Remarks**

Only controls on dialog sheets can have accelerator keys. Worksheet
controls cannot have accelerator keys.

**Related Functions**

PUSHBUTTON.PROPERTIES&nbsp;&nbsp;&nbsp;Sets the properties of the push
button control

EDITBOX.PROPERTIES&nbsp;&nbsp;&nbsp;Sets the properties of an edit box
on a worksheet or dialog sheet

LABEL.PROPERTIES&nbsp;&nbsp;&nbsp;Sets the accelerator property of the
label and group box control

LISTBOX.PROPERTIES&nbsp;&nbsp;&nbsp;Sets the properties of a list box
and drop-down box controls on a worksheet or dialog sheet

Return to [top](#A)

# CHECK.COMMAND

Adds or removes a check mark to or from a command name on a menu. A
check mark beside a command indicates that the command has been chosen.

**Syntax**

**CHECK.COMMAND**(**bar\_num**, **menu**, **command**, **check**,
position)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the menu bar containing the command.
Bar\_num can be the ID number of a built-in or custom menu bar.

Menu&nbsp;&nbsp;&nbsp;&nbsp;is the menu containing the command. Menu can
be either the name of a menu as text or the number of a menu. Menus are
numbered starting with 1 from the left of the screen.

Command&nbsp;&nbsp;&nbsp;&nbsp;is the command you want to check or the
submenu containing the command you want to check. Command can be the
name of the command as text or the number of the command; the first
command on a menu is in position 1.

Check&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
check mark. If check is TRUE, Microsoft Excel adds a check mark to the
command; if FALSE, Microsoft Excel removes the check mark.

position&nbsp;&nbsp;&nbsp;&nbsp;is the name of a command on a submenu
that you want to check.

**Remarks**

  - > The check mark doesn't affect execution of the command. Microsoft
    > Excel automatically adds and deletes check marks to some commands,
    > such as the name of the active workbook in the Window menu. If you
    > have assigned a check mark to a built-in command that Microsoft
    > Excel automatically changes in response to the user's actions, the
    > check mark will be added or removed as appropriate, and any check
    > marks you have added or deleted with CHECK.COMMAND will be
    > ignored.

  - > If you use CHECK.COMMAND with a command on a Microsoft Excel
    > version 4.0 menu bar, the corresponding command on the Microsoft
    > Excel version 5.0 or later menu bar will not be effected.

> &nbsp;

**Example**

The following macro formula adds a check mark to the Sales command on
the Weekly menu on a custom menu bar created by the ADD.BAR function in
a cell named Reports:

CHECK.COMMAND(Reports, "Weekly", "Sales", TRUE)

**Related Functions**

ADD.COMMAND&nbsp;&nbsp;&nbsp;Adds a command to a menu

DELETE.COMMAND&nbsp;&nbsp;&nbsp;Deletes a command from a menu

ENABLE.COMMAND&nbsp;&nbsp;&nbsp;Enables or disables a menu or custom
command

RENAME.COMMAND&nbsp;&nbsp;&nbsp;Changes the name of a command or menu

Return to [top](#A)

# CLEAR

Equivalent to clicking the Clear command on the Edit menu. Clears
contents, formats, notes, or all of these from the active worksheet or
macro sheet. Clears series or formats from the active chart.

**Syntax**

**CLEAR**(type\_num)

**CLEAR**?(type\_num)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 specifying what
to clear. Only values 1, 2, and 3 are valid if the selected item is a
chart.

On a worksheet or macro sheet, or if an entire chart is selected, the
following occurs.

|               |                                                                  |
| ------------- | ---------------------------------------------------------------- |
| **Type\_num** | **Clears**                                                       |
| 1             | All                                                              |
| 2             | Formats (if a chart, clears the chart format or clears pictures) |
| 3             | Contents (if a chart, clears all data series)                    |
| 4             | Comments (this does not apply to charts)                         |

On a chart, if a single point, an entire data series, error bars, or a
trend line is selected, the following occurs.

|               |                                                                 |
| ------------- | --------------------------------------------------------------- |
| **Type\_num** | **Clears**                                                      |
| 1             | Selected series, error bars, or trend line                      |
| 2             | Format in the selected point, series, error bars, or trend line |

If type\_num is omitted, the default values are set as shown in the
following table.

|                            |                           |
| -------------------------- | ------------------------- |
| **Active sheet**           | **Type\_num**             |
| Worksheet                  | 3                         |
| Macro sheet                | 3                         |
| Chart (with no selection)  | 1                         |
| Chart (with item selected) | Deletes the selected item |

**Related Function**

EDIT.DELETE&nbsp;&nbsp;&nbsp;Removes cells from a sheet

Return to [top](#A)

# CLEAR.OUTLINE

Equivalent to clicking the Clear Outline command on the Group And
Outline submenu of the Data menu. Clears the outline within the
selection. If a single cell is selected, it clears the outline from the
entire sheet.

**Syntax**

**CLEAR.OUTLINE**( )

**Related Functions**

AUTO.OUTLINE&nbsp;&nbsp;&nbsp;Creates an outline

OUTLINE&nbsp;&nbsp;&nbsp;Creates an outline and defines settings for
automatically creating outlines

Return to [top](#A)

# CLEAR.ROUTING.SLIP

Equivalent to the Clear button in the Routing Slip dialog box. Clears
the routing slip.

**Syntax**

**CLEAR.ROUTING.SLIP**(reset\_only\_logical)

Reset\_only\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that
specifies whether the routing slip should be cleared.

  - > This option is valid only after every recipient on the routing
    > slip has received and forwarded the workbook. Setting
    > reset\_only\_logical to TRUE in this case is equivalent to the
    > Reset button in the routing slip dialog.

  - > If some recipients have not received or routed the workbook,
    > reset\_only\_logical is ignored.

  - > If reset\_only\_logical is FALSE or omitted and the workbook has
    > been routed to all recipients, then the routing slip is removed
    > from the workbook. A new slip can be subsequently added using
    > ROUTING.SLIP.

Return to [top](#A)

# CLOSE

Closes the active window. In Microsoft Excel for Windows, CLOSE is
equivalent to clicking the Close command on the Document Control menu.
In Microsoft Excel for the Macintosh, CLOSE is equivalent to clicking
the close box.

**Syntax**

**CLOSE**(save\_logical, route\_logical)

Save\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether to save the file before closing the window.

|                   |                                                                                               |
| ----------------- | --------------------------------------------------------------------------------------------- |
| **Save\_logical** | **Result**                                                                                    |
| TRUE              | Saves the file                                                                                |
| FALSE             | Does not save the file                                                                        |
| Omitted           | If you've made changes to the file, displays a dialog box asking if you want to save the file |

Route\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether to route the file after closing it. This argument is ignored if
there is not a routing slip present.

|                    |                                                                                                       |
| ------------------ | ----------------------------------------------------------------------------------------------------- |
| **Route\_logical** | **Result**                                                                                            |
| TRUE               | Routes the file                                                                                       |
| FALSE              | Does not route the file                                                                               |
| Omitted            | If you've specified recipients for routing, displays a dialog box asking if you want to save the file |

**Remarks**

Users of Microsoft Excel versions earlier than 4.0 should note that if
the macro sheet containing the function is the active sheet, CLOSE now
closes the workbook.

**Note&nbsp;&nbsp;&nbsp;**When you use the CLOSE function, Microsoft
Excel does not run any Auto\_Close macros before closing the workbook.

**Related Functions**

CLOSE.ALL&nbsp;&nbsp;&nbsp;Closes all windows

FILE.CLOSE&nbsp;&nbsp;&nbsp;Closes the active workbook

SAVE&nbsp;&nbsp;&nbsp;Saves the active workbook

Return to [top](#A)

# CLOSE.ALL

Equivalent to clicking the Close All command on the File menu. The Close
All command appears when you hold down SHIFT while selecting the File
menu. Closes all protected and unprotected windows and all hidden
windows. If unsaved changes have been made to the workbook in one or
more windows, a message is displayed asking if you want to save each
workbook.

**Syntax**

**CLOSE.ALL**( )

**Related Functions**

CLOSE&nbsp;&nbsp;&nbsp;Closes the active window

FILE.CLOSE&nbsp;&nbsp;&nbsp;Closes the active workbook

QUIT&nbsp;&nbsp;&nbsp;Ends a Microsoft Excel session

SAVE&nbsp;&nbsp;&nbsp;Saves the active workbook

Return to [top](#A)

# COLOR.PALETTE

Copies a color palette from an open workbook to the active workbook. Use
COLOR.PALETTE to share color palettes between workbooks.

**Syntax**

**COLOR.PALETTE**(**file\_text**)

**COLOR.PALETTE**?(file\_text)

File\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a workbook, as a text
string, that you want to copy a color palette from. The workbook
specified by file\_text must be open, or COLOR.PALETTE returns the
\#VALUE\! error value and interrupts the macro. If file\_text is empty
text (""), then COLOR.PALETTE sets colors to the default values.

**Related Function**

EDIT.COLOR&nbsp;&nbsp;&nbsp;Defines a color on the color palette

Return to [top](#A)

# COLUMN.WIDTH

Equivalent to choosing the Width command on the Column submenu of the
Format menu. Changes the width of the columns in the specified
reference.

**Syntax**

**COLUMN.WIDTH**(width\_num, reference, standard, type\_num,
standard\_num)

**COLUMN.WIDTH**?(width\_num, reference, standard, type\_num,
standard\_num)

Width\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies how wide you want the
columns to be in units of one character of the font corresponding to the
Normal cell style. Width\_num is ignored if standard is TRUE or if
type\_num is provided.

Reference&nbsp;&nbsp;&nbsp;&nbsp;specifies the columns for which you
want to change the width.

  - > If reference is specified, it must be either an external reference
    > to the active worksheet, such as \!$A:$C or \!Database, or an
    > R1C1-style reference in the form of text, such as "C1:C3",
    > "C\[-4\]:C\[-2\]", or "Database".

  - > If reference is a relative R1C1-style reference in the form of
    > text, it is assumed to be relative to the active cell.

  - > If reference is omitted, it is assumed to be the current
    > selection.

> &nbsp;

Standard\_num&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Standard Width command from the Column submenu on the Format menu.

  - > If standard is TRUE, Microsoft Excel sets the column width to the
    > currently defined standard (default) width and ignores width\_num.

  - > If standard is FALSE or omitted, Microsoft Excel sets the width
    > according to width\_num or type\_num.

> &nbsp;

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 corresponding
to the Hide, Unhide, or AutoFit Selection commands, respectively, on the
Column submenu of the Format menu.

|               |                                                                                                                                                     |
| ------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Action taken**                                                                                                                                    |
| 1             | Hides the column selection by setting the column width to 0                                                                                         |
| 2             | Unhides the column selection by setting the column width to the value set before the selection was hidden                                           |
| 3             | Sets the column selection to a best-fit width, which varies from column to column depending on the length of the longest data string in each column |

Standard\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies how wide the standard
width is, and is measured in points. If standard\_num is omitted, the
standard width setting remains unchanged.

**Remarks**

  - > Changing the value of standard\_num changes the width of all
    > columns except those that have been set to a custom value.

  - > If any of the argument settings conflict, such as when standard is
    > TRUE and type\_num is 3, Microsoft Excel uses the type\_num
    > argument and ignores any arguments that conflict with type\_num.

  - > If you are recording a macro while using a mouse and you change
    > column widths by dragging the column border, Microsoft Excel
    > records the references of the columns using R1C1-style references
    > in the form of text.

> &nbsp;

**Related Function**

ROW.HEIGHT&nbsp;&nbsp;&nbsp;Changes the heights of rows

Return to [top](#A)

# COMBINATION

Changes the format of the active chart to one of six built-in
combination chart types.

**Syntax**

**COMBINATION**(**type\_num**)

**COMBINATION**?(type\_num)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the
combination chart you want.

|               |                                                                                                                                             |
| ------------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Result**                                                                                                                                  |
| 1             | Column chart overlaid by a line chart                                                                                                       |
| 2             | Column chart overlaid by a line chart with an independent y-axis scale                                                                      |
| 3             | Line chart overlaid by a line chart with an independent y-axis scale                                                                        |
| 4             | Area chart overlaid by a column chart                                                                                                       |
| 5             | Column chart overlaid by a line chart containing three data series (for showing stock volumes related to high, low, and closing prices)     |
| 6             | Column chart overlaid by a line chart containing four data series (for showing stock volumes related to open, high, low, and closing prices |

**Related Functions**

FORMAT.MAIN&nbsp;&nbsp;&nbsp;Formats a main chart

FORMAT.OVERLAY&nbsp;&nbsp;&nbsp;Formats an overlay chart

Return to [top](#A)

# CONSOLIDATE

Equivalent to clicking the Consolidate command on the Data menu.
Consolidates data from multiple ranges on multiple worksheets into a
single range on a single worksheet.

**Syntax**

**CONSOLIDATE**(source\_refs, function\_num, top\_row, left\_col,
create\_links)

**CONSOLIDATE**?(source\_refs, function\_num, top\_row, left\_col,
create\_links)

Source\_refs&nbsp;&nbsp;&nbsp;&nbsp;are references to areas that contain
data to be consolidated on the destination worksheet. Source\_refs must
be in text form and include the full path of the file and the cell
reference or named ranges in the workbook to be consolidated.
Source\_refs are usually external references and must be given as an
array, for example: {"SHEET1\!IncomeOne", "SHEET2\!IncomeTwo"}.

To add or delete source\_refs from an existing consolidation on a
worksheet, reuse the CONSOLIDATE function, specifying the new
source\_refs.

Function\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 11 that
specifies one of the 11 functions you can use to consolidate data. If
function\_num is omitted, the SUM function, number 9, is used. The
functions and their corresponding numbers are listed in the following
table.

|                   |              |
| ----------------- | ------------ |
| **Function\_num** | **Function** |
| 1                 | AVERAGE      |
| 2                 | COUNT        |
| 3                 | COUNTA       |
| 4                 | MAX          |
| 5                 | MIN          |
| 6                 | PRODUCT      |
| 7                 | STDEV        |
| 8                 | STDEVP       |
| 9                 | SUM          |
| 10                | VAR          |
| 11                | VARP         |

The following arguments correspond to text boxes and check boxes in the
Consolidate dialog box. Arguments that correspond to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Top\_row&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Top Row check box.
The default is FALSE.

Left\_col&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Left Column check
box. The default is FALSE.

If top\_row and left\_col are both FALSE or omitted, the data is
consolidated by position.

Create\_links&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Create Links To
Source Data check box.

**Remarks**

  - > If you use the CONSOLIDATE function with no arguments and there is
    > a consolidation on the active worksheet, Microsoft Excel
    > reconsolidates, using the sources, function, and position
    > attributes used to create the existing consolidation.

  - > If you use the CONSOLIDATE function with no arguments and there is
    > no consolidation on the active worksheet, the function returns the
    > \#VALUE\! error value.

> &nbsp;

**Related Functions**

CHANGE.LINK&nbsp;&nbsp;&nbsp;Changes supporting workbook links

LINKS&nbsp;&nbsp;&nbsp;Returns the names of all linked workbooks

OPEN.LINKS&nbsp;&nbsp;&nbsp;Opens specified supporting workbooks

UPDATE.LINK&nbsp;&nbsp;&nbsp;Updates a link to another workbook

Return to [top](#A)

# CONSTRAIN.NUMERIC

Equivalent to pressing the Constrain Numeric button. The Constrain
Numeric button can be found in the Insert category on the Commands tab
(Customize dialog box). The Customize dialog box appears when you choose
the Toolbars command from the View menu and then choose the Customize
command. Constrains handwriting recognition to numbers and punctuation
only. Use this function in a macro to improve the accuracy of
handwriting recognition when the user is entering a series of numbers or
formulas.

**Note&nbsp;&nbsp;&nbsp;**This function is only available if you are
using Microsoft Windows for Pen Computing.

**Syntax**

**CONSTRAIN.NUMERIC**(numeric\_only)

Numeric\_only&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that turns the
numeric constraint on or off. If numeric\_only is TRUE, only numbers and
digits are recognized; if FALSE, all characters are recognized as usual.
if numeric\_only is omitted, the numeric constraint is toggled.

**Remarks**

When the numeric constraint is on, Microsoft Excel recognizes only the
following symbols:

0 1 2 3 4 5 6 7 8 9 $ \# @ % ( ) - + = { } : \< \> , ? | .

**Tip**&nbsp;&nbsp;&nbsp;Use GET.WORKSPACE(45) to make sure you're
running Microsoft Windows for Pen Computing.

Return to [top](#A)

# COPY

Equivalent to clicking the Copy command on the Edit menu. Copies and
pastes data or objects.

**Syntax**

**COPY**(from\_reference, to\_reference)

From\_reference&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cell or
range of cells you want to copy. If from\_reference is omitted, it is
assumed to be the current selection.

To\_reference&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cell or range
of cells where you want to paste what you have copied.

  - > To\_reference should be a single cell or an enlarged multiple of
    > from\_reference. For example, if from\_reference is a 2 by 4
    > rectangle, to\_reference can be a 4 by 8 rectangle.

  - > To\_reference can be omitted so that you can subsequently paste
    > using the PASTE, PASTE.LINK, or PASTE.SPECIAL functions.

> &nbsp;

**Related Functions**

CUT&nbsp;&nbsp;&nbsp;Cuts or moves data or objects

PASTE&nbsp;&nbsp;&nbsp;Pastes cut or copied data

PASTE.LINK&nbsp;&nbsp;&nbsp;Pastes copied data or objects and
establishes a link to the source of the data or object

PASTE.SPECIAL&nbsp;&nbsp;&nbsp;Pastes specific components of copied data

Return to [top](#A)

# COPY.CHART

Equivalent to choosing the Copy Chart command from the Edit menu in
Microsoft Excel for the Macintosh version 1.5 or earlier. This function
is included only for macro compatibility. You can copy a chart with the
COPY.PICTURE function by omitting the appearance\_num argument.

**Syntax**

**COPY.CHART**(**size\_num**)

Size\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number describing how to copy the
picture and is only available if the current selection is a chart.

|               |                                                                          |
| ------------- | ------------------------------------------------------------------------ |
| **Size\_num** | **Action**                                                               |
| 1 or omitted  | Copies the chart in the same size as the window on which it is displayed |
| 2             | Copies what you would see if you printed the chart                       |

**Related Function**

COPY.PICTURE&nbsp;&nbsp;&nbsp;Creates a picture of the current selection
for use in another program

Return to [top](#A)

# COPY.PICTURE

Equivalent to choosing the Copy Picture command from the Edit menu. The
Copy Picture command appears if you hold down SHIFT while choosing the
Edit menu. It copies a chart or range of cells to the Clipboard as a
graphic. Use COPY.PICTURE to create an image of the current selection or
chart for use in another program.

**Syntax**

**COPY.PICTURE**(appearance\_num, size\_num, type\_num)

**COPY.PICTURE**?(appearance\_num, size\_num, type\_num)

**Remarks**

Graphics are created differently on screen and on a printer. Thus, the
printed picture may look different from the one on screen.

Appearance\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number describing how to
copy the picture.

|                     |                                                                                 |
| ------------------- | ------------------------------------------------------------------------------- |
| **Appearance\_num** | **Action**                                                                      |
| 1 or omitted        | Copies a picture as closely as possible to the picture displayed on your screen |
| 2                   | Copies what you would see if you printed the selection                          |

Size\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number describing how to copy the
picture and is only available if the current selection is a chart.

|               |                                                                          |
| ------------- | ------------------------------------------------------------------------ |
| **Size\_num** | **Action**                                                               |
| 1 or omitted  | Copies the chart in the same size as the window on which it is displayed |
| 2             | Copies what you would see if you printed the chart                       |

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the format of
the picture. This argument is available only in Microsoft Excel for
Windows.

|               |                           |
| ------------- | ------------------------- |
| **Type\_num** | **Format of the picture** |
| 1 or omitted  | Picture                   |
| 2             | Bitmap                    |

**Related Functions**

COPY&nbsp;&nbsp;&nbsp;Copies and pastes data or objects

CUT&nbsp;&nbsp;&nbsp;Cuts or moves data or objects

PASTE&nbsp;&nbsp;&nbsp;Pastes cut or copied data

PASTE.PICTURE.LINK&nbsp;&nbsp;&nbsp;Pastes a linked picture of the
currently copied area

PASTE.SPECIAL&nbsp;&nbsp;&nbsp;Pastes specific components of copied data

Return to [top](#A)

# COPY.TOOL

Copies a button face to the Clipboard.

**Syntax**

**COPY.TOOL**(bar\_id, position)

Bar\_id&nbsp;&nbsp;&nbsp;&nbsp;specifies the number or name of a toolbar
from which you want to copy the button face. For detailed information
about bar\_id, see ADD.TOOL.

Position&nbsp;&nbsp;&nbsp;&nbsp;specifies the position of the button
within the toolbar. Position starts with 1 at the left side (if
horizontal) or at the top (if vertical).

**Related Functions**

ADD.TOOL&nbsp;&nbsp;&nbsp;Adds one or more buttons to a toolbar

GET.TOOL&nbsp;&nbsp;&nbsp;Returns information about a button or buttons
on a toolbar

PASTE.TOOL&nbsp;&nbsp;&nbsp;Pastes a button face from the Clipboard to a
specified position on a toolbar

Return to [top](#A)

# CREATE.NAMES

Equivalent to clicking the Create command on the Name submenu of the
Insert menu. Use CREATE.NAMES to quickly create names from text labels
on a sheet.

Arguments are logical values corresponding to check boxes in the Create
Names dialog box. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE or omitted, Microsoft Excel clears the check box.

**Syntax**

**CREATE.NAMES**(top, left, bottom, right)

**CREATE.NAMES**?(top, left, bottom, right)

Top&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Top Row check box.

Left&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Left Column check box.

Bottom&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Bottom Row check box.

Right&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Right Column check box.

**Remarks**

The cell containing the label text that Microsoft Excel uses to create
the names is not included in the resulting named range.

**Related Functions**

APPLY.NAMES&nbsp;&nbsp;&nbsp;Replaces references and values with their
corresponding names

DEFINE.NAME&nbsp;&nbsp;&nbsp;Defines a name on the active sheet or macro
sheet

DELETE.NAME&nbsp;&nbsp;&nbsp;Deletes a name

FORMULA.GOTO&nbsp;&nbsp;&nbsp;Selects a named area or reference on any
open workbook

Return to [top](#A)

# CREATE.OBJECT

Draws an object on a sheet or macro sheet and returns a value
identifying the object created. It is generally easier to use the macro
recorder to enter this function on your macro sheet.

**Syntax 1**

Lines, rectangles, ovals, arcs, pictures, text boxes, and buttons

**CREATE.OBJECT**(**obj\_type**, **ref1**, x\_offset1, y\_offset1,
**ref2**, x\_offset2, y\_offset2, text, fill, editable)

**Syntax 2**

Polygons

**CREATE.OBJECT**(**obj\_type**, **ref1**, x\_offset1, y\_offset1,
**ref2**, x\_offset2,  
y\_offset2, **array**, fill)

**Syntax 3**

Embedded charts

**CREATE.OBJECT**(**obj\_type**, **ref1**, x\_offset1, y\_offset1,
**ref2**, x\_offset2,  
y\_offset2, xy\_series, fill, gallery\_num, type\_num, plot\_visible)

Obj\_type&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the type of
object to create.

|               |                                        |
| ------------- | -------------------------------------- |
| **Obj\_type** | **Object**                             |
| 1             | Line                                   |
| 2             | Rectangle                              |
| 3             | Oval                                   |
| 4             | Arc                                    |
| 5             | Embedded chart                         |
| 6             | Text box                               |
| 7             | Button                                 |
| 8             | Picture (created with the camera tool) |
| 9             | Closed polygon                         |
| 10            | Open polygon                           |
| 11            | Check box                              |
| 12            | Option button                          |
| 13            | Edit box                               |
| 14            | Label                                  |
| 15            | Dialog frame                           |
| 16            | Spinner                                |
| 17            | Scroll bar                             |
| 18            | List box                               |
| 19            | Group box                              |
| 20            | Drop down list box                     |

Ref1&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cell from which the
upper-left corner of the object is drawn, or from which the upper-left
corner of the object's bounding rectangle is defined.

X\_offset1&nbsp;&nbsp;&nbsp;&nbsp;is the horizontal distance from the
upper-left corner of ref1 to the upper-left corner of the object or to
the upper-left corner of the object's bounding rectangle. X\_offset1 is
measured in points. A point is 1/72nd of an inch. If x\_offset1 is
omitted, it is assumed to be 0.

Y\_offset1&nbsp;&nbsp;&nbsp;&nbsp;is the vertical distance from the
upper-left corner of ref1 to the upper-left corner of the object or to
the upper-left corner of the object's bounding rectangle. Y\_offset1 is
measured in points. If y\_offset1 is omitted, it is assumed to be 0.

Ref2&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cell from which the
lower-right corner of the object is drawn, or from which the lower-right
corner of the object's bounding rectangle is defined.

X\_offset2&nbsp;&nbsp;&nbsp;&nbsp;is the horizontal distance from the
upper-left corner of ref2 to the lower-right corner of the object or to
the lower-right corner of the object's bounding rectangle. X\_offset2 is
measured in points. If x\_offset2 is omitted, it is assumed to be 0.

Y\_offset2&nbsp;&nbsp;&nbsp;&nbsp;is the vertical distance from the
upper-left corner of ref2 to the lower-right corner of the object or to
the lower-right corner of the object's bounding rectangle. Y\_offset2 is
measured in points. If y\_offset2 is omitted, it is assumed to be 0.

Text&nbsp;&nbsp;&nbsp;&nbsp;specifies the text that appears in a text
box or button. If text is omitted for a button, the button is named
"Button n", where n is a number. If obj\_type is not 6 or 7, text is
ignored.

Fill&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether the
object is filled or transparent. If fill is TRUE, the object is filled;
if FALSE, the object is transparent; if omitted, the object is filled
with an applicable pattern for the object being created.

Array&nbsp;&nbsp;&nbsp;&nbsp;is an n by 2 array of values, or a
reference to a range of cells containing values, that indicate the
position of each vertex in a polygon, relative to the upper-left corner
of the polygon's bounding rectangle.

  - > A vertex is a point that is defined by a pair of coordinates in
    > one row of array.

  - > If the polygon contains many vertices, one array may not be
    > sufficient to define it. If the number of characters in the
    > formula exceeds 1024, you must include one or more EXTEND.POLYGON
    > functions. If you're recording a macro, Microsoft Excel
    > automatically records EXTEND.POLYGON functions as needed. For more
    > information, see EXTEND.POLYGON.

> &nbsp;

Xy\_series&nbsp;&nbsp;&nbsp;&nbsp;is a number from 0 to 3 that specifies
how data is arranged in a chart and corresponds to options in the Paste
Special dialog box.

|                |                                                                                    |
| -------------- | ---------------------------------------------------------------------------------- |
| **Xy\_series** | **Result**                                                                         |
| 0              | Displays a dialog box if the selection is ambiguous                                |
| 1 or omitted   | First row/column is the first data series                                          |
| 2              | First row/column contains the category (x) axis labels                             |
| 3              | First row/column contains the x-values; the created chart is an xy (scatter) chart |

&nbsp;

  - > Xy\_series is ignored unless obj\_type is 5 (chart).

  - > If you want more control over how the data is arranged, use the
    > plot\_by, categories, and ser\_titles arguments to the
    > CHART.WIZARD function. For more information, see CHART.WIZARD.

> &nbsp;

Gallery\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 15 specifying
the type of embedded chart you want to create.

|                  |              |
| ---------------- | ------------ |
| **Gallery\_num** | **Chart**    |
| 1                | Area         |
| 2                | Bar          |
| 3                | Column       |
| 4                | Line         |
| 5                | Pie          |
| 6                | Radar        |
| 7                | XY (scatter) |
| 8                | Combination  |
| 9                | 3-D area     |
| 10               | 3-D bar      |
| 11               | 3-D column   |
| 12               | 3-D line     |
| 13               | 3-D pie      |
| 14               | 3-D surface  |
| 15               | Doughnut     |

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number identifying a formatting
option for a chart. The formatting options are shown in the dialog box
of the AutoFormat command that corresponds to the type of chart you're
creating. The first formatting option in any gallery is 1.

Plot\_visible&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds
to the Plot Visible Cells Only checkbox in the Chart tab of the Options
dialog box. If FALSE or omitted, all values are plotted.

Editable&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that determines
whether the drop down list box is editable or not. If TRUE, the drop
down list box is editable. If FALSE, the drop down list box is not
editable. If obj\_type is not 20, this argument is ignored.

**Remarks**

  - > CREATE.OBJECT returns the object identifier of the object it
    > created. Object identifiers include text describing the object,
    > such as "Text" or "Oval", and a number indicating the order in
    > which the object was created. For example, CREATE.OBJECT returns
    > "Oval 3" after creating an oval that is the third object in the
    > workbook.

  - > If the offsets are not specified, the object is drawn from the
    > upper-left corner of ref1 to the upper-left corner of ref2.

  - > If the object is not a picture and either ref1 or ref2 is omitted,
    > CREATE.OBJECT returns the \#VALUE\! error value and does not
    > create the object.

  - > CREATE.OBJECT also selects the object.

  - > You must use the COPY function before the CREATE.OBJECT function
    > to create a chart or a picture.

> &nbsp;

**Tip**&nbsp;&nbsp;&nbsp;To assign a macro to an object, use the
ASSIGN.TO.OBJECT function immediately after creating the object.

**Related Functions**

ASSIGN.TO.OBJECT&nbsp;&nbsp;&nbsp;Assigns a macro to an object

EXTEND.POLYGON&nbsp;&nbsp;&nbsp;Adds vertices to a polygon

FORMAT.MOVE&nbsp;&nbsp;&nbsp;Moves the selected object

FORMAT.SHAPE&nbsp;&nbsp;&nbsp;Inserts, moves, or deletes vertices of the
selected polygon

FORMAT.SIZE&nbsp;&nbsp;&nbsp;Sizes an object

GET.OBJECT&nbsp;&nbsp;&nbsp;Returns information about an object

OBJECT.PROPERTIES&nbsp;&nbsp;&nbsp;Determines an object's relationship
to underlying cells

TEXT.BOX&nbsp;&nbsp;&nbsp;Replaces text in a text box

Return to [top](#A)

# CREATE.PUBLISHER

Equivalent to clicking the Create Publisher command on the Publishing
submenu of the Edit menu. Publishes the selected range or chart to an
edition file for use by other Macintosh applications.

**Important**&nbsp;&nbsp;&nbsp;This function is only available if you
are using Microsoft Excel for the Macintosh with system software version
7.0 or later.

**Syntax**

**CREATE.PUBLISHER**(file\_text, appearance, size, formats)

**CREATE.PUBLISHER**?(file\_text, appearance, size, formats)

File\_text&nbsp;&nbsp;&nbsp;&nbsp;is a text string to be used as the
name of the new file that will contain the selected data. If file\_text
is omitted, Microsoft Excel uses the format "\<WorkbookName\> Edition
\#n", where WorkbookName is the name of the workbook from which the
publisher is being created, Edition indicates that the file is an
edition file, and n is a unique integer.

For example, if you omit file\_text and are publishing a selection from
a workbook named Seasonal, and it is your third publisher from that
workbook in the current work session, the default name of the publisher
would be "Seasonal Edition \#3".

Appearance&nbsp;&nbsp;&nbsp;&nbsp;specifies whether the selection is to
be published as shown on screen or as shown when printed. The default
value for appearance is 1 if the selection is a sheet and 2 if the
selection is a chart.

|                |                            |
| -------------- | -------------------------- |
| **Appearance** | **Selection is published** |
| 1              | As shown on screen         |
| 2              | As shown when printed      |

Size&nbsp;&nbsp;&nbsp;&nbsp;specifies the size at which to publish a
chart. Size is only available if a chart is to be published.

|              |                        |
| ------------ | ---------------------- |
| **Size**     | **Chart is published** |
| 1 or omitted | As shown on screen     |
| 2            | As shown when printed  |

Formats&nbsp;&nbsp;&nbsp;&nbsp;is number specifying what file format or
formats CREATE.PUBLISHER should use when it creates the Edition file.

|             |                 |
| ----------- | --------------- |
| **Formats** | **File format** |
| 1           | PICT            |
| 2           | BIFF            |
| 4           | RTF             |
| 8           | VALU            |

&nbsp;

  - > You can also use the sum of the allowable file formats for
    > formats. For example, a value of 6 specifies BIFF and RTF.

  - > If formats is omitted and the document is a sheet, formats is
    > assumed to be 15 (all formats); if the document is a chart,
    > formats is assumed to be 1 (PICT).

> &nbsp;

**Related Functions**

EDITION.OPTIONS&nbsp;&nbsp;&nbsp;Sets publisher and subscriber options

GET.LINK.INFO&nbsp;&nbsp;&nbsp;Returns information about a link

SUBSCRIBE.TO&nbsp;&nbsp;&nbsp;Inserts contents of an edition into the
active workbook

UPDATE.LINK&nbsp;&nbsp;&nbsp;Updates a link to another workbook

Return to [top](#A)

# CUSTOMIZE.TOOLBAR

Equivalent to choosing the Toolbars command from the View menu and
choosing the Customize button in Microsoft Excel 95. Displays the
Customize Toolbars dialog box. In Microsoft Excel 97 or later, this
function displays the Commands tab on the Customize dialog box. The
Customize dialog box appears when you choose the Toolbars command from
the View menu and then choose the Customize command. This function has a
dialog-box syntax only.

**Syntax**

**CUSTOMIZE.TOOLBAR**?(category)

Category&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies which
category of tools you want displayed in the dialog box. If omitted, the
previous setting is used. This argument is for compatibility with
Microsoft Excel 95.

|              |                       |
| ------------ | --------------------- |
| **Category** | **Category of tools** |
| 1            | File                  |
| 2            | Edit                  |
| 3            | Formula               |
| 4            | Formatting            |
| 5            | Text Formatting       |
| 6            | Drawing               |
| 7            | Macro                 |
| 8            | Charting              |
| 9            | Utility               |
| 10           | Data                  |
| 11           | TipWizard             |
| 12           | Auditing              |
| 13           | Forms                 |
| 14           | Custom                |

**Related Functions**

ADD.TOOLBAR&nbsp;&nbsp;&nbsp;Creates a new toolbar with the specified
tools

SHOW.TOOLBAR&nbsp;&nbsp;&nbsp;Hides or displays a toolbar

Return to [top](#A)

# CUSTOM.REPEAT

Allows custom commands to be repeated using the Repeat tool or the
Repeat command on the Edit menu. Also allows custom commands to be
recorded using the macro recorder.

**Syntax**

**CUSTOM.REPEAT**(macro\_text, repeat\_text, record\_text)

Macro\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of, or a reference to,
the macro you want to run when the Repeat command is chosen. If
macro\_text is omitted, no repeat macro is run, but the custom command
can still be recorded.

Repeat\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text you want to use as the
repeat command on the Edit menu (for example, "Repeat Reports"). You can
omit repeat\_text and macro\_text if you only want to record the formula
specified by record\_text when using the macro recorder.

Record\_text&nbsp;&nbsp;&nbsp;&nbsp;is the formula you want to record.
For example, if the user clicks a command named Run Reports in Macro 1,
the record\_text argument would be "=Macro1\!RunReports()", where
RunReports is the name of the macro called by the Run Reports command.

  - > References in record\_text must be in R1C1 format.

  - > If record\_text is omitted, the macro recorder records normally (a
    > RUN function with the first cell of the macro as its argument).

  - > If you are not recording a macro, record\_text is ignored.

> &nbsp;

**Tip**&nbsp;&nbsp;&nbsp;Place CUSTOM.REPEAT at the end of the macro you
will want to repeat. If you place it before the end, then the macro
formulas that follow CUSTOM.REPEAT may interfere with the desired
effects of CUSTOM.REPEAT. The Repeat tool and the Repeat command
continue to change as you click subsequent commands that can be
repeated.

**Example**

The following macro formula specifies that the macro RepeatReport on the
MenuMacros macro sheet in the current workbook will be run when the
Repeat Report command is chosen:

CUSTOM.REPEAT("MenuMacros\!RepeatReport", "Repeat Report")

**Related Function**

CUSTOM.UNDO&nbsp;&nbsp;&nbsp;Specifies a macro to run to undo a custom
command

Return to [top](#A)

# CUSTOM.UNDO

Creates a customized Undo tool and Undo or Redo command on the Edit menu
for custom commands.

**Syntax**

**CUSTOM.UNDO**(**macro\_text**, undo\_text)

Macro\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of, or an R1C1-style
reference to, the macro you want to run when the Undo command is chosen.
Macro\_text can be the name or cell reference of a macro.

Undo\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text you want to use as the
Undo command.

**Example**

The following macro function runs the UndoMult macro when the user
clicks the Undo Times100 command, a custom command that multiples the
current cell by 100.

\=CUSTOM.UNDO("UndoMult", "\&Undo Times100")

**Tip**&nbsp;&nbsp;&nbsp;Use CUSTOM.UNDO directly after the macro
functions you want to be able to repeat, because other macro functions
following CUSTOM.UNDO might reset the Undo command.

**Related Function**

CUSTOM.REPEAT&nbsp;&nbsp;&nbsp;Specifies a macro to run to repeat a
custom command

Return to [top](#A)

# CUT

Equivalent to choosing the Cut command from the Edit menu. Cuts or moves
data or objects.

**Syntax**

**CUT**(from\_reference, to\_reference)

From\_reference&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cell or
range of cells you want to cut. If from\_reference is omitted, it is
assumed to be the current selection.

To\_reference&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cell or range
of cells where you want to paste what you have cut.

  - > To\_reference should be a single cell or an enlarged multiple of
    > from\_reference. For example, if from\_reference is a 2 by 4
    > rectangle, to\_reference can be a 4 by 8 rectangle.

  - > To\_reference can be omitted so that you can paste from\_reference
    > later using the PASTE or PASTE.SPECIAL functions.

> &nbsp;

**Remarks**

The following information may be helpful if you're having problems with
CUT updating references in unexpected ways. When you move cells using
CUT, formulas that referred to from\_reference will refer to
to\_reference, and formulas that referred to to\_reference may return
\#REF\! error values. However, if from\_reference or to\_reference
contains references that are calculated at runtime (for example,
CUT(ACTIVE.CELL(), \!B1)), then Microsoft Excel does not update those
references when the CUT function is run, so no error values are
returned.

**Related Functions**

COPY&nbsp;&nbsp;&nbsp;Copies and pastes data or objects

PASTE&nbsp;&nbsp;&nbsp;Pastes cut or copied data

Return to [top](#A)

# DATA.DELETE

Equivalent to clicking the Delete command on the Data menu in Microsoft
Excel version 4.0. Deletes data that matches the current criteria in the
current database.

In the dialog-box form, DATA.DELETE?, Microsoft Excel displays a message
warning you that matching records will be permanently deleted, and you
can approve or cancel. In the plain form, DATA.DELETE, matching records
are deleted without any message being displayed.

**Syntax**

**DATA.DELETE**( )

**DATA.DELETE**?( )

Return to [top](#A)

# DATA.FIND

Equivalent to clicking the Find and Exit Find commands on the Data menu
in Microsoft Excel version 4.0. Selects records in the database range
which match criteria in the criteria range.

**Syntax**

**DATA.FIND**(**logical**)

Logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies whether
to enter or exit the Data Find mode. If logical is TRUE, Microsoft Excel
carries out the Find command; if FALSE, Microsoft Excel carries out the
Exit Find command. If logical is omitted, the function toggles between
Find and Exit Find.

**Related Functions**

DATA.FIND.NEXT&nbsp;&nbsp;&nbsp;Finds next matching record in a database

DATA.FIND.PREV&nbsp;&nbsp;&nbsp;Finds previous matching record in a
database

Return to [top](#A)

# DATA.FIND.NEXT, DATA.FIND.PREV

Equivalent to pressing the DOWN ARROW or UP ARROW key after the Find
command has been chosen from the Data menu in Microsoft Excel version
4.0. Finds the next or previous matching record in a database. If the
function cannot find a matching record, it returns the logical value
FALSE.

**Syntax**

**DATA.FIND.NEXT**( )

**DATA.FIND.PREV**( )

**Related Function**

DATA.FIND&nbsp;&nbsp;&nbsp;Enters or exits Data Find mode

Return to [top](#A)

# DATA.FORM

Equivalent to clicking the Form command on the Data menu. Displays the
data form.

If Microsoft Excel cannot determine what database or list of information
to use, the function returns the \#VALUE\! error value and interrupts
the macro.

**Syntax**

**DATA.FORM**( )

**Remarks**

  - > You can still use custom data forms created in Microsoft Excel
    > version 4.0 or earlier. To edit the definition table of the custom
    > data form, use the Dialog Editor from Microsoft Excel version 4.0.

  - > The data form can handle up to 32 fields.

> &nbsp;

Return to [top](#A)

# DATA.LABEL

Specifies label contents and position.

**Syntax**

**DATA.LABEL**(show\_option, auto\_text, show\_key)

Show\_option&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies what type
of labels to display.

|                  |                        |
| ---------------- | ---------------------- |
| **Show\_option** | **Type displayed**     |
| 1                | none                   |
| 2                | Show value             |
| 3                | Show percent           |
| 4                | Show label             |
| 5                | Show label and percent |

Auto\_text&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds
the Automatic Checkbox in the Data Labels dialog box. If TRUE, resets a
chart's data labels back to their actual values. If FALSE, they are not
reset. The Automatic Text checkbox appears only if the label has been
selected and its value changed.

Show\_key&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specified
whether to show the legend key next to the label. If TRUE, displays the
legend key. If FALSE or omitted, does not display the legend key.

Return to [top](#A)

# DATA.SERIES

Equivalent to clicking the Series command on the Fill submenu of the
Edit menu. Use DATA.SERIES to enter an interpolated or incrementally
increasing or decreasing series of numbers or dates on a sheet or macro
sheet.

**Syntax**

**DATA.SERIES**(rowcol, type\_num, date\_num, step\_value, stop\_value,
trend)

**DATA.SERIES**?(rowcol, type\_num, date\_num, step\_value, stop\_value,
trend)

Rowcol&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies where the
series should be entered. If rowcol is omitted, the default value is
based on the size and shape of the current selection.

|            |                     |
| ---------- | ------------------- |
| **Rowcol** | **Enter series in** |
| 1          | Rows                |
| 2          | Columns             |

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 that specifies
the type of series.

|               |                    |
| ------------- | ------------------ |
| **Type\_num** | **Type of series** |
| 1 or omitted  | Linear             |
| 2             | Growth             |
| 3             | Date               |
| 4             | AutoFill           |

Date\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 that specifies
the date unit of the series, as shown in the following table. To use the
date\_num argument, the type\_num argument must be 3.

|               |               |
| ------------- | ------------- |
| **Date\_num** | **Date unit** |
| 1 or omitted  | Day           |
| 2             | Weekday       |
| 3             | Month         |
| 4             | Year          |

Step\_value&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies the step
value for the series. If step\_value is omitted, it is assumed to be 1.

Stop\_value&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies the stop
value for the series. If stop\_value is omitted, DATA.SERIES continues
filling the series until the end of the selected range.

Trend&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
Trend check box. If trend is TRUE, Microsoft Excel generates a linear or
exponential trend; if FALSE or omitted, Microsoft Excel generates a
standard data series.

**Remarks**

  - > If you specify a positive value for stop\_value that is lower than
    > the value in the active cell of the selection, DATA.SERIES takes
    > no action.

  - > If type\_num is 4 (AutoFill), Microsoft Excel performs an AutoFill
    > operation just as if you had filled the selection by dragging the
    > fill selection handle or had used the FILL.AUTO macro function.

> &nbsp;

**Related Function**

FILL.AUTO&nbsp;&nbsp;&nbsp;Copies cells or automatically fills a
selection

Return to [top](#A)

# DEFINE.NAME

Equivalent to clicking the Define command on the Name submenu of the
Insert menu. Defines a name on the active sheet or macro sheet. Use
DEFINE.NAME instead of SET.NAME when you want to define a name on the
active sheet.

**Syntax**

**DEFINE.NAME**(**name\_text**, refers\_to, macro\_type, shortcut\_text,
hidden, category, local)

**DEFINE.NAME**?(name\_text, refers\_to, macro\_type, shortcut\_text,
hidden, category, local)

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text you want to use as the
name. Names cannot include spaces, and cannot look like cell references.

Refers\_to&nbsp;&nbsp;&nbsp;&nbsp;describes what name\_text should refer
to, and can be any of the following values.

<table>
<tbody>
<tr class="odd">
<td><strong>If refers_to is</strong></td>
<td><strong>Then name_text is</strong></td>
</tr>
<tr class="even">
<td>A number, text, or logical value</td>
<td>Defined to refer to that value</td>
</tr>
<tr class="odd">
<td>An external reference, such as !$A$1 or SALES!$A$1:$C$3</td>
<td>Defined to refer to those cells</td>
</tr>
<tr class="even">
<td>A formula in the form of text, such as "=2*PI()/360" (if the formula contains references, they must be R1C1-style references, such as<br />
"=R2C2*(1+RC[-1])")</td>
<td>Defined to refer to that formula</td>
</tr>
<tr class="odd">
<td>Omitted</td>
<td>Defined to refer to the current selection</td>
</tr>
</tbody>
</table>

The next two arguments, macro\_type and shortcut\_text, apply only if
the sheet in the active window is a macro sheet.

Macro\_type&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 that
indicates the type of macro.

|                 |                                                      |
| --------------- | ---------------------------------------------------- |
| **Macro\_type** | **Type of macro**                                    |
| 1               | Custom function (also known as a function macro)     |
| 2               | Command macro.                                       |
| 3 or omitted    | None (that is, name\_text does not refer to a macro) |

Shortcut\_text&nbsp;&nbsp;&nbsp;&nbsp;is a text value that specifies the
macro shortcut key. Shortcut\_text must be a single letter, such as "z"
or "Z".

Hidden&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether to
define the name as a hidden name. If hidden is TRUE, Microsoft Excel
defines the name as a hidden name; if FALSE or omitted, Microsoft Excel
defines the name normally.

Category&nbsp;&nbsp;&nbsp;&nbsp;is a number or text identifying the
category of a custom function and corresponds to categories in the
Function Category list box.

  - > Categories are numbered starting with 1, the first category in the
    > list.

  - > If category is text but is not one of the existing function types,
    > Microsoft Excel creates a new category and assigns your custom
    > function to it.

Local&nbsp;&nbsp;&nbsp;&nbsp;is a logical value which, if TRUE, defines
the name on just the current sheet or macro sheet. If FALSE or omitted,
defines the name for all sheets in the workbook.

**Remarks**

  - > You can use hidden names to define values that you want to prevent
    > the user from seeing or changing; they do not appear in the Define
    > Name, Paste Name, or Goto dialog boxes. Hidden names can only be
    > created with the DEFINE.NAME macro function.

  - > If you are recording a macro and you define a name to refer to a
    > formula, Microsoft Excel converts A1-style references to
    > R1C1-style references. For example, if the active cell is C2, and
    > you define the name Previous to refer to =B2, Microsoft Excel
    > records that command as DEFINE.NAME("Previous","=RC\[-1\]").

  - > In DEFINE.NAME?, the dialog-box form of the function, if
    > refers\_to is not specified, the current selection is proposed in
    > the Refers To box. Also, if a name is not specified, text in the
    > active cell is proposed as the name.

> &nbsp;

**Related Functions**

DELETE.NAME&nbsp;&nbsp;&nbsp;Deletes a name

GET.DEF&nbsp;&nbsp;&nbsp;Returns a name matching a definition

GET.NAME&nbsp;&nbsp;&nbsp;Returns the definition of a name

NAMES&nbsp;&nbsp;&nbsp;Returns the names defined in a workbook

SET.NAME&nbsp;&nbsp;&nbsp;Defines a name as a value

Return to [top](#A)

# DEFINE.STYLE

Equivalent to clicking the Define button in the Style dialog box, which
appears when you click the Style command on the Format menu. Creates and
changes cell styles. There are seven syntax forms of this function. Use
syntax 1 of DEFINE.STYLE to define styles based on the format of the
active cell. To create a style by specifying number, font, and other
formats, use syntaxes 2 through 7 of DEFINE.STYLE.

Syntax 1

Syntaxes 2-7

Return to [top](#A)

# DEFINE.STYLE Syntax 1

Equivalent to clicking the Define button in the Style dialog box, which
appears when you click the Style command on the Format menu. Creates and
changes cell styles. There are seven syntax forms of this function. Use
syntax 1 of DEFINE.STYLE to define styles based on the format of the
active cell. To create a style by specifying number, font, and other
formats, use syntaxes 2 through 7 of DEFINE.STYLE.

**Syntax**

**DEFINE.STYLE**(**style\_text**, number, font, alignment, border,
pattern, protection)

**DEFINE.STYLE**?(style\_text, number, font, alignment, border, pattern,
protection)

Style\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name, as text, that you want
to assign to the style.

The following arguments are logical values corresponding to check boxes
in the Style dialog box. If an argument is TRUE, Microsoft Excel selects
the check box and uses the corresponding format of the active cell in
the style; if FALSE, Microsoft Excel clears the check box and omits
formatting descriptions for that attribute. If style\_text is omitted
and all selected cells have identical formatting, the default is TRUE;
if cells have different formatting, the default is FALSE.

Number&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Number check box.

Font&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Font check box.

Alignment&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Alignment check box.

Border&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Border check box.

Pattern&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Pattern check box.

Protection&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Protection check
box.

**Related Functions**

DEFINE.STYLE Syntaxes 2-7

APPLY.STYLE&nbsp;&nbsp;&nbsp;Applies a style to the selection

DELETE.STYLE&nbsp;&nbsp;&nbsp;Deletes a cell style

MERGE.STYLES&nbsp;&nbsp;&nbsp;Imports styles from another workbook into
the active workbook

Return to [top](#A)

# DEFINE.STYLE Syntaxes 2 - 7

Equivalent to clicking the Define button in the Style dialog box, which
appears when you click the Style command on the Format menu. Creates and
changes cell styles. Use one of the following syntax forms of
DEFINE.STYLE to select cell formats for a new style or to alter the
formats of an existing style. Use syntax 1 of DEFINE.STYLE to define
styles based on the format of the active cell.

**Syntax 2**

Number format, using the arguments from the FORMAT.NUMBER function

**DEFINE.STYLE**(**style\_text, attribute\_num**, format\_text)

**Syntax 3**

Font format, using the arguments from the FORMAT.FONT and
FONT.PROPERTIES functions

**DEFINE.STYLE**(**style\_text, attribute\_num**, name\_text, size\_num,
bold, italic, underline, strike, color, outline, shadow, superscript,
subscript)

**Syntax 4**

Alignment, using the arguments from the ALIGNMENT function

**DEFINE.STYLE**(**style\_text**, **attribute\_num**, horiz\_align,
wrap, vert\_align, orientation)

**Syntax 5**

Border, using the arguments from the BORDER function

**DEFINE.STYLE**(**style\_text**, **attribute\_num**, left, right, top,
bottom, left\_color, right\_color, top\_color, bottom\_color)

**Syntax 6**

Pattern, using the arguments from the cell form of the PATTERNS function

**DEFINE.STYLE**(**style\_text, attribute\_num**, apattern, afore,
aback)

**Syntax 7**

Cell protection, using the arguments from the CELL.PROTECTION function

**DEFINE.STYLE**(**style\_text, attribute\_num**, locked, hidden)

Style\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name, as text, that you want
to assign to the style.

Attribute\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 2 to 7 that
specifies which attribute of the style, such as its font, alignment, or
number format, you want to designate with this function.

|                    |                 |
| ------------------ | --------------- |
| **Attribute\_num** | **Specifies**   |
| 2                  | Number format   |
| 3                  | Font format     |
| 4                  | Alignment       |
| 5                  | Border          |
| 6                  | Pattern         |
| 7                  | Cell protection |

**Remarks**

  - > The remaining arguments are different for each form and are
    > identical to arguments in the corresponding function. For example,
    > form 2 of DEFINE.STYLE defines the number format of a style and
    > corresponds to the FORMAT.NUMBER function. The exception is form
    > 5, which does not include every argument for BORDER. For details
    > on the values you can use for these arguments, see the description
    > under the corresponding function.

  - > If you define a style using one of these forms, then any
    > attributes you don't explicitly define are not changed.

> &nbsp;

**Related Functions**

DEFINE.STYLE Syntax 1

ALIGNMENT&nbsp;&nbsp;&nbsp;Aligns or wraps text in cells

APPLY.STYLE&nbsp;&nbsp;&nbsp;Applies a style to the selection

BORDER&nbsp;&nbsp;&nbsp;Adds a border to the selected cell or object

CELL.PROTECTION&nbsp;&nbsp;&nbsp;Allows you to control cell protection
and display

DELETE.STYLE&nbsp;&nbsp;&nbsp;Deletes a cell style

FONT.PROPERTIES&nbsp;&nbsp;&nbsp;Applies a font to the selection

FORMAT.NUMBER&nbsp;&nbsp;&nbsp;Formats numbers, dates, and times in the
selected cells

MERGE.STYLES&nbsp;&nbsp;&nbsp;Imports styles from another workbook into
the active workbook

PATTERNS&nbsp;&nbsp;&nbsp;Changes the appearance of the selected object

Return to [top](#A)

# DELETE.ARROW

Deletes the selected arrow, either drawn as an arrow with the arrow tool
or as a line that is later formatted as an arrow. In Microsoft Excel
version 5.0 or later, arrows are named lines.

**Syntax**

**DELETE.ARROW**( )

If the selection is not an arrow or a line formatted as an arrow, or if
the active window is not a chart, DELETE.ARROW interrupts the macro.

**Tip**&nbsp;&nbsp;&nbsp;Use the SELECT function (chart syntax), with
the number of the arrow (or line) you want to delete in order to select
the arrow before using the DELETE.ARROW function. For example, SELECT
("Line 1"). You can also use the CLEAR function to delete the arrow.

**Related Functions**

CLEAR&nbsp;&nbsp;&nbsp;Clears specified information from the selected
cells or chart

DELETE.OVERLAY&nbsp;&nbsp;&nbsp;Deletes the overlay on a chart

Return to [top](#A)

# DELETE.BAR

Deletes a custom menu bar.

**Syntax**

**DELETE.BAR**(**bar\_num**)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the ID number of the custom menu bar
you want to delete.

**Tip**&nbsp;&nbsp;&nbsp;Rather than trying to discover the ID number of
the menu bar you want to delete, use a reference to the ADD.BAR function
that created the bar. For example, the following macro formula deletes
the menu bar created by the ADD.BAR function in the cell named
ReportsBar:

DELETE.BAR(ReportsBar)

**Related Functions**

ADD.BAR&nbsp;&nbsp;&nbsp;Adds a menu bar

SHOW.BAR&nbsp;&nbsp;&nbsp;Displays a menu bar

Return to [top](#A)

# DELETE.CHART.AUTOFORMAT

Deletes a custom format from the list of formats shown in the Custom
Types tab in the Chart Type dialog box.

**Syntax**

**DELETE.CHART.AUTOFORMAT**(**name\_text**)

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the template name you want to
delete from the list of custom templates.

**Related Function**

ADD.CHART.AUTOFORMAT&nbsp;&nbsp;&nbsp;Adds a custom template

Return to [top](#A)

# DELETE.COMMAND

Deletes a command from a custom or built-in menu. Use DELETE.COMMAND to
remove commands you don't want the user to have access to or to remove
custom commands that you have added.

**Syntax**

**DELETE.COMMAND**(**bar\_num, menu, command**, subcommand)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the menu bar from which you want to
delete the command. Bar\_num can be the ID number of a built-in or
custom menu bar. See ADD.COMMAND for a list of ID numbers for built-in
menu bars and shortcut menus.

Menu&nbsp;&nbsp;&nbsp;&nbsp;is the menu from which you want to delete
the command. Menu can be the name of a menu as text or the number of a
menu. Menus are numbered starting with 1 from the left of the screen.

Command&nbsp;&nbsp;&nbsp;&nbsp;is the command you want to delete, or the
name of a submenu. Command can be the name of the command as text or the
number of the command; the first command on a menu is in position 1.

Subcommand&nbsp;&nbsp;&nbsp;&nbsp;is the command you want to delete from
a submenu. If you use subcommand, you must use command as the name of
the submenu.

**Remarks**

  - > If the specified command does not exist, DELETE.COMMAND returns
    > the \#VALUE\! error value and interrupts the macro.

  - > After a command is deleted, the command number for all commands
    > below that command is decreased by one.

  - > When you delete a built-in command, DELETE.COMMAND returns a
    > unique ID number for that command. You can use this ID number with
    > ADD.COMMAND to restore the built-in command to the original menu.

> &nbsp;

**Example**

The following macro formula removes the Compile Reports command from the
Reports menu on a custom menu bar created by the ADD.BAR function in a
cell named Financials.

DELETE.COMMAND(Financials, "Reports", "Compile Reports...")

**Related Functions**

ADD.COMMAND&nbsp;&nbsp;&nbsp;Adds a command to a menu

CHECK.COMMAND&nbsp;&nbsp;&nbsp;Adds or deletes a check mark to or from a
command

ENABLE.COMMAND&nbsp;&nbsp;&nbsp;Enables or disables a menu or custom
command

RENAME.COMMAND&nbsp;&nbsp;&nbsp;Changes the name of a command or menu

Return to [top](#A)

# DELETE.FORMAT

Equivalent to deleting the specified format in the Number tab in the
Format Cells dialog box, which appears when you click the Cells command
on the Format menu, or in the Number tab for selected chart objects.
Deletes a specified built-in or custom number format.

**Syntax**

**DELETE.FORMAT**(**format\_text**)

Format\_text&nbsp;&nbsp;&nbsp;&nbsp;is the format given as a text
string, for example, "000-00-0000".

**Remarks**

When you delete a custom number format, all numbers formatted with that
number format are formatted with the General format.

**Related Functions**

FORMAT.NUMBER&nbsp;&nbsp;&nbsp;Applies a number format to the selection

GET.CELL&nbsp;&nbsp;&nbsp;Returns information about the specified cell

Return to [top](#A)

# DELETE.MENU

Deletes a menu or submenu. Use DELETE.MENU to delete menus you have
added to menu bars when the supporting macro sheet is closed (using an
Auto\_Close macro), or any time you want to remove a menu.

**Syntax**

**DELETE.MENU**(**bar\_num, menu**, submenu)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the menu bar from which you want to
delete the menu. Bar\_num can be the number of a Microsoft Excel
built-in menu bar or the number returned by a previously run ADD.BAR
function. For a list of ID numbers for built-in menu bars, see
ADD.COMMAND.

Menu&nbsp;&nbsp;&nbsp;&nbsp;is the menu you want to delete. Menu can be
either the name of a menu as text or the number of a menu. Menus are
numbered starting with 1 from the left of the screen. If the specified
menu does not exist, DELETE.MENU returns the \#VALUE\! error value and
interrupts the macro. After a menu is deleted, the menu number for each
menu to the right of that menu is decreased by 1.

Submenu&nbsp;&nbsp;&nbsp;&nbsp;is the name of the submenu you want to
delete or the number of the menu in the list of commands.

**Remarks**

You cannot delete a shortcut menu. Instead, use ENABLE.COMMAND to
prevent the user from accessing a shortcut menu.

**Example**

The following macro formula deletes the Reports menu from the custom
menu bar created by the ADD.BAR function in a cell named Financials:

DELETE.MENU(Financials, "Reports")

**Related Functions**

ADD.MENU&nbsp;&nbsp;&nbsp;Adds a menu to a menu bar

ADD.BAR&nbsp;&nbsp;&nbsp;Adds a menu bar

DELETE.BAR&nbsp;&nbsp;&nbsp;Deletes a menu bar

DELETE.COMMAND&nbsp;&nbsp;&nbsp;Deletes a command from a menu

ENABLE.COMMAND&nbsp;&nbsp;&nbsp;Enables or disables a menu or custom
command

Return to [top](#A)

# DELETE.NAME

Equivalent to clicking the Delete button in the Define Name dialog box,
which appears when you click the Define command on the Name submenu of
the Insert menu. Deletes the specified name.

**Syntax**

**DELETE.NAME**(**name\_text**)

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is a text value specifying the name
that you want to delete.

**Important&nbsp;&nbsp;&nbsp;**Formulas that use names in their
arguments may return incorrect or error values when a name used in the
formula is deleted.

**Related Functions**

DEFINE.NAME&nbsp;&nbsp;&nbsp;Defines a name in the active workbook

GET.NAME&nbsp;&nbsp;&nbsp;Returns the definition of a name

SET.NAME&nbsp;&nbsp;&nbsp;Defines a name as a value

Return to [top](#A)

# DELETE.OVERLAY

Equivalent to clicking the Delete Overlay command on the Chart menu in
Microsoft Excel version 4.0. Deletes all overlays from a chart. If the
chart has no overlay, DELETE.OVERLAY takes no action and returns TRUE.

**Syntax**

**DELETE.OVERLAY**( )

Return to [top](#A)

# DELETE.STYLE

Equivalent to choosing the Delete button from the Style dialog box,
which appears when you choose the Style command from the Format menu.
Deletes a style from a workbook. Cells formatted with the deleted style
revert to the Normal style.

**Syntax**

**DELETE.STYLE**(style\_text)

Style\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a style to be deleted.
If style\_text does not exist, DELETE.STYLE returns the \#VALUE\! error
value and interrupts the macro.

**Remarks**

You can only delete styles from the active workbook. External references
are not permitted as part of the style\_text argument.

**Related Functions**

APPLY.STYLE&nbsp;&nbsp;&nbsp;Applies a style to the selection

DEFINE.STYLE&nbsp;&nbsp;&nbsp;Creates or changes a cell style

MERGE.STYLES&nbsp;&nbsp;&nbsp;Merges styles from another workbook into
the active workbook

Return to [top](#A)

# DELETE.TOOL

Equivalent to selecting a button and dragging it to an area other than a
toolbar. Deletes a button from a toolbar.

**Syntax**

**DELETE.TOOL**(**bar\_id, position**)

Bar\_id&nbsp;&nbsp;&nbsp;&nbsp;specifies the name or number of a toolbar
from which you want to delete a button. For detailed information about
bar\_id, see ADD.TOOL.

Position&nbsp;&nbsp;&nbsp;&nbsp;specifies the position of the button
within the toolbar. Position starts with 1 at the left side (if
horizontal) or at the top (if vertical).

**Related Functions**

ADD.TOOL&nbsp;&nbsp;&nbsp;Adds one or more buttons to a toolbar

ADD.TOOLBAR&nbsp;&nbsp;&nbsp;Creates a new toolbar with the specified
buttons

DELETE.TOOLBAR&nbsp;&nbsp;&nbsp;Deletes custom toolbars

Return to [top](#A)

# DELETE.TOOLBAR

Equivalent to clicking the Delete button in the Toolbars dialog box,
which appears when you click the Customize command (View menu, Toolbars
submenu). Deletes a custom toolbar.

**Syntax**

**DELETE.TOOLBAR**(**bar\_name**)

Bar\_name&nbsp;&nbsp;&nbsp;&nbsp;specifies the name of the toolbar that
you want to delete. For detailed information about bar\_name, see
ADD.TOOL.

**Remarks**

  - > You cannot delete built-in toolbars.

  - > If DELETE.TOOLBAR successfully deletes the toolbar, it returns
    > TRUE. If you try to delete a built-in toolbar, DELETE.TOOLBAR
    > returns the \#VALUE\! error value, interrupts the macro, and takes
    > no other action.

> &nbsp;

**Related Functions**

ADD.TOOL&nbsp;&nbsp;&nbsp;Adds or more buttons to a toolbar

ADD.TOOLBAR&nbsp;&nbsp;&nbsp;Creates a new toolbar with the specified
buttons

RESET.TOOLBAR&nbsp;&nbsp;&nbsp;Resets a built-in toolbar to its initial
default setting

Return to [top](#A)

# DEMOTE

Equivalent to clicking the Group tool. Demotes, or groups, the selected
rows or columns in an outline. Use DEMOTE to change the configuration of
an outline by grouping rows or columns of information.

**Syntax**

**DEMOTE**(**row\_col**)

**DEMOTE**?(row\_col)

Row\_col&nbsp;&nbsp;&nbsp;&nbsp;specifies whether to group rows or
columns.

|              |             |
| ------------ | ----------- |
| **Row\_col** | **Demotes** |
| 1 or omitted | Rows        |
| 2            | Columns     |

**Remarks**

  - > If the selection consists of an entire row or rows, then rows are
    > demoted even if row\_col is 2. Similarly, selection of an entire
    > column overrides row\_col 1.

  - > If the selection is unambiguous (an entire row or column), then
    > DEMOTE? will not display the dialog box.

> &nbsp;

**Related Functions**

PROMOTE&nbsp;&nbsp;&nbsp;Promotes the selection in an outline

SHOW.DETAIL&nbsp;&nbsp;&nbsp;Expands or collapses a portion of an
outline

SHOW.LEVELS&nbsp;&nbsp;&nbsp;Displays a specific number of levels of an
outline

Return to [top](#A)

# DEREF

Returns the value of the cells in a reference.

**Syntax**

**DEREF**(**reference**)

Reference&nbsp;&nbsp;&nbsp;&nbsp;is the cell or cells from which you
want to obtain a value. If reference is the reference of a single cell,
DEREF returns the value of that cell. If reference is the reference of a
range of cells, DEREF returns the array of values in those cells. If
reference refers to the active sheet, it must be an absolute reference.
Relative references are converted to absolute references.

**Remarks**

In most formulas, there is no difference between using a value and using
the reference of a cell containing that value. The reference is
automatically converted to the value, as necessary. For example, if cell
A1 contains the value 2, then the formula =A1+1, like the formula =2+1,
returns the result 3, because the reference A1 is converted to the value
2. However, in a few functions, such as the SET.NAME function,
references are not automatically converted to values. Instead, those
functions behave differently depending on whether an argument is a
reference or a value.

**Example**

See the sixth example for SET.NAME.

**Related Function**

SET.NAME&nbsp;&nbsp;&nbsp;Defines a names as a value

Return to [top](#A)

# DESCR

Generates descriptive statistics for data in the input range.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**DESCR**(**inprng**, outrng, grouped, labels, summary, ds\_large,
ds\_small, confid)

**DESCR**?(inprng, outrng, grouped, labels, summary, ds\_large,
ds\_small, confid)

Inprng&nbsp;&nbsp;&nbsp;&nbsp;is the input range.

Outrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell) in
the output table or the name, as text, of a new sheet to contain the
output table. If FALSE, blank, or omitted, places the output table in a
new workbook.

Grouped&nbsp;&nbsp;&nbsp;&nbsp;is a text character that indicates
whether the data in the input range is organized by row or column.

  - > If grouped is "C" or omitted, then the data is organized by
    > column.

  - > If grouped is "R" then the data is organized by row.

> &nbsp;

Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that describes where
the labels are located in the input range, as shown in the following
table:

|                  |             |                                                   |
| ---------------- | ----------- | ------------------------------------------------- |
| **Labels**       | **Grouped** | **Labels are in**                                 |
| TRUE             | "C"         | First row of the input range.                     |
| TRUE             | "R"         | First column of the input range.                  |
| FALSE or omitted | (ignored)   | No labels. All cells in the input range are data. |

Summary&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If TRUE, DESCR
reports the summary statistics. If FALSE or omitted, no summary
statistics are reported.

Ds\_large&nbsp;&nbsp;&nbsp;&nbsp;is an integer k. If ds\_large is
present, DESCR reports the k-th largest data point. If ds\_large is
omitted, the value is not reported.

Ds\_small&nbsp;&nbsp;&nbsp;&nbsp;is an integer k. If ds\_small is
present, DESCR reports the k-th smallest data point. If ds\_small is
omitted, the value is not reported.

Confid&nbsp;&nbsp;&nbsp;&nbsp;is the confidence level of the mean. If
confid is given, DESCR reports the confidence interval for the input
range. If confid is omitted, the confidence interval is 95%.

Return to [top](#A)

# DIALOG.BOX

Displays the dialog box described in a dialog box definition table.

**Syntax**

**DIALOG.BOX**(**dialog\_ref**)

Dialog\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to a dialog box
definition table on sheet, or an array containing the definition table.

  - > If an OK button in the dialog box is chosen, DIALOG.BOX enters
    > values in fields as specified in the dialog\_ref area and returns
    > the position number of the button chosen. The position numbers
    > start with 1 in the second row of the dialog box definition table.

  - > If the Cancel button in the dialog box is chosen, DIALOG.BOX
    > returns FALSE.

> &nbsp;

The dialog box definition table must be at least seven columns wide and
two rows high. The definitions of each column in a dialog box definition
table are listed in the following table.

|                         |                   |
| ----------------------- | ----------------- |
| **Column type**         | **Column number** |
| Item number             | 1                 |
| Horizontal position     | 2                 |
| Vertical position       | 3                 |
| Item width              | 4                 |
| Item height             | 5                 |
| Text                    | 6                 |
| Initial value or result | 7                 |

The first row of dialog\_ref defines the position, size, and name of the
dialog box. It can also specify the default selected item and the
reference for the Help button. The position is specified in columns 2
and 3, the size in columns 4 and 5, and the name in column 6. To specify
a default item, place the item's position number in column 7. You can
place the reference for the Help button in row 1, column 1 of the table,
but the preferred location is column 7 in the row where the Help button
is defined. Row 1, column 1 is usually left blank.

The following table lists the numbers for the items you can display in a
dialog box.

|                                               |                 |
| --------------------------------------------- | --------------- |
| **Dialog-box item**                           | **Item number** |
| Default OK button                             | 1               |
| Cancel button                                 | 2               |
| OK button                                     | 3               |
| Default Cancel button                         | 4               |
| Static text                                   | 5               |
| Text edit box                                 | 6               |
| Integer edit box                              | 7               |
| Number edit box                               | 8               |
| Formula edit box                              | 9               |
| Reference edit box                            | 10              |
| Option button group                           | 11              |
| Option button                                 | 12              |
| Check box                                     | 13              |
| Group box                                     | 14              |
| List box                                      | 15              |
| Linked list box                               | 16              |
| Icons                                         | 17              |
| Linked file list box (Windows only)           | 18              |
| Linked drive and directory box (Windows only) | 19              |
| Directory text box                            | 20              |
| Drop-down list box                            | 21              |
| Drop-down combination edit/list box           | 22              |
| Picture button                                | 23              |
| Help button                                   | 24              |

**Remarks**

  - > Add 100 to an item number in the above table to define the item as
    > a trigger. A trigger is a dialog box item that, when chosen,
    > returns to your macro (as clicking OK would) but continues to
    > display the dialog box, allowing your macro to change the dialog
    > box definition or display an alert message or another dialog box.
    > The Help button, edit boxes, group boxes, static text, and icons
    > cannot be triggers.

  - > Add 200 to an item number to define it as dimmed. A dimmed (gray)
    > item cannot be chosen or selected. For example, 203 is a dimmed OK
    > button. You can use item 223 to include a picture in your dialog
    > box that does not behave like a button.

  - > If a trigger has been chosen and you still want to clear a dynamic
    > dialog box from the screen, use DIALOG.BOX(FALSE). This is useful
    > if you want to confirm that the dialog box has been filled out
    > correctly before dismissing it.

  - > The dialog box definition table can be an array. If dialog\_ref is
    > an array instead of a reference, DIALOG.BOX returns a modified
    > copy of that array, along with the results of the dialog box in
    > the seventh column. (The first item in the seventh column is the
    > position number of the chosen button or of a triggered item.) This
    > is useful if you want to preserve the original dialog box
    > definition table since DIALOG.BOX does not modify the original
    > array argument. If you cancel the dialog box, or if a dialog box
    > error occurs, DIALOG.BOX returns FALSE instead of an array.

> &nbsp;

**Related Functions**

ALERT&nbsp;&nbsp;&nbsp;Displays a dialog box and a message

INPUT&nbsp;&nbsp;&nbsp;Displays a dialog box for user input

Return to [top](#A)

# DIRECTORY

Sets the current drive and directory or folder to the specified path and
returns the name of the new directory or folder as text. Use DIRECTORY
to get the name of the current directory or folder for use with the OPEN
and SAVE.AS functions or to specify a directory or folder from which to
return a list of files with the FILES function.

**Syntax**

**DIRECTORY**(**path\_text**)

Path\_text&nbsp;&nbsp;&nbsp;&nbsp;is the drive and directory or folder
you want to change to.

  - > If path\_text is not specified, DIRECTORY returns the name of the
    > current directory or folder as text.

  - > If path\_text does not specify a drive, the current drive is
    > assumed.

> &nbsp;

**Examples**

In Microsoft Excel for Windows, the following macro formula sets the
directory to \\EXCEL\\MODELS on the current drive and returns the value
"drive:\\EXCEL\\MODELS":

DIRECTORY("\\EXCEL\\MODELS")

The following macro formula sets the current drive to E and sets the
directory to \\EXCEL\\MODELS on E. It returns the value
"E:\\EXCEL\\MODELS":

DIRECTORY("E:\\EXCEL\\MODELS")

In Microsoft Excel for the Macintosh, the following macro formula sets
the folder to HARD DISK: APPS:EXCEL:FINANCIALS and returns the value
"HARD DISK:APPS:EXCEL:FINANCIALS":

DIRECTORY("HARD DISK:APPS:EXCEL:FINANCIALS")

**Related Function**

FILES&nbsp;&nbsp;&nbsp;Returns the filenames in the specified directory
or folder

Return to [top](#A)

# DISABLE.INPUT

Blocks all input from the keyboard and mouse to Microsoft Excel (except
input to displayed dialog boxes). Use DISABLE.INPUT to prevent input
from the user or from other applications.

**Syntax**

**DISABLE.INPUT**(**logical**)

Logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether
input is currently disabled. If logical is TRUE, input is disabled; if
FALSE, input is reenabled.

**Remarks**

Disabling input can be useful if you are using dynamic data exchange
(DDE) to communicate with Microsoft Excel from another application.

**Important**&nbsp;&nbsp;&nbsp;Be sure to end any macro that uses
DISABLE.INPUT(TRUE) with a DISABLE.INPUT(FALSE) function. If you do not
include DISABLE.INPUT(FALSE) to allow non-dialog-box input, you will not
be able to take any actions on your computer after the macro has
finished.

**Related Functions**

CANCEL.KEY&nbsp;&nbsp;&nbsp;Disables macro interruption

ENTER.DATA&nbsp;&nbsp;&nbsp;Turns Data Entry mode on and off

WORKSPACE&nbsp;&nbsp;&nbsp;Changes workspace settings

Return to [top](#A)

# DISPLAY

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. Use syntax 2 to
control the display of the Info Window.

Syntax 1&nbsp;&nbsp;&nbsp;Controls screen display

Syntax 2&nbsp;&nbsp;&nbsp;Controls display of Info Window

Return to [top](#A)

# DISPLAY Syntax 1

Controls whether the screen displays formulas, gridlines, row and column
headings, and other screen attributes. There are two syntax forms of
this function. Use syntax 1 to control screen display. This function is
provided for compatibility with Microsoft Excel version 4.0. To control
screen display in Microsoft Excel version 5.0 or later, see
OPTIONS.VIEW.

Arguments for this syntax form correspond to options and check boxes in
the Display Options dialog box in Microsoft Excel version 4.0. Arguments
that correspond to check boxes are logical values. If an argument is
TRUE, Microsoft Excel selects the check box; if FALSE, Microsoft Excel
clears the check box. If an argument is omitted, no action is taken.

**Syntax**

**DISPLAY**(formulas, gridlines, headings, zeros, color\_num, reserved,
outline, page\_breaks, object\_num)

**DISPLAY**?(formulas, gridlines, headings, zeros, color\_num, reserved,
outline, page\_breaks, object\_num)

Formulas&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Formulas check box.
The default is FALSE on worksheets and TRUE on macro sheets.

Gridlines&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Gridlines check box.
The default is TRUE.

Headings&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Row & Column Headings
check box. The default is TRUE.

Zeros&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Zero Values check box.
The default is TRUE.

Color\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 0 to 56 corresponding
to the gridline and heading colors in the Display Options dialog box; 0
corresponds to automatic color and is the default value.

Reserved&nbsp;&nbsp;&nbsp;&nbsp;is reserved for certain international
versions of Microsoft Excel.

Outline&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Outline Symbols check
box. The default is TRUE.

Page\_breaks&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Automatic Page
Breaks check box. The default is FALSE.

Object\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 corresponding
to the display options in the Object box.

|                 |                    |
| --------------- | ------------------ |
| **Object\_num** | **Corresponds to** |
| 1 or omitted    | Show All           |
| 2               | Show Placeholders  |
| 3               | Hide               |

**Related Functions**

OPTIONS.VIEW&nbsp;&nbsp;&nbsp;Controls display

WORKSPACE&nbsp;&nbsp;&nbsp;Changes workspace settings

ZOOM&nbsp;&nbsp;&nbsp;Enlarges or reduces a sheet in the active window

Syntax 2&nbsp;&nbsp;&nbsp;Controls display of Info Window

Return to [top](#A)

# DISPLAY Syntax 2

Equivalent to clicking the commands on the Info menu when the Info
Window is active. Controls which commands on the Info Window are in
effect. There are two syntax forms of this function. Use syntax 2 to
control the display of the Info Window. The Info Window must be active
to use this form of DISPLAY. This function is included for compatibility
with Microsoft Excel 95 or earlier; the Info Window is not available in
Microsoft Excel 97 or later.

Arguments in this syntax form correspond to commands on the Info menu
with the same names.

For these arguments:

  - > If the argument is TRUE, Microsoft Excel displays the
    > corresponding Info item.

  - > If the argument is FALSE, Microsoft Excel does not display the
    > corresponding Info item.

  - > If the argument is omitted, the status of the item is unchanged.

> &nbsp;

**Syntax**

For controlling Info Window display

**DISPLAY**(cell, formula, value, format, protection, names, precedents,
dependents, note)

Cell&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to the
Cell command and controls the display of cell information in the Info
Window. If TRUE, cell information will be displayed; if FALSE, cell
information will not be displayed.

Formula&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to
the Formula command and controls the display of formula information in
the Info Window. If TRUE, formula information will be displayed; if
FALSE, formula information will not be displayed.

Value&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to the
Value command and controls the display of value information in the Info
Window. If TRUE, value information will be displayed; if FALSE, value
information will not be displayed.

Format&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to the
Format command and controls the display of format information in the
Info Window. If TRUE, format information will be displayed; if FALSE,
format information will not be displayed.

Protection&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to
the Protection command and controls the display of protection
information in the Info Window. If TRUE, protection information will be
displayed; if FALSE, protection information will not be displayed.

Names&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to the
Names command and controls the display of name information in the Info
Window. If TRUE, name information will be displayed; if FALSE, name
information will not be displayed.

Precedents&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 that specifies
which precedents to list, according to the following table.

Dependents&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 that specifies
which dependents to list, according to the following table.

|                              |             |
| ---------------------------- | ----------- |
| **Precedents or dependents** | **List**    |
| 0                            | None        |
| 1                            | Direct only |
| 2                            | All levels  |

Note&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to the
Note command and controls the display of note information in the Info
Window. If TRUE, note information will be displayed; if FALSE, note
information will not be displayed.

**Related Functions**

SHOW.INFO&nbsp;&nbsp;&nbsp;Controls the display of the Info Window

ZOOM&nbsp;&nbsp;&nbsp;Enlarges or reduces a sheet in the active window

Syntax 1&nbsp;&nbsp;&nbsp;Controls screen display

Return to [top](#A)

# DOCUMENTS

Returns, as a horizontal array in text form, the names of the specified
open workbooks. Use DOCUMENTS to retrieve the names of open workbooks to
use in other functions that manipulate open workbooks.

**Syntax**

**DOCUMENTS**(type\_num, match\_text)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying whether to
include add-in workbooks in the array of workbooks, according to the
following table.

|               |                                                     |
| ------------- | --------------------------------------------------- |
| **Type\_num** | **Returns**                                         |
| 1 or omitted  | Names of all open workbooks except add-in workbooks |
| 2             | Names of add-in workbooks only                      |
| 3             | Names of all open workbooks                         |

Match\_text&nbsp;&nbsp;&nbsp;&nbsp;specifies the workbooks whose names
you want returned and can include wildcard characters. If match\_text is
omitted, DOCUMENTS returns the names of all open workbooks.

**Remarks**

  - > Use the INDEX function to select individual workbook names from
    > the array to use in other functions that take workbook names as
    > arguments.

  - > Use COLUMNS to count the number of entries in the horizontal
    > array.

  - > Use TRANSPOSE to change a horizontal array to a vertical one.

  - > Since the DOCUMENTS function only returns actual workbook names,
    > it ignores any changes made by the WINDOW.TITLE function.

> &nbsp;

**Examples**

In Microsoft Excel for Windows, if your workspace contains windows named
BUDGET.XLS, CHART1, ACTUAL.XLS:1, ACTUAL.XLS:2, and BOOK.XLS, then:

DOCUMENTS (1) equals the four-cell array {"ACTUAL.XLS", "BOOK.XLS",  
"BUDGET.XLS", "CHART1"}

SET.NAME("Document\_array", DOCUMENTS()) defines the name,
Document\_array, as {"ACTUAL.XLS", "BOOK.XLS", "BUDGET.XLS", "CHART1"}

In Microsoft Excel for the Macintosh, if your workspace contains windows
named BUDGET CHART1, ACTUALS, ACTUALS:2, and BOOK then:

DOCUMENTS(1) equals the four-cell array {"ACTUALS", "BOOK", "BUDGET",
"CHART1"}

**Related Functions**

FILES&nbsp;&nbsp;&nbsp;Returns the filenames in the specified directory
or folder

GET.DOCUMENT&nbsp;&nbsp;&nbsp;Returns information about a workbook

GET.WINDOW&nbsp;&nbsp;&nbsp;Returns information about a window

WINDOWS&nbsp;&nbsp;&nbsp;Returns the names of all open windows

Return to [top](#A)

# DUPLICATE

Duplicates the selected object. If an object is not selected, returns
the \#VALUE\! error value and interrupts the macro.

**Syntax**

**DUPLICATE**( )

**Related Functions**

COPY&nbsp;&nbsp;&nbsp;Copies and pastes data or objects

PASTE&nbsp;&nbsp;&nbsp;Pastes cut or copied data

Return to [top](#A)
