<span id="E" class="anchor"></span>This document contains reference
information on the following Excel macro functions:

# E

[ECHO](#echo), [EDITBOX.PROPERTIES](#editbox.properties),
[EDIT.COLOR](#edit.color), [EDIT.DELETE](#edit.delete),
[EDITION.OPTIONS](#edition.options), [EDIT.OBJECT](#edit.object),
[EDIT.REPEAT](#edit.repeat), [EDIT.SERIES](#edit.series),
[EDIT.TOOL](#edit.tool), [ELSE](#else), [ELSE.IF](#else.if),
[EMBED](#embed), [ENABLE.COMMAND](#enable.command),
[ENABLE.OBJECT](#enable.object), [ENABLE.TIPWIZARD](#enable.tipwizard),
[ENABLE.TOOL](#enable.tool), [END.IF](#end.if),
[ENTER.DATA](#enter.data), [ERROR](#error), [ERRORBAR.X,
ERRORBAR.Y](#errorbar.x-errorbar.y), [EVALUATE](#evaluate),
[EXEC](#exec), [EXECUTE](#execute), [EXPON](#expon),
[EXTEND.POLYGON](#extend.polygon), [EXTRACT](#extract)

# F

[FCLOSE](#fclose), [FILE.CLOSE](#file.close),
[FILE.DELETE](#file.delete), [FILES](#files), [FILL.AUTO](#fill.auto),
[FILL.DOWN, FILL.LEFT, FILL.RIGHT,
FILL.UP](#fill.down-fill.left-fill.right-fill.up),
[FILL.GROUP](#fill.group), [FILTER](#filter),
[FILTER.ADVANCED](#filter.advanced),
[FILTER.SHOW.ALL](#filter.show.all), [FIND.FILE](#find.file),
[FONT](#font), [FONT.PROPERTIES](#font.properties), [FOPEN](#fopen),
[FOR](#for), [FOR.CELL](#for.cell), [FORMAT.AUTO](#format.auto),
[FORMAT.CHART](#format.chart), [FORMAT.CHARTTYPE](#format.charttype),
[FORMAT.FONT](#format.font), [FORMAT.LEGEND](#format.legend),
[FORMAT.MAIN](#format.main), [FORMAT.MOVE](#format.move), [FORMAT.MOVE
Syntax 1](#format.move-syntax-1), [FORMAT.MOVE Syntax
2](#format.move-syntax-2), [FORMAT.MOVE Syntax
3](#format.move-syntax-3), [FORMAT.NUMBER](#format.number),
[FORMAT.OVERLAY](#format.overlay), [FORMAT.SHAPE](#format.shape),
[FORMAT.SIZE](#format.size), [FORMAT.SIZE Syntax
1](#format.size-syntax-1), [FORMAT.SIZE Syntax
2](#format.size-syntax-2), [FORMAT.TEXT](#format.text),
[FORMULA](#formula), [FORMULA Syntax 1](#formula-syntax-1), [FORMULA
Syntax 2](#formula-syntax-2), [FORMULA.ARRAY](#formula.array),
[FORMULA.CONVERT](#formula.convert), [FORMULA.FILL](#formula.fill),
[FORMULA.FIND](#formula.find), [FORMULA.FIND.NEXT,
FORMULA.FIND.PREV](#formula.find.next-formula.find.prev),
[FORMULA.GOTO](#formula.goto), [FORMULA.REPLACE](#formula.replace),
[FOURIER](#fourier), [FPOS](#fpos), [FREAD](#fread),
[FREADLN](#freadln), [FREEZE.PANES](#freeze.panes), [FSIZE](#fsize),
[FTESTV](#ftestv), [FULL](#full), [FULL.SCREEN](#full.screen),
[FUNCTION.WIZARD](#function.wizard), [FWRITE](#fwrite),
[FWRITELN](#fwriteln)

# G

[GALLERY.3D.AREA](#gallery.3d.area), [GALLERY.3D.BAR](#gallery.3d.bar),
[GALLERY.3D.COLUMN](#gallery.3d.column),
[GALLERY.3D.LINE](#gallery.3d.line), [GALLERY.3D.PIE](#gallery.3d.pie),
[GALLERY.3D.SURFACE](#gallery.3d.surface),
[GALLERY.AREA](#gallery.area), [GALLERY.BAR](#gallery.bar),
[GALLERY.COLUMN](#gallery.column), [GALLERY.CUSTOM](#gallery.custom),
[GALLERY.DOUGHNUT](#gallery.doughnut), [GALLERY.LINE](#gallery.line),
[GALLERY.PIE](#gallery.pie), [GALLERY.RADAR](#gallery.radar),
[GALLERY.SCATTER](#gallery.scatter), [GET.BAR](#get.bar), [GET.BAR
Syntax 1](#get.bar-syntax-1), [GET.BAR Syntax 2](#get.bar-syntax-2),
[GET.CELL](#get.cell), [GET.CHART.ITEM](#get.chart.item),
[GET.DEF](#get.def), [GET.DOCUMENT](#get.document),
[GET.FORMULA](#get.formula), [GET.LINK.INFO](#get.link.info),
[GET.NAME](#get.name), [GET.NOTE](#get.note), [GET.OBJECT](#get.object),
[GET.PIVOT.FIELD](#get.pivot.field), [GET.PIVOT.ITEM](#get.pivot.item),
[GET.PIVOT.TABLE](#get.pivot.table), [GET.TOOL](#get.tool),
[GET.TOOLBAR](#get.toolbar), [GET.WINDOW](#get.window),
[GET.WORKBOOK](#get.workbook), [GET.WORKSPACE](#get.workspace),
[GOAL.SEEK](#goal.seek), [GOTO](#goto), [GRIDLINES](#gridlines),
[GROUP](#group)

# 

# ECHO

Controls screen updating while a macro is running. If a large macro uses
many commands that update the screen, use ECHO to make the macro run
faster.

**Syntax**

**ECHO**(logical)

Logical    is a logical value specifying whether screen updating is on
or off.

  - > If logical is TRUE, Microsoft Excel selects screen updating.

  - > If logical is FALSE, Microsoft Excel clears screen updating.

  - > If logical is omitted, Microsoft Excel changes the current screen
    > update condition.

>  

**Remarks**

  - > Screen updating is always turned back on when a macro ends.

  - > You can use GET.WORKSPACE to determine whether screen updating is
    > on or off.

>  

**Related Function**

GET.WORKSPACE   Returns information about the workspace

Return to [top](#E)

# EDITBOX.PROPERTIES

Sets the properties of an edit box on a dialog sheet.

**Syntax**

**EDITBOX.PROPERTIES**(validation\_num, multiline\_logical,
vscroll\_logical, password\_logical)

**EDITBOX.PROPERTIES**?(validation\_num, multiline\_logical,
vscroll\_logical, password\_logical)

Validation\_num    is the validation applied to the edit box when the
dialog is dismissed. If the edit box contains a value other than the
type specified (or validation), an error is returned.

|                     |                                |
| ------------------- | ------------------------------ |
| **Validation\_num** | **Type**                       |
| 1                   | Text                           |
| 2                   | Integer                        |
| 3                   | Number (allows floating point) |
| 4                   | Reference                      |
| 5                   | Formula                        |

Multiline\_logical    is a logical value specifying whether word
wrapping is allowed in the edit box control. If TRUE, word wrapping is
allowed. If FALSE, word wrapping is not allowed

Vscroll\_logical    is a logical value specifying whether edit box
displays a vertical scrollbar. If TRUE, a scrollbar is displayed. If
FALSE, a scrollbar is not displayed.

Password\_logical    is a logical value specifying whether edit box
displays characters as the user types. If TRUE, asterisks (\*) are
displayed as the user types. If FALSE, no asterisks are displayed.

**Related Functions**

CHECKBOX.PROPERTIES   Sets various properties of check box and option
box controls

PUSHBUTTON.PROPERTIES   Sets the properties of the push button control

Return to [top](#E)

# EDIT.COLOR

Equivalent to clicking the Modify button on the Color tab, which appears
when you click the Options command on the Tools menu. Defines the color
for one of the 56 color palette boxes.

Use EDIT.COLOR if you want to use a color that is not currently on the
palette and if your system hardware has more than 56 colors available.
After you set the color for the color box, any items previously
formatted with that color are displayed in the new color.

**Syntax**

**EDIT.COLOR**(**color\_num**, red\_value, green\_value, blue\_value)

**EDIT.COLOR**?(**color\_num**, red\_value, green\_value, blue\_value)

Color\_num    is a number from 1 to 56 specifying one of the 56 color
palette boxes for which you want to set the color.

Red\_value, green\_value, and blue\_value    are numbers that specify
how much red, green, and blue are in each color.

  - > In Microsoft Excel for Windows, red\_value, green\_value, and
    > blue\_value are numbers from 0 to 255.

  - > In Microsoft Excel for the Macintosh, red\_value, green\_value,
    > and blue\_value are also numbers from 0 to 255. However, the color
    > editing dialog box displays numbers from 0 to 65, 535. Microsoft
    > Excel automatically converts the numbers between the two ranges.
    > This allows you to display similar colors in all operating
    > environments without modifying your macros.

  - > If red\_value, green\_value, and blue\_value are all set to 255,
    > the resulting color is white. If they are all set to zero, the
    > resulting color is black.

  - > If red\_value, green\_value, or blue\_value is omitted, Microsoft
    > Excel assumes it to be the appropriate value for that color\_num.

>  

**Remarks**

  - > Your system hardware determines the number of unique colors that
    > you can choose from and the number of colors that can be displayed
    > on the screen at the same time.

  - > EDIT.COLOR does not use hue, saturation, or brightness values. If
    > you are using the macro recorder and set the color of a color
    > palette box using hue, saturation, and luminance, Microsoft Excel
    > records the corresponding red, green, and blue values instead.

  - > The dialog-box form of this function, EDIT.COLOR?(color\_num),
    > displays your system's color editing dialog box. The default
    > red\_value, green\_value, and blue\_value are determined by the
    > current settings for the color\_num you specify. Color\_num is a
    > required argument for the dialog-box form of this function.

>  

**Related Function**

COLOR.PALETTE   Copies a color palette from one workbook to another

Return to [top](#E)

# EDIT.DELETE

Equivalent to clicking the Delete command on the Edit menu. Removes the
selected cells from the worksheet and shifts other cells to close up the
space.

**Syntax**

**EDIT.DELETE**(shift\_num)

**EDIT.DELETE**?(shift\_num)

Shift\_num    is a number from 1 to 4 specifying whether to shift cells
left or up after deleting the current selection or else to delete the
entire row or column.

|                |                       |
| -------------- | --------------------- |
| **Shift\_num** | **Result**            |
| 1              | Shifts cells left     |
| 2              | Shifts cells up       |
| 3              | Deletes entire row    |
| 4              | Deletes entire column |

 

  - > If shift\_num is omitted and if one cell or a horizontal range is
    > selected, EDIT.DELETE shifts cells up.

  - > If shift\_num is omitted and a vertical range is selected,
    > EDIT.DELETE shifts cells left.

>  

**Related Function**

CLEAR   Clears specified information from the selected cells or chart

Return to [top](#E)

# EDITION.OPTIONS

Sets options in, or performs actions on, the specified publisher or
subscriber. In Microsoft Excel for Windows, EDITION.OPTIONS also allows
you to cancel a publisher or subscriber created in Microsoft Excel for
the Macintosh.

**Syntax**

**EDITION.OPTIONS**(**edition\_type**, edition\_name, reference,
**option**, appearance, size, formats)

Edition\_type    is the number 1 or 2 specifying the type of edition.

|                   |                     |
| ----------------- | ------------------- |
| **Edition\_type** | **Type of edition** |
| 1                 | Publisher           |
| 2                 | Subscriber          |

Edition\_name    is the name of the edition you want to change the
edition options for or to perform actions on. If edition\_name is
omitted, reference is required.

Reference    specifies the range (given in text form as a name or an
R1C1-style reference) occupied by the publisher or subscriber.

  - > Reference is required if you have more than one publisher or
    > subscriber of edition\_name on the active workbook. Use reference
    > to specify the location of the publisher or subscriber for which
    > you want to set options.

  - > If edition\_type is 1 and the publisher is an embedded chart, or
    > if edition\_type is 2 and the subscriber is a picture, reference
    > is the object identifier as displayed in the reference area.

  - > If reference is omitted, edition\_name is required.

>  

Option    is a number from 1 to 6 specifying the edition option you want
to set or the action you want to take, according to the following two
tables. Options 2 to 6 are only available if you are using Microsoft
Excel for the Macintosh with system software version 7.0 or later.

If a publisher is specified, then option applies as follows.

|            |                                                                        |
| ---------- | ---------------------------------------------------------------------- |
| **Option** | **Action**                                                             |
| 1          | Cancels the publisher                                                  |
| 2          | Sends the edition now                                                  |
| 3          | Selects the range or object published to the specified edition         |
| 4          | Automatically updates the edition when the file is saved               |
| 5          | Updates the edition on request only                                    |
| 6          | Changes the edition file as specified by appearance, size, and formats |

If a subscriber is specified, then option applies as follows.

|            |                                                  |
| ---------- | ------------------------------------------------ |
| **Option** | **Action**                                       |
| 1          | Cancels the subscriber                           |
| 2          | Gets the latest edition                          |
| 3          | Opens the publisher workbook                     |
| 4          | Automatically updates when new data is available |
| 5          | Update on request only                           |

The following three arguments are available only when option is 6.

Appearance    specifies whether the selection is published as shown on
screen or as shown when printed. The default value for appearance is 1
if the selection is a sheet or macro sheet and 2 if the selection is a
chart.

|                |                            |
| -------------- | -------------------------- |
| **Appearance** | **Selection is published** |
| 1              | As shown on screen         |
| 2              | As shown when printed      |

Size    specifies the size of a published chart. Size is only available
if a chart is to be published.

|              |                             |
| ------------ | --------------------------- |
| **Size**     | **Chart size is published** |
| 1 or omitted | As shown on screen          |
| 2            | As shown when printed       |

Formats    is a number specifying the format of the file.

|              |                 |
| ------------ | --------------- |
| **Formats**  | **File format** |
| 1 or omitted | PICT            |
| 2            | BIFF            |
| 4            | RTF             |
| 8            | VALU            |

You can also use the sum of the allowable file formats. For example, a
value of 6 specifies BIFF and RTF.

**Example**

The following macro formula opens the workbook (and application) that
published the edition named Monthly Totals:

EDITION.OPTIONS(2, "Monthly Totals", , 3)

**Related Functions**

CREATE.PUBLISHER   Creates a publisher from the selection

GET.LINK.INFO   Returns information about a link

SUBSCRIBE.TO   Inserts contents of an edition into the active workbook

Return to [top](#E)

# EDIT.OBJECT

Equivalent to clicking the Edit command on the (selected object) Object
submenu of the Edit menu. Starts the application associated with the
selected object and makes the object available for editing or other
actions.

**Syntax**

**EDIT.OBJECT**(verb\_num)

Verb\_num    is a number specifying which verb to use while working with
the object, that is, what you want to do with the object.

  - > The available verbs are determined by the object's source
    > application. 1 often specifies "edit, " and 2 often specifies
    > "play" (for sound, animation, and so on). For more information,
    > consult the documentation for the object's application to see how
    > it supports object linking and embedding (OLE).

  - > If the object does not support multiple verbs, verb\_num is
    > ignored.

  - > If verb\_num is omitted, it is assumed to be 1.

>  

**Remarks**

Your macro pauses while you're editing the object and resumes when you
return to Microsoft Excel.

**Related Function**

INSERT.OBJECT   Creates an object of a specified type

Return to [top](#E)

# EDIT.REPEAT

Equivalent to clicking the Repeat command on the Edit menu. Repeats
certain actions and commands. EDIT.REPEAT is available in the same
situations as the Repeat command.

**Syntax**

**EDIT.REPEAT**( )

Return to [top](#E)

# EDIT.SERIES

Equivalent to clicking the Edit Series command on the Chart menu in
Microsoft Excel version 4.0. Creates or changes chart series by adding a
new SERIES formula or modifying an existing SERIES formula in the
topmost chart type. Chart types are displayed in the following order
from top to bottom: XY (Scatter), Line, Column, Bar, Area.

**Syntax**

**EDIT.SERIES**(**series\_num**, **name\_ref**, **x\_ref**, **y\_ref**,
**z\_ref**, **plot\_order**)

**EDIT.SERIES**?(series\_num, name\_ref, x\_ref, y\_ref, z\_ref,
plot\_order)

Series\_num    is the number of the series you want to change. If
series\_num is 0 or omitted, Microsoft Excel creates a new data series.

Name\_ref    is the name of the data series. It can be an external
reference to a single cell, a name defined as a single cell, or a name
defined as a sequence of characters. Name\_ref can also be text (for
example, "Projected Sales").

X\_ref    is an external reference to the name of the sheet and the
cells that contain one of the following sets of data:

  - > Category labels for all charts except xy (scatter) charts

  - > X-coordinate data for xy (scatter) charts

Y\_ref    is an external reference to the name of the sheet and the
cells that contain values (or y-coordinate data in xy (scatter) charts)
for all 2-D charts. Y\_ref is required in 2-D charts but does not apply
to 3-D charts.

Z\_ref    is an external reference to the name of the sheet and the
cells that contain values for all 3-D charts. Z\_ref is required in 3-D
charts but does not apply to 2-D charts.

Plot\_order    is a number specifying whether the data series is plotted
first, second, and so on, in the chart type.

  - > If you assign a plot\_order to a series, Microsoft Excel plots
    > that series in the order you specify, and the series that
    > previously had that plot order (and any series following it) has
    > its plot order increased by one.

  - > If you add a series to a chart with an overlay, the number of
    > series in the main chart does not change, so if the series is
    > added to the main chart, then the series that was plotted last in
    > the main chart will be plotted first in the overlay chart. To
    > change which series is plotted first in the overlay chart, use the
    > (chart type) Group command from the Format menu, and then select
    > the Series Order tab in the Format (chart type) Group dialog box.
    > You can also use the FORMAT.OVERLAY function.

  - > If you omit plot\_order when you add a new series, then Microsoft
    > Excel plots that series last and assigns it the correct
    > plot\_order value.

  - > The maximum value for plot\_order is 255.

**Remarks**

To change where a series is plotted within a chart, you can change the
chart type, using the FORMAT.CHART function, or the plot order. Plot
order affects where the series appears within the chart type only.

X\_ref, y\_ref, and z\_ref can be arrays or references to a nonadjacent
selection, although they cannot be names that refer to a nonadjacent
selection. If you specify a nonadjacent selection for any of these
arguments, make sure to enclose the reference to the selection in
parentheses so that Microsoft Excel does not treat the components of the
references as separate arguments.

**Tip**   To delete a data series, use the SELECT("Sn") macro function,
where n is the series number, followed by the FORMULA("") macro
function. You can also use the CLEAR function instead of FORMULA.

**Related Function**

FORMAT.CHART

Return to [top](#E)

# EDIT.TOOL

Displays the Button Editor dialog box, which you use to change the
appearance of a button on a toolbar.

**Syntax**

**EDIT.TOOL**(**bar\_id**, **position**)

Bar\_id    is the number of the toolbar containing the button you want
to edit. For a list of toolbar numbers, see ADD.TOOL. Use the
GET.TOOLBAR function to return the information about a toolbar.

Position    is the position on the toolbar of the button you want to
edit. Buttons are numbered from the left starting at 1. Gaps between
buttons are counted as positions.

**Related Functions**

ADD.TOOL   Adds a button to a toolbar

GET.TOOLBAR   Returns information about a toolbar

Return to [top](#E)

# ELSE

Used with IF, ELSE.IF, and END.IF to control which functions are carried
out in a macro. ELSE signals the beginning of a group of formulas in a
macro sheet that will be carried out if the results of all preceding
ELSE.IF statements and the preceding IF statement are FALSE. Use ELSE
with IF, ELSE.IF, and END.IF when you want to perform multiple actions
based on a condition. This method is preferable to using GOTO because it
makes your macros more structured.

**Syntax**

**ELSE**( )

**Remarks**

ELSE must be entered in a cell by itself. In other words, the cell can
contain only "=ELSE()".

For more information about ELSE, ELSE.IF, END.IF, and IF, and for
examples of these functions, see form 2 of the IF function.

**Related Functions**

ELSE.IF   Specifies an action to take if an IF or another ELSE.IF
function returns FALSE

END.IF   Ends a group of macro functions started with an IF statement

IF   Specifies an action to take if a logical test is TRUE

Return to [top](#E)

# ELSE.IF

Used with IF, ELSE, and END.IF to control which functions are carried
out in a macro. ELSE.IF signals the beginning of a group of formulas in
a macro sheet that will be carried out if the preceding IF or ELSE.IF
function returns FALSE and if logical\_test is TRUE. Use ELSE.IF with
IF, ELSE, and END.IF when you want to perform multiple actions based on
a condition. This method is preferable to using GOTO because it makes
your macros more structured.

**Syntax**

**ELSE.IF**(**logical\_test**)

Logical\_test    is a logical value that ELSE.IF uses to determine what
functions to carry out next—that is, where to branch.

  - > If logical\_test is TRUE, Microsoft Excel carries out the
    > functions between the ELSE.IF function and the next ELSE.IF, ELSE,
    > or END.IF function.

  - > If logical\_test is FALSE, Microsoft Excel immediately branches to
    > the next ELSE.IF, ELSE, or END.IF function.

>  

**Remarks**

  - > ELSE.IF must be entered in a cell by itself.

  - > Logical\_test will always be evaluated, even if the ELSE.IF
    > section is not reached (due to a previous IF or ELSE.IF
    > logical\_test evaluating to TRUE). For this reason, you should not
    > use formulas that carry out actions for logical\_test. If you need
    > to base the ELSE.IF condition on the return value of a formula
    > that carries out an action, use the form "ELSE, IF(logical\_test),
    > and END.IF" in place of "ELSE.IF(logical\_test)."

>  

For more information about ELSE, ELSE.IF, END.IF, and IF, and for
examples of these functions, see form 2 of the IF function.

**Related Functions**

ELSE   Specifies an action to take if an IF function returns FALSE

END.IF   Ends a group of macro functions started with an IF statement

IF   Specifies an action to take if a logical test is TRUE

Return to [top](#E)

# EMBED

Displayed in the formula bar when an embedded object is selected. EMBED
cannot be entered on a sheet or used in a macro.

**Syntax**

**EMBED**(**object\_class**, item)

Object\_class    is the name of the application and document type that
created the embedded object. For example, the object\_class arguments
used when Microsoft Excel sheets are embedded in other applications are
"Excel.sheet.5" and "Excel.Chart.5".

Item    is the area selected to copy, and determines the view on the
embedded document. When item is empty text (""), EMBED creates a view on
the entire document.

**Remarks**

If you delete the EMBED formula, the embedded object remains on the
sheet as a graphic, and the link to the creating application is deleted.
Double-clicking the object no longer starts the creating application.

Return to [top](#E)

# ENABLE.COMMAND

Enables or disables a custom command or menu. Disabled commands appear
dimmed and can't be chosen. Use ENABLE.COMMAND to control which commands
the user can click in a menu bar.

**Syntax**

**ENABLE.COMMAND**(**bar\_num**, **menu**, **command**, **enable**,
subcommand)

Bar\_num    is the menu bar in which a command resides. Bar\_num can be
the number of a built-in menu bar or the number returned by a previously
run ADD.BAR function. See ADD.COMMAND for a list of the built-in menu
bar numbers.

Menu    is the menu on which the command resides. Menu can be either the
name of a menu as text or the number of a menu. Menus are numbered
starting with 1 from the left of the screen.

Command    is the command you want to enable or disable. Command can be
either the name of the command as text or the number of the command. The
top command on a menu is command 1. If command is 0, ENABLE.COMMAND
enables or disables the entire menu.

Enable    is a logical value specifying whether the command should be
enabled or disabled. If enable is TRUE, Microsoft Excel enables the
command; if FALSE, it disables the command.

Subcommand    is the name of the command on a submenu that you want to
enable. If you use subcommand, you must use command as the name of the
submenu. Use subcommand 0 to enable an entire submenu.

**Remarks**

  - > You cannot disable built-in commands. If the specified command is
    > a built-in command or does not exist, ENABLE.COMMAND returns the
    > \#VALUE\! error value and interrupts the macro.

  - > You can hide any shortcut menu from users by using ENABLE.COMMAND
    > with command set to 0.

>  

**Example**

The following macro formula disables a custom command that had been
added previously to the View menu on the worksheet and macro sheet menu
bar:

ENABLE.COMMAND(10, "View", "Audit...", FALSE)

**Related Functions**

ADD.BAR   Adds a menu bar

ADD.COMMAND   Adds a command to a menu

CHECK.COMMAND   Adds or deletes a check mark to or from a command

DELETE.COMMAND   Deletes a command from a menu

RENAME.COMMAND   Changes the name of a command or menu

Return to [top](#E)

# ENABLE.OBJECT

Enables or disables a drawing object or the selected drawing object. A
disabled object will not run any macro events assigned to it, and the
controls will be grayed out.

**Syntax**

**ENABLE.OBJECT**(object\_id\_text, **enable\_logical**)

Object\_id\_text    is the name of the object(s) as text. If omitted,
the selected object(s) are assumed.

Enable\_logical    is a logical value that specifies whether the object
is to be enabled. If TRUE, the object is enabled. If FALSE, the object
is disabled.

**Examples**

ENABLE.OBJECT("Button 2",FALSE) disables the button with object name
Button 2 on the dialog box.

**Related Function**

SET.CONTROL.VALUE   Changes the value of the active control

Return to [top](#E)

# ENABLE.TIPWIZARD

This function should not be used. The TipWizard has been removed from
Microsoft Excel.

Return to [top](#E)

# ENABLE.TOOL

Enables or disables a button on a toolbar. An enabled button can be
accessed by the user. Disabled buttons may still be visible but cannot
be accessed. Use ENABLE.TOOL to control which buttons the user can click
in a particular situation.

**Syntax**

**ENABLE.TOOL**(**bar\_id, position**, enable)

Bar\_id    is the number or name of a toolbar on which the button
resides. For detailed information about bar\_id, see ADD.TOOL.

Position    specifies the position of the button on the toolbar.
Position starts with 1 at the left side (if horizontal) or from the top
(if vertical).

Enable    specifies whether the button can be accessed. If enable is
TRUE or omitted, the user can access the button; if FALSE, the user
cannot access it.

**Remarks**

Microsoft Excel sounds a tone if you click a disabled button.

**Example**

The following macro formula enables the fourth button in Toolbar1:

ENABLE.TOOL("Toolbar1", 4, TRUE)

**Related Function**

GET.TOOL   Returns information about a button or buttons on a toolbar

Return to [top](#E)

# END.IF

Ends a block of functions associated with the preceding IF function. You
must include one and only one END.IF function for each macro-sheets-only
syntax form (syntax 2) of the IF function in a macro. Syntax 1 of the IF
function, which can be used on both worksheets and macro sheets, does
not require an END.IF function. Use END.IF with IF, ELSE, and ELSE.IF
when you want to perform multiple actions based on a condition. This
method is preferable to using GOTO because it makes your macros more
structured.

**Syntax**

**END.IF**( )

**Remarks**

  - > If you accidentally omit an END.IF function, your macro will end
    > with an error at the cell containing the first IF function that
    > does not have a corresponding END.IF function.

  - > END.IF must be entered in a cell by itself.

  - > For more information about ELSE, ELSE.IF, END.IF, and IF, and for
    > examples of these functions, see form 2 of the IF function.

>  

**Related Functions**

ELSE   Specifies an action to take if an IF function returns FALSE

ELSE.IF   Specifies an action to take if an IF or another ELSE.IF
function returns FALSE

IF   Specifies an action to take if a logical test is TRUE

Return to [top](#E)

# ENTER.DATA

Turns on Data Entry mode and allows you to select and to enter data into
the unlocked cells in the current selection only (the data entry area).
Use ENTER.DATA when you want to enter data only in a specific part of
your sheet. You can then use that part of the sheet as a simple data
form.

**Syntax**

**ENTER.DATA**(logical)

Logical    is a logical value that turns Data Entry mode on or off.

  - > If logical is TRUE, Data Entry mode is turned on; if FALSE, Data
    > Entry mode is turned off and data entry, cell movement, and cell
    > selection return to normal. If logical is omitted, ENTER.DATA
    > toggles Data Entry mode.

  - > Logical can also be the number 2. This setting turns on Data Entry
    > mode and prevents the ESC key from turning it off.

  - > Logical can also be a reference. Using a reference for this
    > argument turns on Data Entry mode for the supplied reference.

>  

**Remarks**

  - > In Data Entry mode, you can move the active cell and select cell
    > ranges only in the data entry area. The arrow keys and the TAB and
    > SHIFT+TAB keys move from one unlocked cell to the next. The HOME
    > and END keys move to the first and last cell in the data entry
    > area, respectively. You cannot select entire rows or columns, and
    > clicking a cell outside the data entry area does not select it.

  - > The only commands available in Data Entry mode are commands
    > normally available to protected workbooks.

  - > To turn off Data Entry mode, press ESC (unless logical is 2),
    > activate another sheet in the active workbook window, or use
    > another ENTER.DATA function. If you use another ENTER.DATA
    > function, you will usually design your macros in one of two ways:
    
      - > The macro turns on Data Entry mode, pauses while you enter
        > data, resumes, and then turns off Data Entry mode.
    
      - > The macro turns on Data Entry mode and ends. After entering
        > data, another macro turns off Data Entry mode; this latter
        > macro could be assigned to a "Finished" button, for example.

> With either method, you can use Microsoft Excel's ON functions to
> resume or run other macros based on an event, such as pressing the
> CONTROL+D keys.
> 
>  

**Tips**

  - > Normally you use Data Entry mode to enter data, but you can also
    > prevent someone from entering data or moving the active cell by
    > locking all the cells in the current selection before turning on
    > Data Entry mode. This is useful if you want a user to view a range
    > of cells but not change it or move the active cell. Similarly, if
    > you unlock certain cells, you can restrict the user's movement to
    > the Data Entry area only.

  - > To prevent someone from activating another workbook, which would
    > turn off Data Entry mode, use the ON.WINDOW function or an
    > Auto\_Deactivate macro.

**Related Functions**

DISABLE.INPUT   Blocks all input to Microsoft Excel

FORMULA   Enters values into a cell or range or onto a chart

Return to [top](#E)

# ERROR

Specifies what action to take if an error is encountered while a macro
is running. Use ERROR to control whether Microsoft Excel error messages
are displayed, or to run your own macro when an error is encountered.

**Syntax**

**ERROR**(**enable\_logical**, macro\_ref)

Enable\_logical    is a logical value or number that selects or clears
error-checking.

  - > If enable\_logical is FALSE or 0, all error-checking is cleared.
    > If error-checking is cleared and an error is encountered while a
    > macro is running, Microsoft Excel ignores it and continues.
    > Error-checking is selected again by an ERROR(TRUE) statement, or
    > when the macro stops running.

  - > If enable\_logical is TRUE or 1, you can either select normal
    > error-checking (by omitting the other argument) or specify a macro
    > to run when an error is encountered by using the macro\_ref
    > argument. When normal error-checking is active, the Macro Error
    > dialog box is displayed when an error is encountered. You can halt
    > the macro, start single-stepping through the macro, continue
    > running the macro normally, or go to the macro cell where the
    > error occurred.

  - > If enable\_logical is 2 and macro\_ref is omitted, error-checking
    > is normal except that if the user clicks the Cancel button in an
    > alert message, ERROR returns FALSE and the macro is not
    > interrupted.

  - > If enable\_logical is 2 and macro\_ref is given, the macro goes to
    > that macro\_ref when an error is encountered. If the user clicks
    > the Cancel button in an alert message, FALSE is returned and the
    > macro is not interrupted.

>  

Macro\_ref    specifies a macro to run if enable\_logical is TRUE, 1, or
2 and an error is encountered. It can be either the name of the macro or
a cell reference. If enable\_logical is FALSE or 0, macro\_ref is
ignored.

**Important   **Both ERROR(FALSE) and ERROR(TRUE, macro\_ref ) keep
Microsoft Excel from displaying any messages at all, including the
message asking whether to save changes when you close an unsaved
workbook. If you want alert messages but not error messages to be
displayed, use ERROR(2, macro\_ref ).

**Remarks**

You can use GET.WORKSPACE to determine whether error-checking is on or
off.

**Examples**

ERROR(FALSE) clears error-checking.

ERROR(TRUE, Recover) selects error-checking and runs the macro named
Recover when an error is encountered.

The following macro runs the macro ForceMenus if an error occurs in the
current macro:

\=ERROR(TRUE, ForceMenus)

**Related Functions**

CANCEL.KEY   Disables macro interruption

LAST.ERROR   Returns the reference of the cell where the last error
occurred

ON.KEY   Runs a macro when a specified key is pressed

Return to [top](#E)

# ERRORBAR.X, ERRORBAR.Y

Adds error bars to the selected series in a chart. ERRORBAR.X adds bars
showing the error factor for the X (category) axis and works for XY
(scatter) charts only. ERRORBAR.Y adds bars showing the error factor for
the Y (value) axis for all charts.

**Syntax**

**ERRORBAR.X**(include, type, amount, minus)

**ERRORBAR.Y**(include, type, amount, minus)

Include    specifies the type of error value to include:

|              |                         |
| ------------ | ----------------------- |
| **Include**  | **Type of error value** |
| 1 or omitted | Plus and minus          |
| 2            | Plus                    |
| 3            | Minus                   |
| 4            | None                    |

Type    specifies the type of error bars to display:

|              |                                                            |
| ------------ | ---------------------------------------------------------- |
| **Type**     | **Type of error displayed**                                |
| 1 or omitted | Fixed amount                                               |
| 2            | Percent                                                    |
| 3            | Multiplying factor standard deviation (default value is 1) |
| 4            | Standard error                                             |
| 5            | Custom                                                     |

Amount    is the range of error values to display. This argument depends
on the value of type:

|                |                                              |
| -------------- | -------------------------------------------- |
| **If type is** | **then amount**                              |
| 1 or omitted   | Can be any number greater than 0             |
| 2              | Can be any number greater than 0             |
| 3              | Can be any number greater than or 0          |
| 4              | Not required                                 |
| 5              | Is the positive amount for custom error bars |

Minus    is the negative amount for custom error bars. Applicable only
if type is 5.

**Remarks**

For the amount argument, standard deviation(s) can be calculated using
this equation:

![](media/image1.png)

The standard deviation is multiplied by the value specified by amount
and the error bars are placed this distance from the arithmetic mean.
Therefore, these error bars are plotted along the arithmetic mean, not
attached to data series.

Microsoft Excel calculates the standard error using the following
equation:

![](media/image2.png)

Both the standard deviation and standard error functions use the
following variables:

|              |                                           |
| ------------ | ----------------------------------------- |
| **Variable** | **Equals**                                |
| s            | Series number                             |
| i            | Point number in series s                  |
| m            | Number of series for point y in chart     |
| n            | Number of points in each series           |
| Yi           | Data value of series s and the ith point  |
| Ny           | Total number of data values in all series |
| M            | Arithmetic mean                           |

Return to [top](#E)

# EVALUATE

Evaluates a formula or expression that is in the form of text and
returns the result. To run a macro or subroutine, use the RUN function.

**Syntax**

**EVALUATE**(**formula\_text**)

Formula\_text    is the expression in the form of text that you want to
evaluate.

**Remarks**

Using EVALUATE is similar to selecting an expression within a formula in
the formula bar and pressing the Recalculate key (F9 in Microsoft Excel
for Windows and COMMAND+= in Microsoft Excel for the Macintosh).
EVALUATE replaces an expression with a value.

**Example**

Suppose you want to know the value of a cell named LabResult1,
LabResult2, or LabResult3, where the 1, 2, or 3 is specified by the name
TrialNum whose value may change as the macro runs. You can use the
following formula to calculate the value:

EVALUATE("LabResult"\&TrialNum)

**Related Function**

RUN   Runs a macro

Return to [top](#E)

# EXEC

Starts a separate program. Use EXEC to start other programs with which
you want to communicate. Use EXEC with Microsoft Excel's other DDE
functions (INITIATE, EXECUTE, and SEND.KEYS) to create a channel to
another program and to send keystrokes and commands to the program.
(SEND.KEYS is available only in Microsoft Excel for Windows.)

Syntax 1 is for Microsoft Excel for Windows. Syntax 2 is for Microsoft
Excel for the Macintosh.

**Syntax 1**

For Microsoft Excel for Windows

**EXEC**(program\_text, window\_num)

**Syntax 2**

For Microsoft Excel for the Macintosh

**EXEC**(program\_text, , background, preferred\_size\_only)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for the last two arguments of this
function.

Program\_text    is the name, as a text string, of any executable file
or, in Microsoft Excel for Windows, any data file that is associated
with an executable file.

  - > Use paths when the file or program to be started is not in the
    > current directory or folder.

  - > In Microsoft Excel for Windows, program\_text can include any
    > arguments and switches that are accepted by the program to be
    > started. Also, if program\_text is the name of a file associated
    > with a specific installed program, EXEC starts the program and
    > loads the specified file.

>  

**Note   **In Microsoft Excel for the Macintosh, you must use an extra
comma after the program\_text argument. This skips the window\_num
argument that does not apply to the Macintosh.

Window\_num    is a number from 1 to 3 that specifies how the window
containing the program should appear. Window\_num is only available for
use with Microsoft Excel for Windows. The window\_num argument is
allowed on the Macintosh, but it is ignored.

|                 |                    |
| --------------- | ------------------ |
| **Window\_num** | **Window appears** |
| 1               | Normal size        |
| 2 or omitted    | Minimized size     |
| 3               | Maximized size     |

Background    is a logical value that determines whether the program
specified by program\_text is opened as the active program or in the
background, leaving Microsoft Excel as the active program. If background
is TRUE, the program is started in the background; if FALSE or omitted,
the program is started in the foreground. Background is only available
for use with Microsoft Excel for the Macintosh and system software
version 7.0 or later.

Preferred\_size\_only    is a logical value that determines the amount
of memory allocated to the program. If preferred\_size\_only is TRUE,
the program is opened with its preferred memory allocation; if FALSE or
omitted, it opens with the available memory if greater than its minimum
requirement. Preferred\_size\_only is only available for use with
Microsoft Excel for the Macintosh and system software version 7.0 or
later. For information about changing the preferred memory size, see
your Macintosh documentation.

**Remarks**

In Microsoft Excel for Windows and in Microsoft Excel for the Macintosh
with system software version 7.0, if the EXEC function is successful, it
returns the task ID number of the started program. The task ID number is
a unique number that identifies a program. Use the task ID number in
other macro functions, such as APP.ACTIVATE, to refer to the program. In
Microsoft Excel for the Macintosh with system software version 6.0, if
EXEC is successful, it returns TRUE. If EXEC is unsuccessful, it returns
the \#VALUE\! error value.

**Examples**

In Microsoft Excel for Windows, the following macro formula starts the
program SEARCH.EXE. Use paths when the file or program to be started is
not in the current directory:

EXEC("C:\\WINDOWS\\SEARCH.EXE")

The following macro formula starts Microsoft Word for Windows and loads
the document SALES.DOC:

EXEC("C:\\WINWORD\\WINWORD.EXE C:\\MYFILES\\SALES.DOC")

In Microsoft Excel for the Macintosh, the following macro formula starts
Microsoft Word:

EXEC("HARD DISK:APPS:WORD")

**Related Functions**

APP.ACTIVATE   Switches to another application

EXECUTE   Carries out a command in another application

INITIATE   Opens a channel to another application

SEND.KEYS   Sends a key sequence to an application

TERMINATE   Closes a channel to another application

REQUEST   Requests an array of a specific type of information from an
application with which you have a dynamic data exchange (DDE) link

POKE   Sends data to another application with which you have a dynamic
data exchange (DDE) link

Return to [top](#E)

# EXECUTE

Carries out commands in another program with which you have a dynamic
data exchange (DDE) link. Use with EXEC, INITIATE, and SEND.KEYS to run
another program through Microsoft Excel. (SEND.KEYS is available only in
Microsoft Excel for Windows.)

**Important**   Microsoft Excel for the Macintosh requires system
software version 7.0 or later for this function.

**Syntax**

**EXECUTE**(**channel\_num, execute\_text**)

Channel\_num    is a number returned by a previously run INITIATE
function. Channel\_num refers to a channel through which Microsoft Excel
communicates with another program.

Execute\_text    is a text string representing commands you want to
carry out in the program specified by channel\_num. The form of
execute\_text depends on the program you are referring to. To include
specific key sequences in execute\_text, use the format described under
key\_text in the ON.KEY function.

If EXECUTE is not successful, it returns one of the following error
values:

|                    |                                                                                                                  |
| ------------------ | ---------------------------------------------------------------------------------------------------------------- |
| **Value returned** | **Situation**                                                                                                    |
| \#VALUE\!          | Channel\_num is not a valid channel number.                                                                      |
| \#N/A              | The program you are accessing is busy.                                                                           |
| \#DIV/0\!          | The program you are accessing does not respond after a certain length of time or you have pressed ESC to cancel. |
| \#REF\!            | The keys specified in execute\_text are refused by the application which you want to access.                     |

**Remarks**

Commands sent to another program with EXECUTE will not work when a
dialog box is displayed in the program. In Microsoft Excel for Windows,
you can use SEND.KEYS to send commands that make selections in a dialog
box.

**Examples**

The following macro formula sends the number 25 and a carriage return to
the application identified by channel\_num 14:

EXECUTE(14, "25\~")

**Related Functions**

EXEC   Starts another application

INITIATE   Opens a channel to another application

POKE   Sends data to another application

REQUEST   Returns data from another application

SEND.KEYS   Sends a key sequence to an application

TERMINATE   Closes a channel to another application

Return to [top](#E)

# EXPON

Predicts a value based on the forecast for the prior period, adjusted
for the error in that prior forecast.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**EXPON**(**inprng**, outrng, damp, stderrs, chart)

Inprng    is the input range.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Damp    is the damping factor. If omitted, damp is 0.3.

Stderrs    is a logical value. If TRUE, standard error values are
included in the output table. If FALSE, standard errors are not
included.

Chart    is a logical value. If TRUE, EXPON generates a chart for the
actual and forecast values. If FALSE, the chart is not generated.

**Related Function**

MOVEAVG   Returns values along a moving average trend

Return to [top](#E)

# EXTEND.POLYGON

Adds vertices to a polygon. This function must immediately follow a
CREATE.OBJECT function or another EXTEND.POLYGON function. Use multiple
EXTEND.POLYGON functions to create arbitrarily complex polygons. It is
generally easier to use the macro recorder to enter this function on
your macro sheet.

**Syntax**

**EXTEND.POLYGON**(**array**)

Array    is an array of values, or a reference to a range of cells
containing values, that indicate the position of vertices in the
polygon. The position is measured in points and is relative to the
upper-left corner of the polygon's bounding rectangle.

  - > A vertex is a point. Each vertex is defined by a pair of
    > coordinates in one row of array.

  - > The polygon is defined by the array argument to the CREATE.OBJECT
    > function and to all the immediately following EXTEND.POLYGON
    > functions.

  - > If the polygon contains many vertices, one array may not be
    > sufficient to define it. If the number of elements in the formula
    > exceeds 1024, you must include additional EXTEND.POLYGON
    > functions. If you're recording a macro, Microsoft Excel
    > automatically records additional EXTEND.POLYGON functions as
    > needed.

**Related Functions**

CREATE.OBJECT   Creates an object

FORMAT.SHAPE   Inserts, moves, or deletes vertices of the selected
polygon

Return to [top](#E)

# EXTRACT

Equivalent to choosing the Extract command from the Data menu in
Microsoft Excel version 4.0. Finds database records that match the
criteria defined in the criteria range and copies them into a separate
extract range.

**Syntax**

**EXTRACT**(unique)

**EXTRACT**?(unique)

Unique    is a logical value corresponding to the Unique Records Only
check box in the Extract dialog box.

  - > If unique is TRUE, Microsoft Excel selects the check box and
    > excludes duplicate records from the extract list.

  - > If unique is FALSE or omitted, Microsoft Excel clears the check
    > box and extracts all records matching the criteria.

>  

**Related Functions**

DATA.FIND   Finds records in a database

SET.CRITERIA   Defines the name Criteria for the selected range on the
active sheet

SET.DATABASE   Defines the name Database for the selected range on the
active sheet

SET.EXTRACT   Defines the name Extract for the selected range on the
active sheet

Return to [top](#E)

# FCLOSE

Closes the specified file.

**Syntax**

**FCLOSE**(**file\_num**)

File\_num    is the number of the file you want to close. File\_num is
returned by the FOPEN function that originally opened the file. If
file\_num is not a valid file number, FCLOSE halts the macro and returns
the \#VALUE\! error value.

**Examples**

The following function closes the file identified by FileNumber:

FCLOSE(FileNumber)

**Related Functions**

CLOSE   Closes the active window

FILE.CLOSE   Closes the active workbook

FOPEN   Opens a file with the type of permission specified

Return to [top](#E)

# FILE.CLOSE

Equivalent to clicking the Close command on the File menu. Closes the
active workbook.

**Syntax**

**FILE.CLOSE**(save\_logical, route\_logical)

Save\_logical    is a logical value specifying whether to save the file
before closing it.

|                   |                                                                                                       |
| ----------------- | ----------------------------------------------------------------------------------------------------- |
| **Save\_logical** | **Result**                                                                                            |
| TRUE              | Saves the workbook                                                                                    |
| FALSE             | Does not save the workbook                                                                            |
| Omitted           | If you've made changes to the workbook, displays a dialog box asking if you want to save the workbook |

Route\_logical    is a logical value that specifies whether to route the
file after closing it. This argument is ignored if there is not a
routing slip present.

|                    |                         |
| ------------------ | ----------------------- |
| **Route\_logical** | **Result**              |
| TRUE               | Routes the file         |
| FALSE              | Does not route the file |

Omitted   If you've specified recipients for routing, displays a dialog
box asking if you want to save the file

**Remarks**

If you make any changes to the structure of a workbook, such as the name
of sheets, their order, and so on, then a message will be displayed
reminding you that there are unsaved changes, regardless of the
save\_logical value.

**Note   **When you use the FILE.CLOSE function, Microsoft Excel does
not run any Auto\_Close macros before closing the workbook.

**Related Functions**

CLOSE   Closes the active window

CLOSE.ALL   Closes all unprotected windows

FCLOSE   Closes a text file

Return to [top](#E)

# FILE.DELETE

Deletes a file from the disk. Although you will normally delete files
manually, you can, for example, use FILE.DELETE in a macro to delete
temporary files created by the macro.

**Syntax**

**FILE.DELETE**(**file\_text**)

**FILE.DELETE**?(file\_text)

File\_text    is the name of the file to delete.

**Remarks**

  - > If Microsoft Excel can't find file\_text, it displays a message
    > saying that it cannot delete the file. To avoid this, include the
    > entire path in file\_text. See the following second and fifth
    > examples. You can also use FILES to generate an array of filenames
    > and then check if the file you want to delete is in the array.

  - > If a file is open when you delete it, the file is removed from the
    > disk but remains open in Microsoft Excel.

  - > In the dialog-box form, FILE.DELETE?, you can use an asterisk (\*)
    > to represent any series of characters and a question mark (?) to
    > represent any single character. See the following third and sixth
    > examples.

>  

**Examples**

In Microsoft Excel for Windows, the following macro formula deletes a
file called CHART1.XLS from the current directory:

FILE.DELETE("CHART1.XLS")

The following macro formula deletes a file called 92INFO.XLS kept in the
EXCEL\\SALES subdirectory:

FILE.DELETE("C:\\EXCEL\\SALES\\92INFO.XLS")

The following macro formula displays the Delete dialog box listing all
documents whose extensions begin with the letters "XL":

FILE.DELETE?("\*.XL?")

In Microsoft Excel for the Macintosh, the following macro formula
deletes a file called CHART1 from the current folder:

FILE.DELETE("CHART1")

The following macro formula deletes a file called 1992 INFO kept in a
series of nested folders:

FILE.DELETE("HARD DISK:EXCEL 5:SALES WORKSHEETS:1992 INFO")

The following macro formula displays the Delete dialog box listing all
documents beginning with the word "Clients":

FILE.DELETE?("Clients\*")

**Related Functions**

FILE.CLOSE   Closes the active workbook

FILES   Returns the filenames in the specified directory or folder

Return to [top](#E)

# FILES

Returns a horizontal text array of the names of all files in the
specified directory or folder. Use FILES to build a list of filenames
upon which you want your macro to operate.

**Syntax**

**FILES**(directory\_text)

Directory\_text    specifies which directories or folders to return
filenames from.

  - > Directory\_text accepts an asterisk (\*) to represent a series of
    > characters and a question mark (?) to represent a single character
    > in filenames.

  - > If directory\_text is not specified, FILES returns filenames from
    > the current directory.

>  

**Remarks**

If you enter FILES in a single cell, only one filename is returned. You
will normally use FILES with SET.NAME to assign the returned array to a
name. See the last example below.

**Tips**   You can use COLUMNS to count the number of entries in the
returned array. You can use TRANSPOSE to change a horizontal array to a
vertical one.

**Examples**

In Microsoft Excel for Windows, the following macro formula returns the
names of all files starting with the letter F in the current directory
or folder:

FILES("F\*.\*")

When entered as an array formula in several cells, the following macro
formula returns the filenames in the current directory to those cells.
If the directory contains fewer files than can fit in the selected
cells, the \#N/A error value appears in the extra cells.

FILES()

In Microsoft Excel for Windows, the following macro formula returns all
files starting with "SALE" and ending with the .XLS extension in the
\\EXCEL\\CHARTS subdirectory:

FILES("C:\\EXCEL\\CHARTS\\SALE\*.XLS")

In Microsoft Excel for the Macintosh, the following macro formula
returns all files starting with "SALE" in the nested CHART folder:

FILES("DISK:EXCEL:CHART:SALE\*")

The following macro stores the names of the files in the current
directory in the named array FileArray

SET.NAME("FileArray",FILES())

**Related Functions**

DOCUMENTS   Returns the names of the specified open workbooks

FILE.DELETE   Deletes a file

OPEN   Opens a workbook

SET.NAME   Defines a name as a value

Return to [top](#E)

# FILL.AUTO

Equivalent to copying cells or automatically filling a selection by
dragging the fill selection handle with the mouse (the AutoFill
feature).

**Syntax**

**FILL.AUTO**(destination\_ref, copy\_only)

Destination\_ref    is the range of cells into which you want to fill
data. The top, bottom, left, or right end of destination\_ref must
include all of the cells in the source reference (the current
selection).

Copy\_only    is a number specifying whether to copy cells or perform an
AutoFill operation.

|              |                      |
| ------------ | -------------------- |
| **Value**    | **Result**           |
| 0 or omitted | Normal AutoFill      |
| 1 or TRUE    | Copy cells           |
| 2            | Copy formats         |
| 3            | Fill values          |
| 4            | Increment            |
| 5            | Increment by day     |
| 6            | Increment by weekday |
| 7            | Increment by month   |
| 8            | Increment by year    |
| 9            | Linear trend         |
| 10           | Growth trend         |

**Related Functions**

COPY   Copies and pastes data or objects

DATA.SERIES   Fills a range of cells with a series of numbers or dates

Return to [top](#E)

# FILL.DOWN, FILL.LEFT, FILL.RIGHT, FILL.UP

Equivalent to clicking the Down, Left, Right, and Up commands,
respectively, on the Fill submenu of the Edit menu.

**Syntax**

**FILL.DOWN**( )

**FILL.LEFT**( )

**FILL.RIGHT**( )

**FILL.UP**( )

FILL.DOWN copies the contents and formats of the cells in the top row of
a selection into the rest of the rows in the selection.

FILL.LEFT copies the contents and formats of the cells in the right
column of a selection into the rest of the columns in the selection.

FILL.RIGHT copies the contents and formats of the cells in the left
column of a selection into the rest of the columns in the selection.

FILL.UP copies the contents and formats of the cells in the bottom row
of a selection into the rest of the rows in the selection.

**Remarks**

If you have a multiple selection, each range in the selection is filled
separately with the contents of the source range.

**Related Functions**

COPY   Copies and pastes data or objects

DATA.SERIES   Fills a range of cells with a series of numbers or dates

FILL.AUTO   Copies cells or automatically fills a selection

FORMULA.FILL   Enters a formula in the specified range

Return to [top](#E)

# FILL.GROUP

Equivalent to choosing the Across Worksheets command from the Fill
submenu on the Edit menu. Copies the contents of the active worksheet's
selection to the same area on all other worksheets in the group. Use
FILL.GROUP to fill a range of cells on all worksheets in a group at
once.

**Syntax**

**FILL.GROUP**(**type\_num**)

**FILL.GROUP**?(type\_num)

Type\_num    is a number from 1 to 3 that corresponds to the choices in
the Fill Across Worksheets dialog box.

|               |                                |
| ------------- | ------------------------------ |
| **Type\_num** | **Type of information filled** |
| 1             | All                            |
| 2             | Contents                       |
| 3             | Formats                        |

**Related Functions**

NEW   Creates a new workbook

WORKBOOK.SELECT   Selects one or more sheets in a workbook

Return to [top](#E)

# FILTER

Filters lists of data one column at a time. Only one list can be
filtered on any one sheet at a time.

**Syntax**

**FILTER**(field\_num, criteria1, operation, criteria2)

**FILTER**?(field\_num, criteria1, operation, criteria2)

Field\_num    is the number of the field that you want to filter. Fields
are numbered from left to right starting with 1.

Criteria1    is a text string specifying criteria for filtering a list,
such as "\>2". If you want to include all items in the list, omit this
argument.

Operation    is a number that specifies how you want criteria2 used with
criteria1:

|            |                    |
| ---------- | ------------------ |
| **Number** | **Operation Used** |
| 1          | AND                |
| 2          | OR                 |

Criteria2    is a text string specifying criteria for filtering a list,
such as "\>2". If you include this argument, operation is required.

**Remarks**

If you omit all arguments, FILTER toggles the display of filter arrows.

**Related Function**

FILTER.ADVANCED   Lets you set options for filtering a list

Return to [top](#E)

# FILTER.ADVANCED

Equivalent to choosing the Advanced Filter command from the Filter
submenu on the Data menu. Lets you set options for filtering a list.

**Syntax**

**FILTER.ADVANCED**(**operation**, **list\_ref**, criteria\_ref,
copy\_ref, unique)

**FILTER.ADVANCED**?(operation, list\_ref, criteria\_ref, copy\_ref,
unique)

Operation    is a number specifying whether to copy the filter list to a
new location. To filter a list without copying, use 1; to copy the
filter list to a new location, use 2.

List\_ref    specifies the location of the list to be filtered. If
operation is 1, then list\_ref must be on the active sheet.

Criteria\_ref    is a reference to a range containing criteria for
filtering the list. If omitted, uses "All" as the criteria.

Copy\_ref    is a reference on the active sheet where you want the
filtered list copied. Ignored if operation is 1.

Unique    is a logical value that specifies whether only unique records
are displayed. To display only unique records, use TRUE. To display all
records that match the criteria, use FALSE or omit this argument.

**Related Function**

FILTER   Filters lists of data one column at a time

Return to [top](#E)

# FILTER.SHOW.ALL

Equivalent to choosing the Show All command from the Filter submenu on
the Data menu. Displays all items in a filtered list.

**Syntax**

**FILTER.SHOW.ALL**()

Return to [top](#E)

# FIND.FILE

Equivalent to choosing the Find File command from the File menu in
Microsoft Excel version 5.0. Lets you search for files based on criteria
such as author or creation date.

**Syntax**

**FIND.FILE**?( )

**Remarks**

This function has a dialog-box form only.

Return to [top](#E)

# FONT

Equivalent to clicking the Font command on the Options menu in Microsoft
Excel for the Macintosh version 1.5 or earlier. This function is
included only for macro compatibility. Sets the font for the Normal
style. Microsoft Excel now uses the FONT.PROPERTIES and DEFINE.STYLE
functions. For more information, see FONT.PROPERTIES and DEFINE.STYLE.

**Syntax**

**FONT**(**name\_text, size\_num**)

**FONT**?(name\_text, size\_num)

**Related Functions**

DEFINE.STYLE   Creates or changes a cell style

FONT.PROPERTIES   Sets various font properties

Return to [top](#E)

# FONT.PROPERTIES

Equivalent to choosing the Cells command from the Format menu. Applies a
font and other attributes to the selection. Applies to cells, charts,
and text boxes and buttons on worksheets and macro sheets.

**Syntax**

**FONT.PROPERTIES**(font, font\_style, size, strikethrough, superscript,
subscript, outline, shadow, underline, color, normal, background,
start\_char, char\_count)

**FONT.PROPERTIES**?(font, font\_style, size, strikethrough,
superscript, subscript, outline, shadow, underline, color, normal,
background, start\_char, char\_count)

Arguments correspond to check boxes or options in the Font tab on the
Format Cells dialog box. Arguments that correspond to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box. If an
argument is omitted, the format is not changed.

Font    is the name of the font as it appears on the Font tab. For
example, Courier is a font name.

Font\_style    is the name of the font style as it appears on the Font
tab. For example, Bold Italic is a font style.

Size    is the font size, in points.

Strikethrough    corresponds to the Strikethrough check box.

Superscript    corresponds to the Superscript check box

Subscript    corresponds to the Subscript check box

Outline    corresponds to the Outline check box. Outline fonts are
available in Microsoft Excel for the Macintosh. For macro compatibility,
this argument is ignored by Microsoft Excel for Windows..

Shadow    corresponds to the Shadow check box. Shadow fonts are
available in Microsoft Excel for the Macintosh. For macro compatibility,
this argument is ignored by Microsoft Excel for Windows.

**Note   **For macro compatibility with Microsoft Excel for the
Macintosh, the presence of the outline and shadow arguments do not
prevent the macro from working on Microsoft Excel for Windows, nor does
their absence prevent it from working on the Macintosh.

Underline    corresponds to the Underline Drop-down box.

|               |                   |
| ------------- | ----------------- |
| **Underline** | **Type applied**  |
| 0             | None              |
| 1             | Single            |
| 2             | Double            |
| 3             | Single Accounting |
| 4             | Double Accounting |

Color    is a number from 0 to 56 corresponding to the colors listed in
the Color box; 0 corresponds to automatic color.

Normal    corresponds to the Normal Font check box. Applies the default
font for your system

Background    is a number from 1 to 3 specifying which type of
background to apply to text in a chart.

|                |                                |
| -------------- | ------------------------------ |
| **Background** | **Type of background applied** |
| 1              | Automatic                      |
| 2              | Transparent                    |
| 3              | Opaque                         |

Start\_char    specifies the first character to be formatted. If
start\_char is omitted, it is assumed to be 1 (the first character in
the cell or text box).

Char\_count    specifies how many characters to format. If char\_count
is omitted, Microsoft Excel formats all characters in the cell or text
box starting at start\_char.

**Remarks**

Some extended TrueType styles do not have corresponding arguments to
FONT.PROPERTIES. To access an extended TrueType font style, append the
style name to the font name in the font argument. For example, the font
Taipei can be formatted in an upside-down style by specifying "Taipei
Upside-down" as the font argument. For more information about TrueType,
see your Microsoft Windows documentation.

**Related Functions**

ALIGNMENT   Aligns or wraps text in cells

FORMAT.NUMBER   Applies a number format to the selection

FORMAT.TEXT   Formats a worksheet text box or a chart text item

Return to [top](#E)

# FOPEN

Opens a file with the type of permission specified. Unlike OPEN, FOPEN
does not load the file into memory and display it; instead, FOPEN
establishes a channel with the file so that you can exchange information
with it. If the file is opened successfully, FOPEN returns a file ID
number. If it can't open the file, FOPEN returns the \#N/A error value.
Use the file ID number with other file functions (such as FREAD, FWRITE,
and FSIZE) when you want to get information from or send information to
the file.

**Syntax**

**FOPEN**(**file\_text**, access\_num)

File\_text    is the name of the file you want to open.

Access\_num    is a number from 1 to 3 specifying what type of
permission to allow to the file:

|                 |                                                                       |
| --------------- | --------------------------------------------------------------------- |
| **Access\_num** | **Type of permission**                                                |
| 1 or omitted    | Can read and write to the file (read/write permission)                |
| 2               | Can read the file, but can't write to the file (read-only permission) |
| 3               | Creates a new file with read/write permission                         |

 

  - > If the file doesn't exist and access\_num is 3, FOPEN creates a
    > new file.

  - > If the file does exist and access\_num is 3, FOPEN replaces the
    > contents of the file with any information you supply using the
    > FWRITE or FWRITELN functions.

  - > If the file doesn't exist and access\_num is 1 or 2, FOPEN returns
    > the \#N/A error value.

>  

**Remarks**

Use FCLOSE to close a file after you finish using it.

**Example**

The following function opens a file identified as FileName using
read-only permission:

FOPEN(FileName, 2)

**Related Functions**

FCLOSE   Closes a text file

FREAD   Reads characters from a text file

FWRITE   Writes characters to a text file

OPEN   Opens a workbook

Return to [top](#E)

# FOR

Starts a FOR-NEXT loop. The instructions between FOR and NEXT are
repeated until the loop counter reaches a specified value. Use FOR when
you need to repeat instructions a specified number of times. Use
FOR.CELL when you need to repeat instructions over a range of cells.

**Syntax**

**FOR**(**counter\_text, start\_num, end\_num**, step\_num)

Counter\_text    is the name of the loop counter in the form of text.

Start\_num    is the value initially assigned to counter\_text.

End\_num    is the last value assigned to counter\_text.

Step\_num    is a value added to the loop counter after each iteration.
If step\_num is omitted, it is assumed to be 1.

**Remarks**

  - > Microsoft Excel follows these steps as it executes a FOR-NEXT
    > loop:

<table>
<tbody>
<tr class="odd">
<td><strong>Step</strong></td>
<td><strong>Action</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Sets counter_text to the value start_num.</td>
</tr>
<tr class="odd">
<td>2</td>
<td><p>If counter_text is greater than end_num (or less than end_num if step_num is negative), the loop ends, and the macro continues with the function after the NEXT function.</p>
<p>If counter_text is less than or equal to end_num (or greater than or equal to end_num if step_num is negative), the macro continues in the loop.</p></td>
</tr>
<tr class="even">
<td>3</td>
<td>Carries out functions up to the following NEXT function. The NEXT function must be below the FOR function and in the same column.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Adds step_num to the loop counter.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns to the FOR function and proceeds as described in step 2.</td>
</tr>
</tbody>
</table>

 

  - > You can interrupt a FOR-NEXT loop by using the BREAK function.

>  

**Example**

The following macro starts a FOR-NEXT loop that is executed once for
every open window:

FOR("Counter", 1, COLUMNS(WINDOWS()))

**Related Functions**

BREAK   Interrupts a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

FOR.CELL   Starts a FOR.CELL-NEXT loop

NEXT   Ends a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

WHILE   Starts a WHILE-NEXT loop

Return to [top](#E)

# FOR.CELL

Starts a FOR.CELL-NEXT loop. This function is similar to FOR, except
that the instructions between FOR.CELL and NEXT are repeated over a
range of cells, one cell at a time, and there is no loop counter.

**Syntax**

**FOR.CELL**(**ref\_name**, area\_ref, skip\_blanks)

Ref\_name    is the name in the form of text that Microsoft Excel gives
to the one cell in the range that is currently being operated on;
ref\_name refers to a new cell during each loop.

Area\_ref    is the range of cells on which you want the FOR.CELL-NEXT
loop to operate and can be a multiple selection. If area\_ref is
omitted, it is assumed to be the current selection.

Skip\_blanks    is a logical value specifying whether Microsoft Excel
skips blank cells as it operates on the cells in area\_ref.

|                  |                                    |
| ---------------- | ---------------------------------- |
| **Skip\_blanks** | **Result**                         |
| TRUE             | Skips blank cells in area\_ref     |
| FALSE or omitted | Operates on all cells in area\_ref |

**Remarks**

FOR.CELL operates on each cell in a row from left to right one area at a
time before moving to the next row in the selection.

**Example**

The following macro starts a FOR.CELL-NEXT loop and uses the name
CurrentCell to refer to the cell in the range that is currently being
operated on:

FOR.CELL("CurrentCell", SELECTION(), TRUE)

**Related Functions**

BREAK   Interrupts a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

FOR   Starts a FOR-NEXT loop

NEXT   Ends a FOR-NEXT, FOR.CELL-NEXT, or WHILE-NEXT loop

WHILE   Starts a WHILE-NEXT loop

Return to [top](#E)

# FORMAT.AUTO

Equivalent to clicking the AutoFormat command on the Format menu when a
worksheet is active or clicking the AutoFormat button. Formats the
selected range of cells from a built-in gallery of formats.

**Syntax**

**FORMAT.AUTO**(format\_num, number, font, alignment, border, pattern,
width)

**FORMAT.AUTO**?(format\_num, number, font, alignment, border, pattern,
width)

Format\_num    is a number from 1 to 17 corresponding to the formats in
the Table Format list box in the AutoFormat dialog box.

|                 |                                                     |
| --------------- | --------------------------------------------------- |
| **Format\_num** | **Table Format**                                    |
| 0               | None                                                |
| 1 or omitted    | Classic 1                                           |
| 2               | Classic 2                                           |
| 3               | Classic 3                                           |
| 4               | Accounting 1                                        |
| 5               | Accounting 2                                        |
| 6               | Accounting 3                                        |
| 7               | Colorful 1                                          |
| 8               | Colorful 2                                          |
| 9               | Colorful 3                                          |
| 10              | List 1                                              |
| 11              | List 2                                              |
| 12              | List 3                                              |
| 13              | 3D Effects 1                                        |
| 14              | 3D Effects 2                                        |
| 15              | Japan 1 (Far East versions of Microsoft Excel only) |
| 16              | Japan 2 (Far East versions of Microsoft Excel only) |
| 17              | Accounting 4                                        |
| 18              | Simple                                              |

The following arguments are logical values corresponding to the Formats
To Apply check boxes in the AutoFormat dialog box. If an argument is
TRUE or omitted, Microsoft Excel selects the check box; if FALSE,
Microsoft Excel clears the check box.

Number    corresponds to the Number check box.

Font    corresponds to the Font check box.

Alignment    corresponds to the Alignment check box.

Border    corresponds to the Border check box.

Pattern    corresponds to the Pattern check box.

Width    corresponds to the Column Width/Row Height check box.

**Related Functions**

ALIGNMENT   Aligns or wraps text in cells

BORDER   Adds a border to the selected cell or object

FONT.PROPERTIES   Applies a font to the selection

FORMAT.NUMBER   Applies a number format to the selection

PATTERNS   Changes the appearance of the selected object

Return to [top](#E)

# FORMAT.CHART

Equivalent to choosing the Options button in the Chart Type dialog box,
which is available when you choose the Chart Type command from the
Format menu when a chart is active. Formats the chart according to the
arguments you specify.

**Syntax**

**FORMAT.CHART**(layer\_num, view, overlap, angle, gap\_width,
gap\_depth, chart\_depth, doughnut\_size, axis\_num, drop, hilo,
up\_down, series\_line, labels, vary)

**FORMAT.CHART**?(layer\_num, view, overlap, angle, gap\_width,
gap\_depth, chart\_depth, doughnut\_size, axis\_num, drop, hilo,
up\_down, series\_line, labels, vary)

Several of the following arguments are logical values corresponding to
check boxes in the Options tab of Format (chart type) Group dialog box.
If an argument is TRUE, Microsoft Excel selects the corresponding check
box; if FALSE, Microsoft Excel clears the check box. If an argument is
omitted, the setting is unchanged.

Layer\_num    is a number specifying which chart you want to change.

View    is a number specifying one of the subtypes in the Subtype tab of
the Format (type) Group dialog box. The subtype varies depending on the
type of chart.

Overlap    is a number from -100 to 100 specifying how you want bars or
columns to be positioned. It corresponds to the Overlap edit box in the
Options tab on the Format Bar Group Dialog box, which appears when you
choose the Bar Group from the Format menu. Overlap is ignored if
type\_num is not 2 or 3 (bar or column chart).

  - > If overlap is positive, it specifies the percentage of overlap you
    > want for bars or columns. For example, 50 would cause one-half of
    > a bar or column to be covered by an adjacent bar or column. A
    > value of zero prevents bars or columns from overlapping.

  - > If overlap is negative, then bars or columns are separated by the
    > specified percentage of the maximum available distance between any
    > two bars or columns.

  - > If overlap is omitted, it is assumed to be 0 (bars or columns do
    > not overlap), or it is unchanged if a value was previously set.

Angle    is a number from 0 to 360 specifying the angle of the first pie
or doughnut slice (in degrees) if the chart is a pie or doughnut chart.
If angle is omitted, it is assumed to be 0, or it is unchanged if a
value was previously set.

Gap\_width    is a number from 0 to 500 specifying the space between bar
or column clusters as a percentage of the width of a bar or column. It
corresponds to the Gap Width edit box in the Options tab on the Format
Bar Group Dialog box, which appears when you choose the Bar Group from
the Format menu.

  - > Gap\_width is ignored if type\_num is not 2, 3, 8, or 12 (bar or
    > column chart).

  - > If Gap\_width is omitted, it is assumed to be 50, or it is
    > unchanged if a value was previously set.

>  

The next two arguments are for 3-D charts only, and correspond to check
boxes in the Options tab of Format (chart type) Group dialog box.

Gap\_depth    is a number from 0 to 500 specifying the depth of the gap
in front of and behind a bar, column, area, or line as a percentage of
the depth of the bar, column, area, or line.

  - > Gap\_depth is ignored if the chart is a pie chart or if it is not
    > a 3-D chart.

  - > If gap\_depth is omitted and the chart is a 3-D chart, gap\_depth
    > is assumed to be 50, or it is unchanged if a value was previously
    > set. If gap\_depth is omitted and the view is side-by-side,
    > stacked, or stacked 100%, gap\_depth is assumed to be 0, or it is
    > unchanged if a value was previously set.

>  

Chart\_depth    is a number from 20 to 2000 specifying the visual depth
of the chart as a percentage of the width of the chart.

  - > Chart\_depth is ignored if the chart is not a 3-D chart.

  - > If Chart\_depth is omitted, it is assumed to be 100, or it is
    > unchanged if a value was previously set.

>  

Doughnut\_size    specifies the size of the hole in a doughnut chart.
Can be a value from 10% - 90%. Default is 50%.

Axis\_num    is a number specifying whether to plot the chart on the
primary axis or the secondary axis.

Drop    corresponds to the Drop Lines check box. Drop is available only
for area and line charts.

Hilo    corresponds to the Hi-Lo Lines check box. Hilo is available only
for 2-D line charts.

The next four arguments are logical values corresponding to check boxes
in the Options tab of the Format (chart type) Group dialog box. If an
argument is TRUE, Microsoft Excel selects the corresponding check box;
if FALSE, Microsoft Excel clears the check box. If an argument is
omitted, the setting is unchanged.

Up\_down    corresponds to the Up/Down Bars check box. Up\_down is
available only for 2-D line charts.

Series\_line    corresponds to the Series Lines check box. Series\_line
is available only for 2-D stacked bar and column charts.

Labels    corresponds to the Radar Axis Labels check box. Labels is
available only for radar charts.

Vary    corresponds to the Vary Colors By Point check box. Vary applies
only to charts with one data series and is not available for area
charts.

**Related Functions**

FORMAT.MAIN   Formats a chart according to the arguments you specify

FORMAT.OVERLAY   Formats an overlay chart

Return to [top](#E)

# FORMAT.CHARTTYPE

Changes the chart type for a selected data series, a group of data
series, or an entire chart.

**Syntax**

**FORMAT.CHARTTYPE**(**apply\_to**, group\_num, dimension,
**type\_num**)

**FORMAT.CHARTTYPE**?(apply\_to, group\_num, dimension, type\_num)

Apply\_to    is a number from 1 to 3 specifying what part of a chart the
new chart type effects.

|           |                      |
| --------- | -------------------- |
| **Value** | **Part of chart**    |
| 1         | Selected data series |
| 2         | Group of data series |
| 3         | Entire chart         |

Group\_num    corresponds to the number of the group you want to change
as listed in the Group list box of the Chart Type dialog box, which
appears when you click Chart Type from the Format menu while a chart is
active. Groups are numbered starting with 1 for the group at the top of
the list. This argument is required if apply\_to equals 2; otherwise it
is ignored.

Dimension    specifies whether to apply a 2-D or 3-D chart type. Use 1
for a 2-D chart type or 2 for a 3-D chart type. If omitted, uses the
same dimension as the series, group, or chart to be changed.

Type\_num    specifies the chart type to apply. Meaning of type\_num
varies depending on the value of dimension:

|               |                                  |
| ------------- | -------------------------------- |
| **Type\_num** | **Chart type if dimension is 1** |
| 1             | Area or 3-D area                 |
| 2             | Bar or 3-D bar                   |
| 3             | Column or 3-D column             |
| 4             | Line or 3-D line                 |
| 5             | Pie or 3-D pie                   |
| 6             | Doughnut or 3-D surface          |
| 7             | Radar                            |
| 8             | XY (scatter)                     |

**Related Function**

FORMAT.CHART   Formats the selected chart

Return to [top](#E)

# FORMAT.FONT

Equivalent to choosing the Cells command from the Format menu, and then
selecting Font tab from the Format Cells dialog box. This function is
included for compatibility with Microsoft Excel version 4.0. Use
FONT.PROPERTIES to set various font properties. FORMAT.FONT has three
syntax forms. Syntax 1 is for cells; syntax 2 is for text boxes and
buttons; syntax 3 is used with all chart items (axes, labels, text, and
so on).

**Syntax 1**

Cells

**FORMAT.FONT**(name\_text, size\_num, bold, italic, underline, strike,
color, outline, shadow)

**FORMAT.FONT**?(name\_text, size\_num, bold, italic, underline, strike,
color, outline, shadow)

**Syntax 2**

Text boxes and buttons on worksheets and macro sheets

**FORMAT.FONT**(name\_text, size\_num, bold, italic, underline, strike,
color, outline, shadow, object\_id\_text, start\_num, char\_num)

**FORMAT.FONT**?(name\_text, size\_num, bold, italic, underline, strike,
color, outline, shadow, object\_id\_text, start\_num, char\_num)

**Syntax 3**

Chart items including unattached chart text

**FORMAT.FONT**(color, backgd, apply, name\_text, size\_num, bold,
italic, underline, strike, outline, shadow, object\_id\_text,
start\_num, char\_num)

**FORMAT.FONT**?(color, backgd, apply, name\_text, size\_num, bold,
italic, underline, strike, outline, shadow, object\_id\_text,
start\_num, char\_num)

Arguments correspond to check boxes and list box items in the Font tab
on the Format Cells dialog box. Arguments that correspond to check boxes
are logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box. If an
argument is omitted, the format is not changed.

Name\_text    is the name of the font as it appears in the Font tab. For
example, Courier is a font name.

Size\_num    is the font size, in points.

Bold    corresponds to the Bold item in the Font Style list box. Makes
the selection bold, if applicable.

Italic    corresponds to the Italic item in the Font Style list box.
Makes the selection italic, if applicable.

Underline    corresponds to the Underline check box.

Strike    corresponds to the Strikethrough check box.

Color    is a number from 0 to 56 corresponding to the colors in the
Font tab; 0 corresponds to automatic color.

Outline    corresponds to the Outline check box. Outline fonts are
available in Microsoft Excel for the Macintosh. For macro compatibility,
this argument is ignored by Microsoft Excel for Windows.

Shadow    corresponds to the Shadow check box. Shadow fonts are
available in Microsoft Excel for the Macintosh. For macro compatibility,
this argument is ignored by Microsoft Excel for Windows.

**Note   **For macro compatibility with Microsoft Excel for the
Macintosh, the presence of the outline and shadow arguments do not
prevent the macro from working on Microsoft Excel for Windows, nor does
their absence prevent it from working on the Macintosh.

Object\_id\_text    identifies the text box you want to format (for
example, "Text 1", "Text 2", and so on). You can also use the object
number alone without the text identifier. For compatibility with earlier
versions of Microsoft Excel. This argument is ignored in Microsoft Excel
version 5.0 or later.

Start\_num    specifies the first character to be formatted. If
start\_num is omitted, it is assumed to be 1 (the first character in the
text box).

Char\_num    specifies how many characters to format. If char\_num is
omitted, Microsoft Excel formats all characters in the text box starting
at start\_num.

Backgd    is a number from 1 to 3 specifying which type of background to
apply to text in a chart.

|            |                                |
| ---------- | ------------------------------ |
| **Backgd** | **Type of background applied** |
| 1          | Automatic                      |
| 2          | Transparent                    |
| 3          | Opaque                         |

Apply    corresponds to the Apply To All check box. This argument
applies to data labels only.

**Remarks**

Some extended TrueType styles do not have corresponding arguments to
FORMAT.FONT. To access an extended TrueType font style, append the style
name to the font name in name\_text. For example, the font Taipei can be
formatted in an upside-down style by specifying "Taipei Upside-down" as
the name\_text argument. For more information about TrueType, see your
Microsoft Windows documentation.

**Related Functions**

ALIGNMENT   Aligns or wraps text in cells

FONT.PROPERTIES   Sets various font attributes

FORMAT.NUMBER   Applies a number format to the selection

FORMAT.TEXT   Formats a worksheet text box or a chart text item

Return to [top](#E)

# FORMAT.LEGEND

Equivalent to clicking the Selected Legend command on the Format menu
when a chart is active. Determines the position and orientation of the
legend on a chart and returns TRUE; returns an error message if the
legend is not already selected.

**Syntax**

**FORMAT.LEGEND**(**position\_num**)

**FORMAT.LEGEND**?(position\_num)

Position\_num    is a number from 1 to 5 specifying the position of the
legend.

|                   |                        |
| ----------------- | ---------------------- |
| **Position\_num** | **Position of legend** |
| 1                 | Bottom                 |
| 2                 | Corner                 |
| 3                 | Top                    |
| 4                 | Right                  |
| 5                 | Left                   |

**Related Functions**

FORMAT.MOVE   Moves the selected object

FORMAT.SIZE   Sizes an object

LEGEND   Adds or deletes a chart legend

Return to [top](#E)

# FORMAT.MAIN

Equivalent to clicking the Main Chart command on the Format menu in
Microsoft Excel version 4.0. Formats a chart according to the arguments
you specify. This function is included for compatibility with Microsoft
Excel version 4.0. In Microsoft Excel version 5.0 or later, this is
equivalent to clicking the Chart Type command on the Format menu. You
can also use the FORMAT.CHART function.

**Syntax**

**FORMAT.MAIN**(**type\_num**, view, overlap, gap\_width, vary, drop,
hilo, angle, gap\_depth, chart\_depth, up\_down, series\_line, labels,
doughnut\_size)

**FORMAT.MAIN**?(type\_num, view, overlap, gap\_width, vary, drop, hilo,
angle, gap\_depth, chart\_depth, up\_down, series\_line, labels,
doughnut\_size)

Type\_num    is a number specifying the type of chart.

|               |              |
| ------------- | ------------ |
| **Type\_num** | **Chart**    |
| 1             | Area         |
| 2             | Bar          |
| 3             | Column       |
| 4             | Line         |
| 5             | Pie          |
| 6             | XY (Scatter) |
| 7             | 3-D Area     |
| 8             | 3-D Column   |
| 9             | 3-D Line     |
| 10            | 3-D Pie      |
| 11            | Radar        |
| 12            | 3-D Bar      |
| 13            | 3-D Surface  |
| 14            | Doughnut     |

View    is a number specifying one of the views in the Data View box in
the Main Chart dialog box. The view varies depending on the type of
chart.

Overlap    is a number from -100 to 100 specifying how you want bars or
columns to be positioned. It corresponds to the Overlap box in the Main
Chart dialog box. Overlap is ignored if type\_num is not 2 or 3 (bar or
column chart).

  - > If overlap is positive, it specifies the percentage of overlap you
    > want for bars or columns. For example, 50 would cause one-half of
    > a bar or column to be covered by an adjacent bar or column. A
    > value of zero prevents bars or columns from overlapping.

  - > If overlap is negative, then bars or columns are separated by the
    > specified percentage of the maximum available distance between any
    > two bars or columns.

  - > If overlap is omitted, it is assumed to be 0 (bars or columns do
    > not overlap), or it is unchanged if a value was previously set.

Gap\_width    is a number from 0 to 500 specifying the space between bar
or column clusters as a percentage of the width of a bar or column. It
corresponds to the Gap Width box in the Main Chart dialog box.

  - > Gap\_width is ignored if type\_num is not 2, 3, 8, or 12 (bar or
    > column chart).

  - > If gap\_width is omitted, it is assumed to be 50, or it is
    > unchanged if a value was previously set.

Several of the following arguments are logical values corresponding to
check boxes in the Main Chart dialog box. If an argument is TRUE,
Microsoft Excel selects the corresponding check box; if FALSE, Microsoft
Excel clears the check box. If an argument is omitted, the setting is
unchanged.

Vary    corresponds to the Vary By Categories check box. Vary applies
only to charts with one data series and is not available for area
charts.

Drop    corresponds to the Drop Lines check box. Drop is available only
for area and line charts.

Hilo    corresponds to the Hi-Lo Lines check box. Hilo is available only
for line charts.

Angle    is a number from 0 to 360 specifying the angle of the first pie
slice (in degrees) if the chart is a pie chart. If angle is omitted, it
is assumed to be 0, or it is unchanged if a value was previously set.

The next two arguments are for 3-D charts only.

Gap\_depth    is a number from 0 to 500 specifying the depth of the gap
in front of and behind a bar, column, area, or line as a percentage of
the depth of the bar, column, area, or line.

  - > Gap\_depth is ignored if the chart is a pie chart or if it is not
    > a 3-D chart.

  - > If gap\_depth is omitted and the chart is a 3-D chart, gap\_depth
    > is assumed to be 50, or it is unchanged if a value was previously
    > set. If gap\_depth is omitted and the view is side-by-side,
    > stacked, or stacked 100%, gap\_depth is assumed to be 0, or it is
    > unchanged if a value was previously set.

>  

Chart\_depth    is a number from 20 to 2000 specifying the visual depth
of the chart as a percentage of the width of the chart. Chart\_depth
corresponds to the Chart Depth box in the Main Chart dialog box.

  - > Chart\_depth is ignored if the chart is not a 3-D chart.

  - > If chart\_depth is omitted, it is assumed to be 100, or it is
    > unchanged if a value was previously set.

>  

The next three arguments are logical values corresponding to check boxes
in the Main Chart dialog box. If an argument is TRUE, Microsoft Excel
selects the corresponding check box; if FALSE, Microsoft Excel clears
the check box. If an argument is omitted, the setting is unchanged. The
final argument is for compatibility with Microsoft Excel version 4.0.

Up\_down    corresponds to the Up/Down Bars check box. Up\_down is
available only for line charts.

Series\_line    corresponds to the Series Lines check box. Series\_line
is available only for stacked bar and column charts.

Labels    corresponds to the Radar Axis Labels check box. Labels is
available only for radar charts.

Doughnut\_size    specifies the size of the hole in a doughnut chart.
Can be a value from 10% - 90%. Default is 50%

**Related Functions**

FORMAT.CHART   Formats a chart

FORMAT.OVERLAY   Formats an overlay chart

Return to [top](#E)

# FORMAT.MOVE

Equivalent to moving an object with the mouse. Moves the selected object
to the specified position and, if successful, returns TRUE. If the
selected object cannot be moved, FORMAT.MOVE returns FALSE. There are
three syntax forms of this function. Use syntax 1 to move worksheet
objects. Use syntax 2 to move chart items. Use syntax 3 to move
pie-chart and doughnut-chart items. It is generally easier to use the
macro recorder to enter this function on your macro sheet.

Syntax 1   Moves worksheet items

Syntax 2   Moves chart items

Syntax 3   Moves pie-chart and doughnut-chart items

Return to [top](#E)

# FORMAT.MOVE Syntax 1

Equivalent to moving an object with the mouse. Moves the selected object
to the specified position and, if successful, returns TRUE. If the
selected object cannot be moved, FORMAT.MOVE returns FALSE. There are
three syntax forms of this function. Use syntax 1 to move worksheet
objects. Use syntax 2 to move chart items. Use syntax 3 to move
pie-chart and doughnut-chart items. It is generally easier to use the
macro recorder to enter this function on your macro sheet.

**Syntax**

**FORMAT.MOVE**(**x\_offset**, **y\_offset**, reference)

**FORMAT.MOVE**?(x\_offset, y\_offset, reference)

X\_offset    specifies the horizontal position to which you want to move
the object and is measured in points from the upper-left corner of the
object to the upper-left corner of the cell specified by reference. A
point is 1/72nd of an inch.

Y\_offset    specifies the vertical position to which you want to move
the object and is measured in points from the upper-left corner of the
object to the upper-left corner of the cell specified by reference.

Reference    specifies which cell or range of cells to place the object
in relation to.

  - > If reference is a range of cells, only the upper-left cell is
    > used.

  - > If reference is omitted, it is assumed to be cell A1.

>  

**Remarks**

The position of an object is based on its upper-left corner. For ovals
and arcs, the position is based on the upper-left corner of the bounding
rectangle of the object.

**Example**

The following macro formula moves an object on the active worksheet so
that it is 10 points horizontally offset and 15 points vertically offset
from cell D4:

FORMAT.MOVE(10, 15, \!$D$4)

**Related Functions**

CREATE.OBJECT   Creates an object

FORMAT.SIZE   Sizes an object

WINDOW.MOVE   Moves a window

Syntax 2   Moves chart items

Syntax 3   Moves pie-chart and doughnut-chart items

Return to [top](#E)

# FORMAT.MOVE Syntax 2

Equivalent to moving an object with the mouse. Moves the base of the
selected object to the specified position and, if successful, returns
TRUE. If the selected object cannot be moved, FORMAT.MOVE returns FALSE.
There are three syntax forms of this function. Use syntax 3 to move
pie-chart and doughnut-chart items. Use syntax 1 to move worksheet
objects. It is generally easier to use the macro recorder to enter this
function on your macro sheet.

**Syntax**

**FORMAT.MOVE**(**x\_pos**, **y\_pos**)

**FORMAT.MOVE**?(x\_pos, y\_pos)

X\_pos    specifies the horizontal position to which you want to move
the object and is measured in points from the base of the object to the
lower-left corner of the window. A point is 1/72nd of an inch.

Y\_pos    specifies the vertical position to which you want to move the
object and is measured in points from the base of the object to the
lower-left corner of the window.

**Remarks**

  - > The base of a text label on a chart is the lower-left corner of
    > the text rectangle.

  - > The base of an arrow is the end without the arrowhead.

  - > The base of a pie slice is the point.

>  

**Example**

On a chart, the following macro formula moves the base of the selected
chart object 10 points to the right of and 20 points above the
lower-left corner of the window:

FORMAT.MOVE(10, 20)

**Related Functions**

FORMAT.SIZE   Sizes an object

WINDOW.MOVE   Moves a window

Syntax 1   Moves worksheet items

Syntax 3   Moves pie-chart and doughnut-chart items

Return to [top](#E)

# FORMAT.MOVE Syntax 3

Equivalent to exploding by moving a pie-chart or doughnut-chart slice
with the mouse. Sets the percentage of pie-chart or doughnut-chart slice
explosion, and, if successful, returns TRUE. If the selected object
cannot be exploded, returns FALSE. There are three syntax forms of this
function. Use syntax 1 to move worksheet items. Use syntax 2 to move
chart items. It is generally easier to use the macro recorder to enter
this function on your macro sheet.

**Syntax**

**FORMAT.MOVE**(**explosion\_num**)

Explosion\_num    is a number specifying the explosion percentage for
the selected pie slice or the entire chart (if the series is selected).
Zero is no explosion (the tip of the slice is in the center of the pie).

**Related Functions**

FORMAT.SIZE   Sizes an object

Syntax 1   Moves worksheet items

Syntax 2   Moves chart items

WINDOW.MOVE   Moves a window

Return to [top](#E)

# FORMAT.NUMBER

Equivalent to choosing the Number tab in the Format Cells dialog box,
which appears when you choose Cells from the Format menu. Formats
numbers, dates, and times in the selected cells, data labels, and axis
labels on charts. Use FORMAT.NUMBER to apply built-in formats or to
create and apply custom formats.

**Syntax**

**FORMAT.NUMBER**(**format\_text**)

**FORMAT.NUMBER**?(format\_text)

Format\_text    is a format string, such as "\#, \#\#0.00", specifying
which format to apply to the selection.

**Related Functions**

DELETE.FORMAT   Deletes the specified custom number format

FONT.PROPERTIES   Applies a font to the selection

FORMAT.TEXT   Formats a sheet text box or a chart text item

Return to [top](#E)

# FORMAT.OVERLAY

Equivalent to clicking the Overlay command on the Format menu in
Microsoft Excel version 4.0. Formats the overlay chart according to the
arguments you specify.

**Syntax**

**FORMAT.OVERLAY**(**type\_num**, view, overlap, gap\_width, vary, drop,
hilo, angle, series\_dist, series\_num, up\_down, series\_line, labels)

**FORMAT.OVERLAY**?(type\_num, view, overlap, gap\_width, vary, drop,
hilo, angle, series\_dist, series\_num, up\_down, series\_line, labels)

Type\_num    is a number specifying the type of chart.

|               |              |
| ------------- | ------------ |
| **Type\_num** | **Chart**    |
| 1             | Area         |
| 2             | Bar          |
| 3             | Column       |
| 4             | Line         |
| 5             | Pie          |
| 6             | XY (Scatter) |
| 11            | Radar        |
| 14            | Doughnut     |

View    is a number specifying one of the views in the Data View box in
the Overlay dialog box. The view varies depending on the type of chart.

Overlap    is a number from -100 to 100 specifying how you want bars or
columns to be positioned. It corresponds to the Overlap box in the
Overlay dialog box. Overlap is ignored if type\_num is not 2 or 3 (bar
or column chart).

  - > If overlap is positive, it specifies the percentage of overlap you
    > want for bars or columns. For example, 50 would cause one-half of
    > a bar or column to be covered by an adjacent bar or column.

  - > If overlap is negative, then bars or columns are separated by the
    > specified percentage of the maximum available distance between any
    > two bars or columns.

  - > If overlap is omitted, it is assumed to be 0 (bars or columns do
    > not overlap), or it is unchanged if a value was previously set.

>  

Gap\_width    is a number from 0 to 500 specifying the space between bar
or column clusters as a percentage of the width of a bar or column.

  - > Gap\_width is ignored if type\_num is not 2 or 3 (bar or column
    > chart).

  - > If gap\_width is omitted, it is assumed to be 50, or it is
    > unchanged if a value was previously set.

>  

Several of the following arguments are logical values corresponding to
check boxes in the Overlay dialog box. If an argument is TRUE, Microsoft
Excel selects the corresponding check box; if FALSE, Microsoft Excel
clears the check box. If an argument is omitted, the setting is
unchanged.

Vary    corresponds to the Vary By Categories check box. Vary is not
available for area charts.

Drop    corresponds to the Drop Lines check box. Drop is available only
for area and line charts.

Hilo    corresponds to the Hi-Lo Lines check box. Hilo is available only
for line charts.

Angle    is a number from 0 to 360 specifying the angle of the first pie
slice (in degrees) if the chart is a pie chart. If angle is omitted, it
is assumed to be 0, or it is unchanged if a value was previously set.

Series\_dist    is the number 1 or 2 and specifies automatic or manual
series distribution.

  - > If series\_dist is 1 or omitted, Microsoft Excel uses automatic
    > series distribution.

  - > If series\_dist is 2, Microsoft Excel uses manual series
    > distribution, and you must specify which series is first in the
    > distribution by using the series\_num argument.

>  

Series\_num    is the number of the first series in the overlay chart
and corresponds to the First Overlay Series box in the Overlay dialog
box. If series\_dist is 1 (automatic series distribution), this argument
is ignored.

Up\_down    corresponds to the Up/Down Bars check box. Up\_down is
available only for line charts.

Series\_line    corresponds to the Series Lines check box. Series\_line
is available only for stacked bar and column charts.

Labels    corresponds to the Radar Axis Labels check box. Labels is
available only for radar charts.

**Related Functions**

DELETE.OVERLAY   Deletes the overlay on a chart

FORMAT.CHART   Formats a chart

Return to [top](#E)

# FORMAT.SHAPE

Equivalent to clicking the reshape button on the Drawing toolbar and
then inserting, moving, or deleting vertices of the selected polygon. A
vertex is a point defined by a pair of coordinates in one row of the
array that defines the polygon. The array is created by CREATE.OBJECT
and EXTEND.POLYGON functions.

**Syntax**

**FORMAT.SHAPE**(**vertex\_num, insert**, reference, x\_offset,
y\_offset)

Vertex\_num    is a number corresponding to the vertex you want to
insert, move, or delete.

Insert    is a logical value specifying whether to insert a vertex, or
move or delete a vertex.

  - > If insert is TRUE, Microsoft Excel inserts a vertex between the
    > vertices vertex\_num and vertex\_num-1. The number of the new
    > vertex then becomes vertex\_num. The number of the vertex
    > previously identified by vertex\_num becomes vertex\_num+1, and so
    > on.

  - > If insert is FALSE, Microsoft Excel deletes the vertex (if the
    > remaining arguments are omitted) or moves the vertex to the
    > position specified by the remaining arguments.

>  

Reference    is the reference from which the vertex you are inserting or
moving is measured; that is, the cell or range of cells to use as the
basis for the x and y offsets.

  - > If reference is a range of cells, only the upper-left cell is
    > used.

  - > If reference is omitted, the vertex is measured from the
    > upper-left corner of the polygon's bounding rectangle.

>  

X\_offset    is the horizontal distance from the upper-left corner of
reference to the vertex. X\_offset is measured in points. A point is
1/72nd of an inch. If reference is omitted, x\_offset specifies the
horizontal distance from the upper-left corner of the polygon bounding
rectangle.

Y\_offset    is the vertical distance from the upper-left corner of
reference to the vertex. Y\_offset is measured in points. If reference
is omitted, y\_offset specifies the vertical distance from the
upper-left corner of the polygon bounding rectangle.

**Remarks**

You cannot delete a vertex if only two vertices remain.

**Examples**

The following macro formula deletes the second vertex of the selected
polygon:

FORMAT.SHAPE(2, FALSE)

The following macro formula moves the thirteenth vertex 6 points to the
right and 4 points below the upper-left corner of cell B5 on the active
worksheet:

FORMAT.SHAPE(13, FALSE, \!$B$5, 6, 4)

The following macro formula inserts a new vertex between vertices 2 and
3. The new vertex is 60 points to the right and 75 points below the
upper-left corner of the polygon's bounding rectangle:

FORMAT.SHAPE(3, TRUE, , 60, 75)

**Related Functions**

CREATE.OBJECT   Creates an object

EXTEND.POLYGON   Adds vertices to a polygon

Return to [top](#E)

# FORMAT.SIZE

Equivalent to sizing an object with the mouse. Sizes the selected object
and returns TRUE. If the selected chart object cannot be sized,
FORMAT.SIZE returns FALSE. There are two syntax forms of this function.
Use syntax 1 to size worksheet objects and chart items absolutely. Use
syntax 2 relative to a cell or range of cells to size only worksheet
objects. It is generally easier to use the macro recorder to enter this
function on your macro sheet.

Syntax 1   Sizes worksheet objects and chart items

Syntax 2   Sizes worksheet objects relative to a cell or range

Return to [top](#E)

# FORMAT.SIZE Syntax 1

Equivalent to sizing an object with the mouse. Sizes the selected object
and returns TRUE. If the selected chart object cannot be sized,
FORMAT.SIZE returns FALSE. There are two syntax forms of this function.
Use syntax 1 to size worksheet objects and chart items absolutely. Use
syntax 2 relative to a cell or range of cells to size only worksheet
objects. It is generally easier to use the macro recorder to enter this
function on your macro sheet.

**Syntax**

**FORMAT.SIZE**(width, height)

**FORMAT.SIZE**?(width, height)

Width    specifies the width of the selected object, measured in points.
A point is 1/72nd of an inch.

Height    specifies the height of the selected object, measured in
points.

You do not always have to use both arguments. For example, if you
specify height and not width, the height changes but the width does not.

**Remarks**

  - > The base of a text label on a chart is the lower-left corner of
    > the text rectangle.

  - > The base of an arrow is the end without the arrowhead.

>  

**Related Functions**

FORMAT.MOVE   Moves the selected object

SIZE   Changes the size of a window

Syntax 2   Sizes worksheet objects relative to a cell or range

Return to [top](#E)

# FORMAT.SIZE Syntax 2

Equivalent to sizing an object with the mouse. Sizes the selected
worksheet object and returns TRUE. If the selected object cannot be
sized, FORMAT.SIZE returns FALSE. There are two syntax forms of this
function. Use syntax 2 to size worksheet objects relative to a cell or
range of cells. Use syntax 1 to size worksheet objects and chart items.
It is generally easier to use the macro recorder to enter this function
on your macro sheet.

**Syntax**

**FORMAT.SIZE**(x\_off, y\_off, reference)

**FORMAT.SIZE**?(x\_off, y\_off, reference)

X\_off    specifies the width of the selected object and is measured in
points from the lower-right corner of the object to the upper-left
corner of reference. A point is 1/72nd of an inch. If omitted, x\_off is
assumed to be 0. If reference is omitted, x\_off specifies the
horizontal size.

Y\_off    specifies the height of the selected object and is measured in
points from the lower-right corner of the object to the upper-left
corner of reference. If omitted, y\_off is assumed to be 0. If reference
is omitted, y\_off specifies the vertical size.

Reference    specifies the cell or range of cells to use as the basis
for the offset and for sizing. If reference is a range of cells, only
the upper-left cell in the range is used.

**Related Functions**

FORMAT.MOVE   Moves the selected object

SIZE   Changes the size of a window

Syntax 1   Sizes worksheet objects and chart items

Return to [top](#E)

# FORMAT.TEXT

Formats the selected worksheet text box or button or any text item on a
chart.

**Syntax**

**FORMAT.TEXT**(x\_align, y\_align, orient\_num, auto\_text, auto\_size,
show\_key, show\_value, add\_indent)

**FORMAT.TEXT**?(x\_align, y\_align, orient\_num, auto\_text,
auto\_size, show\_key, show\_value, add\_indent)

Arguments correspond to check boxes or options in the various tabs on
Format Object dialog box. Arguments that correspond to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box; if omitted,
the current setting is used.

X\_align    is a number from 1 to 4 specifying the horizontal alignment
of the text.

|              |                          |
| ------------ | ------------------------ |
| **X\_align** | **Horizontal alignment** |
| 1            | Left                     |
| 2            | Center                   |
| 3            | Right                    |
| 4            | Justify                  |

Y\_align    is a number from 1 to 4 specifying the vertical alignment of
the text.

|              |                        |
| ------------ | ---------------------- |
| **Y\_align** | **Vertical alignment** |
| 1            | Top                    |
| 2            | Center                 |
| 3            | Bottom                 |
| 4            | Justify                |

Orient\_num    is a number from 0 to 3 specifying the orientation of the
text.

|                 |                      |
| --------------- | -------------------- |
| **Orient\_num** | **Text orientation** |
| 0               | Horizontal           |
| 1               | Vertical             |
| 2               | Upward               |
| 3               | Downward             |

Auto\_text    corresponds to the Automatic Text check box. If the
selected text was created with the Data Labels command from the Insert
menu and later edited, this option restores the original text.
Auto\_text is ignored for text boxes on worksheets and macro sheets.

Auto\_size    corresponds to the Automatic Size check box. If you have
changed the size of the border around the selected text, this option
restores the border to automatic size. Automatic size makes the border
fit exactly around the text no matter how you change the text.

Show\_key    corresponds to the Show Legend Key Next to Label check box
in the Data Labels dialog box. This argument applies only if the
selected text is an attached data label on a chart.

Show\_value    corresponds to the Show Value option button in the Format
Data Labels dialog box. This argument applies only if the selected text
is an attached data label on a chart.

The following list summarizes which arguments apply to each type of text
item.

Add\_indent   This argument is for only Far East versions of Microsoft
Excel.

|                              |                                             |
| ---------------------------- | ------------------------------------------- |
| **Text item**                | **Arguments that apply**                    |
| Worksheet text box or button | X\_align, y\_align, orient\_num, auto\_size |
| Attached data label          | All arguments                               |
| Unattached text label        | X\_align, y\_align, orient\_num, auto\_size |
| Tickmark label               | Orient\_num                                 |

**Related Functions**

CREATE.OBJECT   Creates an object

FONT.PROPERTIES   Applies a font to the selection

FORMULA   Enters values into a cell or range or onto a chart

Return to [top](#E)

# FORMULA

Enters a formula in the active cell or in a reference. There are two
syntax forms of this function. Use syntax 1 to enter numbers, text,
references, and formulas in a worksheet. Although syntax 1 can also be
used to enter values on a macro sheet, you will not generally use
FORMULA for this purpose. Use syntax 2 to enter a formula in a chart.
For information about setting values on a macro sheet, see "Remarks" in
the following topics.

Syntax 1   Enters numbers, text, references, and formulas in a worksheet

Syntax 2   Enters formulas in a chart

Return to [top](#E)

# FORMULA Syntax 1

Enters a formula in the active cell or in a reference. If the active
sheet is a worksheet, using FORMULA is equivalent to entering
formula\_text in the cell specified by reference. Formula\_text is
entered just as if you typed it in the formula bar.

There are two syntax forms of this function. Use syntax 1 to enter
numbers, text, references, and formulas in a worksheet. Although syntax
1 can also be used to enter values on a macro sheet, you will not
generally use FORMULA for this purpose. Use syntax 2 to enter a formula
in a chart. For information about setting values on a macro sheet, see
"Remarks" later in this topic.

**Syntax**

**FORMULA**(**formula\_text**, reference)

Formula\_text    can be text, a number, a reference, or a formula in the
form of text, or a reference to a cell containing any of the above.

  - > If formula\_text contains references, they must be R1C1-style
    > references, such as "=RC\[1\]\*(1+R1C1)". If you are recording a
    > macro when you enter a formula, Microsoft Excel converts A1-style
    > references to R1C1-style references. For example, if you enter the
    > formula =B2\*(1+$A$1) in cell C2 while recording, Microsoft Excel
    > records that action as =FORMULA("=RC\[-1\]\*(1+R1C1)").

  - > If formula\_text is a formula, the formula is entered. Text
    > arguments must be surrounded by double sets of quotation marks.
    > For example, to enter the formula =IF($A$1="Hello World", 1, 0) in
    > the active cell with the FORMULA function, you would use the
    > formula FORMULA("=IF(R1C1=""Hello World"", 1, 0)")

  - > If formula\_text is a number, text, or logical value, the value is
    > entered as a constant.

Reference    specifies where formula\_text is to be entered. It can be a
reference to a cell in the active workbook or an external reference to a
workbook. If reference is omitted, formula\_text is entered in the
active cell.

**Remarks**

Consider the following guidelines as you choose a function to set values
on a worksheet or macro sheet:

  - > Use FORMULA to enter formulas and change values in a worksheet
    > cell.

  - > SET.VALUE changes values on the macro sheet. Use SET.VALUE to
    > assign initial values to a reference and to store values during
    > the calculation of the macro.

  - > SET.NAME creates names on the macro sheet. Use SET.NAME to create
    > a name and immediately assign a value to the name.

>  

**Examples**

If the active sheet is a worksheet, the following macro formula enters
the number constant 523 in the active cell:

FORMULA(523)

If the active sheet is a worksheet, the following macro formula enters
the result of the INPUT function in cell A5:

FORMULA(INPUT("Enter a formula:", 0), \!$A$5)

If you're using R1C1-style references and the active sheet is a
worksheet, the following macro formula enters the formula
=RC\[-1\]\*(1+R1C1) in the active cell:

FORMULA("=RC\[-1\]\*(1+R1C1)")

If the active sheet is a worksheet, the following macro formulas enter
the number 1000 in the cell two rows down and three columns right from
the active cell. The R1C1-style formula is shorter, but the OFFSET
method may provide faster performance in larger macro sheets.

FORMULA(1000, OFFSET(ACTIVE.CELL(), 2, 3))

FORMULA(1000, "R\[2\]C\[3\]")

The following macro formula enters the phrase "Year to Date" in cell B4
on the sheet named SALES 1993:

FORMULA("Year to Date", 'SALES 1993'\!B4)

**Related Functions**

FORMULA.ARRAY   Enters an array

FORMULA.FILL   Enters a formula in the specified range

SET.VALUE   Sets the value of a cell on a macro sheet

FORMULA Syntax 2   Enters formulas in a chart

Return to [top](#E)

# FORMULA Syntax 2

Enters a text label or SERIES formula in a chart. To enter formulas on a
worksheet or macro sheet, use syntax 1 of this function.

**Syntax**

**FORMULA**(**formula\_text**)

Formula\_text    is the text label or SERIES formula you want to enter
into the chart.

|                                                                                                                             |                                                             |
| --------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------- |
| **If**                                                                                                                      | **Then**                                                    |
| Formula\_text can be treated as a text label and the current selection is a text label                                      | The selected text label is replaced with formula\_text.     |
| Formula\_text can be treated as a text label and there is no current selection or the current selection is not a text label | Formula\_text creates a new unattached text label.          |
| Formula\_text can be treated as a SERIES formula and the current selection is a SERIES formula                              | The selected SERIES formula is replaced with formula\_text. |
| Formula\_text can be treated as a SERIES formula and the current selection is not a SERIES formula                          | Formula\_text creates a new SERIES formula.                 |

**Remarks**

You would normally use the EDIT.SERIES function to create or edit a
chart series. For more information, see EDIT.SERIES.

**Example**

The following macro formula enters a SERIES formula on the chart. If the
current selection is a SERIES formula, it is replaced:

FORMULA("=SERIES(""Title"", , {1, 2, 3}, 1)")

**Related Functions**

EDIT.SERIES   Creates or changes a chart series

FORMULA, Syntax 1   Enters numbers, text, references, and formulas in a
worksheet

Return to [top](#E)

# FORMULA.ARRAY

Enters a formula as an array formula in the range specified or in the
current selection. Equivalent to entering an array formula while
pressing CTRL+SHIFT+ENTER in Microsoft Excel for Windows or
COMMAND+ENTER in Microsoft Excel for the Macintosh.

**Syntax**

**FORMULA.ARRAY**(**formula\_text**, reference)

Formula\_text    is the text you want to enter in the array. For more
information on formula\_text, see the first form of FORMULA.

Reference    specifies where formula\_text is entered. It can be a
reference to a cell on the active worksheet or an external reference to
a named workbook. Reference must be a R1C1-style reference in text form.
If reference is omitted, formula\_text is entered in the active cell.

**Examples**

If the selection is D25:E25, the following macro formula enters the
array formula {=D22:E22+D23:E23} in the range D25:E25:

FORMULA.ARRAY("=R\[-3\]C:R\[-3\]C\[1\]+R\[-2\]C:R\[-2\]C\[1\]")

Regardless of the selection, the following macro formula enters the
array formula {=D22:E22+D23:E23} in the range D25:E25:

FORMULA.ARRAY("=R\[-3\]C:R\[-3\]C\[1\]+R\[-2\]C:R\[-2\]C\[1\]",
"R25C4:R25C5")

To use FORMULA.ARRAY to put an array in a specific workbook, specify the
name of the workbook as an external reference in the reference argument.
Using "\[SALES.XLS\]North\!R25C3:R25C4" as the reference argument in the
preceding example would enter the array in cells C25:D25 on the
worksheet named North in the workbook SALES.XLS. Using
"SALES\!R25C3:R25C4" as the reference argument would enter the array in
the same cells in the worksheet named SALES.

**Related Functions**

FORMULA   Enters values into a cell or range or onto a chart

FORMULA.FILL   Enters a formula in the specified range

Return to [top](#E)

# FORMULA.CONVERT

Changes the style and type of references in a formula between A1 and
R1C1 and between relative and absolute. Use FORMULA.CONVERT to convert
references of one style or type to another style or type.

**Syntax**

**FORMULA.CONVERT**(**formula\_text, from\_a1**, to\_a1, to\_ref\_type,
rel\_to\_ref)

Formula\_text    is the formula, given as text, containing the
references you want to change. Formula\_text must be a valid formula,
and an equal sign must be included.

From\_a1    is a logical value specifying whether the references in
formula\_text are in A1 or R1C1 style. If from\_a1 is TRUE, references
are in A1 style; if FALSE, references are in R1C1 style.

To\_a1    is a logical value specifying the form for the references
FORMULA.CONVERT returns. If to\_a1 is TRUE, references are returned in
A1 style; if FALSE, references are returned in R1C1 style. If to\_a1 is
omitted, the reference style is not changed.

To\_ref\_type    is a number from 1 to 4 specifying the reference type
of the returned formula. If to\_ref\_type is omitted, the reference type
is not changed.

|                   |                               |
| ----------------- | ----------------------------- |
| **To\_ref\_type** | **Reference type returned**   |
| 1                 | Absolute                      |
| 2                 | Absolute row, relative column |
| 3                 | Relative row, absolute column |
| 4                 | Relative                      |

Rel\_to\_ref    is an absolute reference that specifies what cell the
relative references are or should be relative to.

**Examples**

Use FORMULA.CONVERT to convert relative references entered by the user
in an INPUT function or custom dialog box into absolute references. The
following macro formula converts the given formula to an absolute,
R1C1-style reference:

FORMULA.CONVERT("=A1:A10", TRUE, FALSE, 1) equals "=R1C1:R10C1"

The following macro formula converts the references in the given formula
to relative, A1-style references:

FORMULA.CONVERT("=SUM(R10C2:R15C2)", FALSE, TRUE, 4) equals
"=SUM(B10:B15)"

**Tip**   To put the converted formula into a cell or range of cells,
use the FORMULA.CONVERT function as the formula\_text argument to the
FORMULA function.

**Related Functions**

ABSREF   Returns the absolute reference of a range of cells to another
range

FORMULA   Enters values into a cell or range or onto a chart

RELREF   Returns a relative reference

Return to [top](#E)

# FORMULA.FILL

Enters a formula in the range specified or in the current selection.
Equivalent to entering a formula in a range of cells while pressing
CTRL+ENTER in Microsoft Excel for Windows or OPTION+ENTER in Microsoft
Excel for the Macintosh.

**Syntax**

**FORMULA.FILL**(**formula\_text**, reference)

Formula\_text    is the text with which you want to fill the range. For
more information on formula\_text, see FORMULA.

Reference    specifies where formula\_text is entered. It can be a
reference to a range in the active worksheet or an external reference to
a named workbook. If omitted, formula\_text is entered in the current
selection.

**Related Functions**

DATA.SERIES   Fills a range of cells with a series of numbers or dates

FORMULA   Enters values into a cell or range or onto a chart

FORMULA.ARRAY   Enters an array

Return to [top](#E)

# FORMULA.FIND

Equivalent to clicking the Find command on the Edit menu. Selects the
next or previous cell containing the specified text and returns TRUE. If
a matching cell is not found, FORMULA.FIND returns FALSE and displays a
message.

**Syntax**

**FORMULA.FIND**(**text, in\_num, at\_num, by\_num**, dir\_num,
match\_case)

**FORMULA.FIND**?(text, in\_num, at\_num, by\_num, dir\_num,
match\_case)

Text    is the text you want to find. Text corresponds to the Find What
box in the Find dialog box.

In\_num    is a number from 1 to 3 specifying where to search.

|             |              |
| ----------- | ------------ |
| **In\_num** | **Searches** |
| 1           | Formulas     |
| 2           | Values       |
| 3           | Notes        |

At\_num    is the number 1 or 2 and specifies whether to find cells
containing only text or also cells containing text within a longer
string of characters.

|             |                                                  |
| ----------- | ------------------------------------------------ |
| **At\_num** | **Searches for text as**                         |
| 1           | A whole string (the only value in the cell)      |
| 2           | Either a whole string or part of a longer string |

By\_num    is the number 1 or 2 and specifies whether to search by rows
or by columns.

|             |                 |
| ----------- | --------------- |
| **By\_num** | **Searches by** |
| 1           | Rows            |
| 2           | Columns         |

Dir\_num    is the number 1 or 2 and specifies whether to search for the
next or previous occurrence of text.

|              |                                 |
| ------------ | ------------------------------- |
| **Dir\_num** | **Searches for**                |
| 1 or omitted | The next occurrence of text     |
| 2            | The previous occurrence of text |

Match\_case    is a logical value corresponding to the Match Case check
box in the Find dialog box. If match\_case is TRUE, Microsoft Excel
matches characters exactly, including uppercase and lowercase; if FALSE
or omitted, matching is not case-sensitive.

**Remarks**

  - > In Microsoft Excel for Windows, the dialog-box form of
    > FORMULA.FIND is equivalent to pressing SHIFT+F5.

  - > If more than one cell is selected when you use FORMULA.FIND,
    > Microsoft Excel searches only that selection.

>  

Return to [top](#E)

# FORMULA.FIND.NEXT, FORMULA.FIND.PREV

Finds the next and previous cells on the worksheet, as specified in the
Find dialog box, and returns TRUE. (To see the Find dialog box, click
Find on the Edit menu.) If a matching cell is not found, the functions
return FALSE. For more information see FORMULA.FIND.

**Syntax**

**FORMULA.FIND.NEXT**( )

**FORMULA.FIND.PREV**( )

**Related Functions**

DATA.FIND   Selects records in a database that match the specified
criteria

FORMULA.FIND   Finds text in a workbook

Return to [top](#E)

# FORMULA.GOTO

Equivalent to clicking the Go To command on the Edit menu or to pressing
F5. Scrolls through the worksheet and selects a named area or reference.
Use FORMULA.GOTO to select a range on any open workbook; use SELECT to
select a range on the active workbook.

**Syntax**

**FORMULA.GOTO**(reference, corner)

**FORMULA.GOTO**?(reference, corner)

Reference    specifies where to scroll and what to select.

  - > Reference should be either an external reference to a workbook, an
    > R1C1-style reference in the form of text (see the second example
    > following), or a name.

  - > If the Go To command has already been carried out, reference is
    > optional. If reference is omitted, it is assumed to be the
    > reference of the cells you selected before the previous Go To
    > command or FORMULA.GOTO macro function was carried out. This
    > feature distinguishes FORMULA.GOTO from SELECT.

>  

Corner    is a logical value that specifies whether to scroll through
the window so that the upper-left cell in reference is in the upper-left
corner of the active window. If corner is TRUE, Microsoft Excel places
reference in the upper-left corner of the window; if FALSE or omitted,
Microsoft Excel scrolls through normally.

**Tip**   Microsoft Excel keeps a list of the cells you've selected with
previous FORMULA.GOTO functions or Go To commands. When you use
FORMULA.GOTO with GET.WORKSPACE(41), which returns a horizontal array of
previous Go To selections, you can backtrack through multiple previous
selections. See the last example below.

**Remarks**

  - > If you are recording a macro when you click the Go To command, the
    > reference you enter in the Reference box of the Go To dialog box
    > is recorded as text in the R1C1 reference style.

  - > If you are recording a macro when you double-click a cell that has
    > precedents on another worksheet, Microsoft Excel records a
    > FORMULA.GOTO function.

**Examples**

Each of the following macro formulas goes to cell A1 on the active
worksheet:

FORMULA.GOTO(\!$A$1)

FORMULA.GOTO("R1C1")

Each of the following macro formulas goes to the cells named Sales on
the active worksheet and scrolls through the worksheet so that the
upper-left corner of Sales is in the upper-left corner of the window:

FORMULA.GOTO(\!Sales, TRUE)

FORMULA.GOTO("Sales", TRUE)

The following macro formula goes to the cells that were selected by the
third most recent FORMULA.GOTO function or Go To command:

FORMULA.GOTO(INDEX(GET.WORKSPACE(41), 1, 3))

**Related Functions**

GOTO   Directs macro execution to another cell

HSCROLL   Horizontally scrolls through a sheet by percentage or by
column or row number

SELECT   Selects a cell, worksheet object, or chart item

VSCROLL   Vertically scrolls through a sheet by percentage or by column
or row number

Return to [top](#E)

# FORMULA.REPLACE

Equivalent to clicking the Replace command on the Edit menu. Finds and
replaces characters in cells on your worksheet.

**Syntax**

**FORMULA.REPLACE**(**find\_text, replace\_text**, look\_at, look\_by,
active\_cell, match\_case)

**FORMULA.REPLACE**?(find\_text, replace\_text, look\_at, look\_by,
active\_cell, match\_case)

Find\_text    is the text you want to find. You can use the wildcard
characters, question mark (?) and asterisk (\*), in find\_text. A
question mark matches any single character; an asterisk matches any
sequence of characters. If you want to find an actual question mark or
asterisk, type a tilde (\~) before the character.

Replace\_text    is the text you want to replace find\_text with.

Look\_at    is a number specifying whether you want find\_text to match
the entire contents of a cell or any string of matching characters.

|              |                                   |
| ------------ | --------------------------------- |
| **Look\_at** | **Looks for find\_text**          |
| 1 or omitted | As the entire contents of a cell  |
| 2            | As part of the contents of a cell |

Look\_by    is a number specifying whether to search horizontally
(through rows) or vertically (through columns).

|              |                          |
| ------------ | ------------------------ |
| **Look\_by** | **Looks for find\_text** |
| 1 or omitted | By rows                  |
| 2            | By columns               |

Active\_cell    is a logical value specifying the cells in which
find\_text is to be replaced.

  - > If active\_cell is TRUE, find\_text is replaced in the active cell
    > only.

  - > If active\_cell is FALSE, find\_text is replaced in the entire
    > selection, or, if the selection is a single cell, in the entire
    > sheet.

>  

Match\_case    is a logical value corresponding to the Match Case check
box in the Replace dialog box. If match\_case is TRUE, Microsoft Excel
selects the check box; if FALSE, Microsoft Excel clears the check box.
If match\_case is omitted, the status of the check box is unchanged.

**Remarks**

  - > In FORMULA.REPLACE?, the dialog-box form of the function, omitted
    > arguments are assumed to be the same arguments used in the
    > previous replace operation. If there was no previous replace
    > operation, omitted text arguments are assumed to be "" (empty
    > text).

  - > The result of FORMULA.REPLACE must be a valid cell entry. For
    > example, you cannot replace "=" with "= =" at the beginning of a
    > formula.

  - > If more than a single cell is selected before you use
    > FORMULA.REPLACE, only the selected cells are searched.

>  

**Related Function**

FORMULA.FIND   Finds text in a workbook

Return to [top](#E)

# FOURIER

Performs a Fourier transform.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**FOURIER**(**inprng**, outrng, inverse, labels)

**FOURIER**?(inprng, outrng, inverse, labels)

Inprng    is the input range. The number of cells in the input range
must be equal to a power of two (2, 4, 8, 16, ...).

Outrng    is the first cell in the output range or the name, as text, of
a new sheet to contain the output table. If FALSE, blank, or omitted,
places the output table in a new workbook.

Inverse    is a logical value. If TRUE, an inverse Fourier transform is
performed. If FALSE or omitted, a forward Fourier transform is
performed.

Labels    is a logical value.

  - > If labels is TRUE, then the first row or column of inprng contains
    > labels.

  - > If labels is FALSE or omitted, all cells in inprng are considered
    > data. Microsoft Excel generates appropriate data labels for the
    > output table.

>  

**Related Function**

SAMPLE   Samples data

Return to [top](#E)

# FPOS

Sets the position of a file. The position of a file is where a character
is read from or written to by an FREAD, FREADLN, FWRITE, or FWRITELN
function. Use FPOS when you want to write characters to or read
characters from specific locations. For example, to append text to the
end of a file, you must set the position to the end of the file;
otherwise, you might accidentally overwrite existing characters in the
file.

**Syntax**

**FPOS**(**file\_num**, position\_num)

File\_num    is the unique ID number of the file for which you want to
set the position. File\_num is returned by a previously executed FOPEN
function. If file\_num is not valid, FPOS returns the \#VALUE\! error
value.

Position\_num    is the location in the file that a character will be
read from or written to.

  - > The first position in a file is 1, the location of the first byte.

  - > The last position in the file is the same as the value returned by
    > FSIZE. For example, the last position in a file with 280 bytes is
    > 280.

  - > If position\_num is omitted, FPOS returns the current position of
    > the file—that is, the number corresponding to where the next
    > character will be read from or written to.

>  

Whenever you read a character from or write a character to a file, the
file's position is automatically incremented.

**Examples**

The following statement starts a loop that executes until the position
in the open file identified as FileNumber reaches the end of the file:

\=WHILE(FPOS(FileNumber)\<=FSIZE(FileNumber))

**Related Functions**

FCLOSE   Closes a text file

FOPEN   Opens a file with the type of permission specified

FREAD   Reads characters from a text file

FREADLN   Reads a line from a text file

FWRITE   Writes characters to a text file

FWRITELN   Writes a line to a text file

Return to [top](#E)

# FREAD

Reads characters from a file, starting at the current position in the
file. (For more information about a file's position, see FPOS.) If FREAD
is successful, it returns the text to the cell containing FREAD and
set's the file's position to the start of the following line. If the end
of the file is reached or if FREAD can't read the file, it returns the
\#N/A error value. Use FREAD instead of FREADLN when you need to read a
specific number of characters from a text file.

**Syntax**

**FREAD**(**file\_num, num\_chars**)

File\_num    is the unique ID number of the file you want to read data
from. File\_num is returned by a previously executed FOPEN function. If
file\_num is not valid, FREAD returns the \#VALUE\! error value.

Num\_chars    specifies how many bytes to read from the file. FREAD can
read up to 255 bytes at a time.

**Example**

The following function reads the next 200 bytes from the open file
identified as FileNumber:

FREAD(FileNumber, 200)

**Related Functions**

FOPEN   Opens a file with the type of permission specified

FPOS   Sets the position in a text file

FREADLN   Reads a line from a text file

FWRITE   Writes characters to a text file

Return to [top](#E)

# FREADLN

Reads characters from a file, starting at the current position in the
file and continuing to the end of the line, placing the characters in
the cell containing FREADLN. (For more information about a file's
position, see FPOS.) If FREADLN is successful, it returns the text it
read, up to but not including the carriage-return and linefeed
characters at the end of the line (in Microsoft Excel for Windows) or
the carriage-return character at the end of the line (in Microsoft Excel
for the Macintosh). If the current file position is the end of the file
or if FREADLN can't read the file, it returns the \#N/A error value.

**Syntax**

**FREADLN**(**file\_num**)

File\_num    is the unique ID number of the file you want to read data
from. File\_num is returned by a previously executed FOPEN function. If
file\_num is not valid, FREADLN returns the \#VALUE\! error value.

**Example**

The following function reads the next line from the open file identified
as FileNumber:

FREADLN(FileNumber)

**Related Functions**

FOPEN   Opens a file with the type of permission specified

FPOS   Sets the position in a text file

FREAD   Reads characters from a text file

FWRITE   Writes characters to a text file

FWRITELN   Writes a line to a text file

Return to [top](#E)

# FREEZE.PANES

Equivalent to clicking the Freeze Panes or Unfreeze Panes command on the
Window menu. Splits the active window into panes, creates frozen panes,
or freezes or unfreezes existing panes. Use FREEZE.PANES to keep row or
column titles on the screen while scrolling to other parts of the sheet.

**Syntax**

**FREEZE.PANES**(logical, col\_split, row\_split)

Logical    is a logical value specifying which command FREEZE.PANES is
equivalent to.

  - > If logical is TRUE, the function is equivalent to the Freeze Panes
    > command. It freezes panes if they exist, or creates them, splits
    > them at the specified position, and freezes them if they do not
    > exist. If the panes are already frozen, FREEZE.PANES takes no
    > action.

  - > If logical is FALSE, the function is equivalent to the Unfreeze
    > Panes command. If no panes exist, FREEZE.PANES takes no action.

  - > If logical is omitted, FREEZE.PANES creates and then freezes panes
    > if no panes exist, freezes existing panes if they're not currently
    > frozen, or unfreezes existing panes if they're currently frozen.

>  

Col\_split    specifies where to split the window vertically and is
measured in columns from the left of the window.

Row\_split    specifies where to split the window horizontally and is
measured in rows from the top of the window.

Col\_split and row\_split are ignored unless logical is TRUE and split
panes do not exist.

**Remarks**

To create panes without freezing or unfreezing them, use the SPLIT
function. You can freeze the panes later using the FREEZE.PANES
function.

**Related Functions**

ACTIVATE   Switches to a window

SPLIT   Splits a window

Return to [top](#E)

# FSIZE

Returns the number of bytes in a file. Use FSIZE to determine the size
of the file, which is the same as the position of the last byte in the
file.

**Syntax**

**FSIZE**(**file\_num**)

File\_num    is the unique ID number of the file whose size you want to
know. File\_num is returned by a previously executed FOPEN function. If
file\_num is not valid, FSIZE returns the \#VALUE\! error value.

**Example**

The following function returns the size in bytes of the open file
identified as FileNumber:

FSIZE(FileNumber)

**Related Functions**

FOPEN   Opens a file with the type of permission specified

FPOS   Sets the position in a text file

Return to [top](#E)

# FTESTV

Performs a two-sample F-test.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**FTESTV**(**inprng1, inprng2**, outrng, labels)

**FTESTV**?(inprng1, inprng2. outrng, labels)

Inprng1    is the input range for the first data set.

Inprng2    is the input range for the second data set.

Outrng    is the first cell (the upper-left cell) in the output table or
the name, as text, of a new sheet to contain the output table. If FALSE,
blank, or omitted, places the output table in a new workbook.

Labels    is a logical value.

  - > If labels is TRUE, then the first row or column of inprng1 and
    > inprng2 contain labels.

  - > If labels is FALSE or omitted, all cells in inprng1 and inprng2
    > are considered data. Microsoft Excel generates appropriate data
    > labels for the output table.

Return to [top](#E)

# FULL

Equivalent to pressing CTRL+F10 (full size) and CTRL+F5 (previous size)
or double-clicking the title bar in Microsoft Excel for Windows version
3.0 or earlier. Equivalent to double-clicking the title bar or clicking
the zoom box in Microsoft Excel for the Macintosh version 3.0 or
earlier. This function is included only for macro compatibility. To
perform the equivalent of a FULL(TRUE) function in Microsoft Excel
version 4.0 or later, use the WINDOW.MAXIMIZE function. To perform the
equivalent of a FULL(FALSE) function in Microsoft Excel version 4.0 or
later, use the WINDOW.RESTORE function.

**Syntax**

**FULL**(logical)

Return to [top](#E)

# FULL.SCREEN

Equivalent to clicking the Full Screen command on the View menu.

**Syntax**

**FULL.SCREEN**(logical)

Logical    switches to full screen if TRUE or omitted; exits full screen
mode if FALSE.

Return to [top](#E)

# FUNCTION.WIZARD

Displays the Paste Function dialog box, which you can use to enter
functions into cells.

**Syntax**

**FUNCTION.WIZARD**?( )

**Remarks**

If you know the function or formula that you want to insert into a cell,
use the FORMULA function.

**Related Function**

FORMULA   Enters values into a cell or range or onto a chart

Return to [top](#E)

# FWRITE

Writes text to a file, starting at the current position in that file.
(For more information about a file's position, see FPOS.) If FWRITE
can't write to the file, it returns the \#N/A error value.

**Syntax**

**FWRITE**(**file\_num, text**)

File\_num    is the unique ID number of the file you want to write data
to. File\_num is returned by a previously executed FOPEN function. If
file\_num is not valid, FWRITE returns the \#VALUE\! error value.

Text    is the text you want to write to the file.

**Example**

The following function writes the current month to the open file
identified as FileNumber:

FWRITE(FileNumber, TEXT(MONTH(NOW()),"mmmm"))

**Related Functions**

FOPEN   Opens a file with the type of permission specified

FPOS   Sets the position in a text file

FREAD   Reads characters from a text file

FWRITELN   Writes a line to a text file

Return to [top](#E)

# FWRITELN

Writes text, followed by a carriage return and linefeed, to a file,
starting at the current position in that file. (For more information
about a file's position, see FPOS.) If FWRITELN can't write to the file,
it returns the \#N/A error value. Use FWRITELN instead of FWRITE when
you want to append a carriage return and linefeed to each group of
characters that you write to a text file.

**Syntax**

**FWRITELN**(**file\_num, text**)

File\_num    is the unique ID number of the file you want to write data
to. File\_num is returned by a previously executed FOPEN function. If
file\_num is not valid, FWRITELN returns the \#VALUE\! error value.

Text    is the text you want to write to the file.

**Remarks**

In Microsoft Excel for Windows, FWRITELN writes text followed by a
carriage return and a line feed. In Microsoft Excel for the Macintosh,
FWRITELN writes text followed by a carriage return only.

**Example**

The following function writes the current month to the open file
identified as FileNumber and starts a new line in the file:

FWRITELN(FileNumber, TEXT(MONTH(NOW()),"mmmm"))

**Related Functions**

FOPEN   Opens a file with the type of permission specified

FPOS   Sets the position in a text file

FREAD   Reads characters from a text file

FWRITE   Writes characters to a text file

Return to [top](#E)

# GALLERY.3D.AREA

Changes the format of the active chart to a 3-D area chart.

**Syntax**

**GALLERY.3D.AREA**(**type\_num**)

**GALLERY.3D.AREA**?(type\_num)

Type\_num    is the number of the 3-D Area format that you want to apply
to the chart.

Return to [top](#E)

# GALLERY.3D.BAR

Changes the active chart to a 3-D bar chart.

**Syntax**

**GALLERY.3D.BAR**(**type\_num**)

**GALLERY.3D.BAR**?(type\_num)

Type\_num    is the number of the 3-D Bar format that you want to apply
to the chart.

Return to [top](#E)

# GALLERY.3D.COLUMN

Changes the format of the active chart to a 3-D column chart.

**Syntax**

**GALLERY.3D.COLUMN**(**type\_num**)

**GALLERY.3D.COLUMN**?(type\_num)

Type\_num    is the number of the 3-D Column format that you want to
apply to the chart.

Return to [top](#E)

# GALLERY.3D.LINE

Changes the format of the active chart to a 3-D line chart.

**Syntax**

**GALLERY.3D.LINE**(**type\_num**)

**GALLERY.3D.LINE**?(type\_num)

Type\_num    is the number of the 3-D Line format that you want to apply
to the chart.

Return to [top](#E)

# GALLERY.3D.PIE

Changes the format of the active chart to a 3-D pie chart.

**Syntax**

**GALLERY.3D.PIE**(**type\_num**)

**GALLERY.3D.PIE**?(type\_num)

Type\_num    is the number of the 3-D Pie format that you want to apply
to the chart.

Return to [top](#E)

# GALLERY.3D.SURFACE

Changes the active chart to a 3-D surface chart.

**Syntax**

**GALLERY.3D.SURFACE**(**type\_num**)

**GALLERY.3D.SURFACE**?(type\_num)

Type\_num    is the number of the 3-D Surface format that you want to
apply to the chart.

Return to [top](#E)

# GALLERY.AREA

Changes the format of the active chart to an area chart.

**Syntax**

**GALLERY.AREA**(**type\_num**, delete\_overlay)

**GALLERY.AREA**?(type\_num, delete\_overlay)

Type\_num    is the number of a format in the AutoFormat dialog box when
a chart is active dialog box that you want to apply to the area chart.

Delete\_overlay    is a logical value specifying whether to delete an
overlay chart.

  - > If delete\_overlay is TRUE, Microsoft Excel deletes all overlays,
    > if present, and applies the new format to the main chart.

  - > If delete\_overlay is FALSE or omitted, Microsoft Excel applies
    > the new format to either the main chart or the overlay, depending
    > on the location of the selected series.

>  

Return to [top](#E)

# GALLERY.BAR

Changes the format of the active chart to a bar chart.

**Syntax**

**GALLERY.BAR**(**type\_num**, delete\_overlay)

**GALLERY.BAR**?(type\_num, delete\_overlay)

Type\_num    is the number of the format that you want to apply to the
bar chart.

Delete\_overlay    is a logical value specifying whether to delete an
overlay chart.

  - > If delete\_overlay is TRUE, Microsoft Excel deletes all overlays,
    > if present, and applies the new format to the main chart.

  - > If delete\_overlay is FALSE or omitted, Microsoft Excel applies
    > the new format to either the main chart or the overlay, depending
    > on the location of the selected series.

>  

Return to [top](#E)

# GALLERY.COLUMN

Changes the format of the active chart to a column chart.

**Syntax**

**GALLERY.COLUMN**(**type\_num**, delete\_overlay)

**GALLERY.COLUMN**?(type\_num, delete\_overlay)

Type\_num    is the number of the format you want to apply to the column
chart.

Delete\_overlay    is a logical value specifying whether to delete an
overlay chart.

  - > If delete\_overlay is TRUE, Microsoft Excel deletes all overlays,
    > if present, and applies the new format to the main chart.

  - > If delete\_overlay is FALSE or omitted, Microsoft Excel applies
    > the new format to either the main chart or the overlay, depending
    > on the location of the selected series.

>  

Return to [top](#E)

# GALLERY.CUSTOM

Changes the format of the active chart to the custom format.

**Syntax**

**GALLERY.CUSTOM**(**name\_text**)

Name\_text    is the name of the custom template you want to apply.

**Related Functions**

ADD.CHART.AUTOFORMAT   Formats a chart using a custom gallery

DELETE.CHART.AUTOFORMAT   Deletes a custom gallery

Return to [top](#E)

# GALLERY.DOUGHNUT

Changes the format of the active chart to a doughnut chart.

**GALLERY.DOUGHNUT**(**type\_num**, delete\_overlay)

**GALLERY.DOUGHNUT**?(type\_num, delete\_overlay)

Type\_num    is the number of the format you want to apply to the
doughnut chart.

Delete\_overlay    is a logical value specifying whether to delete an
overlay chart.

  - > If delete\_overlay is TRUE, Microsoft Excel deletes all overlays,
    > if present, and applies the new format to the main chart.

  - > If delete\_overlay is FALSE or omitted, Microsoft Excel applies
    > the new format to either the main chart or the overlay, depending
    > on the location of the selected series.

>  

Return to [top](#E)

# GALLERY.LINE

Changes the format of the active chart to a line chart.

**Syntax**

**GALLERY.LINE**(**type\_num**, delete\_overlay)

**GALLERY.LINE**?(type\_num, delete\_overlay)

Type\_num    is the number of the format you want to apply to the line
chart.

Delete\_overlay    is a logical value specifying whether to delete an
overlay chart.

  - > If delete\_overlay is TRUE, Microsoft Excel deletes all overlays,
    > if present, and applies the new format to the main chart.

  - > If delete\_overlay is FALSE or omitted, Microsoft Excel applies
    > the new format to either the main chart or the overlay, depending
    > on the location of the selected series.

>  

Return to [top](#E)

# GALLERY.PIE

Changes the format of the active chart to a pie chart.

**Syntax**

**GALLERY.PIE**(**type\_num**, delete\_overlay)

**GALLERY.PIE**?(type\_num, delete\_overlay)

Type\_num    is the number of the format you want to apply to the pie
chart.

Delete\_overlay    is a logical value specifying whether to delete an
overlay chart.

  - > If delete\_overlay is TRUE, Microsoft Excel deletes all overlays,
    > if present, and applies the new format to the main chart.

  - > If delete\_overlay is FALSE or omitted, Microsoft Excel applies
    > the new format to either the main chart or the overlay, depending
    > on the location of the selected series.

>  

Return to [top](#E)

# GALLERY.RADAR

Changes the format of the active chart to a radar chart.

**Syntax**

**GALLERY.RADAR**(**type\_num**, delete\_overlay)

**GALLERY.RADAR**?(type\_num, delete\_overlay)

Type\_num    is the number of the format you want to apply to the radar
chart.

Delete\_overlay    is a logical value specifying whether to delete an
overlay chart.

  - > If delete\_overlay is TRUE, Microsoft Excel deletes all overlays,
    > if present, and applies the new format to the main chart.

  - > If delete\_overlay is FALSE or omitted, Microsoft Excel applies
    > the new format to either the main chart or the overlay, depending
    > on the location of the selected series.

>  

Return to [top](#E)

# GALLERY.SCATTER

Changes the format of the active chart to an xy (scatter) chart.

**Syntax**

**GALLERY.SCATTER**(**type\_num**, delete\_overlay)

**GALLERY.SCATTER**?(type\_num, delete\_overlay)

Type\_num    is the number of the format you want to apply to the xy
(scatter) chart.

Delete\_overlay    is a logical value specifying whether to delete an
overlay chart.

  - > If delete\_overlay is TRUE, Microsoft Excel deletes all overlays,
    > if present, and applies the new format to the main chart.

  - > If delete\_overlay is FALSE or omitted, Microsoft Excel applies
    > the new format to either the main chart or the overlay, depending
    > on the location of the selected series.

>  

Return to [top](#E)

# GET.BAR

Returns the number of the active menu bar. There are two syntax forms of
GET.BAR. Use syntax 1 to return information that you can use with other
functions that manipulate menu bars. Use syntax 2 to return information
that you can use with functions that add, delete, or alter menu
commands.

Syntax 1   Returns the number of the active menu bar

Syntax 2   Returns the name or position number of a specified command on
a menu or of a specified menu on a menu bar

Return to [top](#E)

# GET.BAR Syntax 1

Returns the number of the active menu bar. There are two syntax forms of
GET.BAR. Use syntax 1 to return information that you can use with other
functions that manipulate menu bars. For a list of the ID numbers for
Microsoft Excel's built-in menu bars, see ADD.COMMAND.

**Syntax**

**GET.BAR**( )

**Example**

The following macro formula assigns the name OldBar to the number of the
active menu bar. This is useful if you will need to restore the current
menu bar after displaying another custom menu bar.

SET.NAME("OldBar", GET.BAR())

**Related Functions**

ADD.BAR   Adds a menu bar

SHOW.BAR   Displays a menu bar

GET.BAR Syntax 2   Returns the name or position number of a specified
command on a menu or of a specified menu on a menu bar

Return to [top](#E)

# GET.BAR Syntax 2

Returns the name or position number of a specified command on a menu or
of a specified menu on a menu bar. There are two syntax forms of
GET.BAR. Use syntax 2 to return information that you can use with
functions that add, delete, or alter menu commands.

**Syntax**

**GET.BAR**(**bar\_num, menu, command**, subcommand)

Bar\_num    is the number of a menu bar containing the menu or command
about which you want information. Bar\_num can be the number of a
built-in menu bar or the number returned by a previously run ADD.BAR
function. For a list of the ID numbers for Microsoft Excel's built-in
menu bars, see ADD.COMMAND.

Menu    is the menu on which the command resides or the menu whose name
or position you want. Menu can be the name of the menu as text or the
number of the menu. Menus are numbered starting with 1 from the left of
the menu bar.

Command    is the command or submenu whose name or number you want
returned. Command can be the name of the command from the menu as text,
in which case the number is returned, or the number of the command from
the menu, in which case the name is returned. Commands are numbered
starting with 1 from the top of the menu. If command is 0, the name or
position number of the menu is returned. If an ellipsis (...) follows a
command name, such as the Open... command on the File menu, then you
must include the ellipsis when referring to that command. See the
following examples.

Subcommand    returns the name (if number is used for subcommand) or
position (if name is used for subcommand) of a command on a submenu. If
the command argument refers to an empty submenu, or is a command instead
of a submenu, then using subcommand returns \#N/A.

**Remarks**

  - > If an ampersand is used to indicate the access key in the name of
    > a custom command, the ampersand is included in the name returned
    > by GET.BAR. All built-in commands have an ampersand before the
    > letter used as the access key.

  - > If the command name or position specified does not exist, GET.BAR
    > returns the \#N/A error value.

>  

**Examples**

In the default worksheet and macro sheet menu bar:

GET.BAR(10, "File", "Print...") equals 14

GET.BAR(10, "File", 14) equals "\&Print...^tCTRL+P" (where ^t is a tab
character)

GET.BAR(10, 1, "Open") equals \#N/A

GET.BAR(10, 1, "Open...") equals 2

**Related Functions**

ADD.COMMAND   Adds a command to a menu

DELETE.COMMAND   Deletes a command from a menu

GET.TOOLBAR   Retrieves information about a toolbar

RENAME.COMMAND   Changes the name of a command or menu

GETBAR Syntax 1   Returns the number of the active menu bar

Return to [top](#E)

# GET.CELL

Returns information about the formatting, location, or contents of a
cell. Use GET.CELL in a macro whose behavior is determined by the status
of a particular cell.

**Syntax**

**GET.CELL**(**type\_num**, reference)

Type\_num    is a number that specifies what type of cell information
you want. The following list shows the possible values of type\_num and
the corresponding results.

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

Reference    is a cell or a range of cells from which you want
information.

  - > If reference is a range of cells, the cell in the upper-left
    > corner of the first range in reference is used.

  - > If reference is omitted, the active cell is assumed.

>  

**Tip   **Use GET.CELL(17) to determine the height of a cell and
GET.CELL(44) - GET.CELL(42) to determine the width.

**Examples**

The following macro formula returns TRUE if cell B4 on sheet Sheet1 is
bold:

GET.CELL(20, Sheet1\!$B$4)

You can use the information returned by GET.CELL to initiate an action.
The following macro formula runs a custom function named BoldCell if the
GET.CELL formula returns FALSE:

IF(GET.CELL(20, Sheet1\!$B$4), , BoldCell())

**Related Functions**

ABSREF   Returns the absolute reference of a range of cells to another
range

ACTIVE.CELL   Returns the reference of the active cell

GET.FORMULA   Returns the contents of a cell

GET.NAME   Returns the definition of a name

GET.NOTE   Returns characters from a note

RELREF   Returns a relative reference

Return to [top](#E)

# GET.CHART.ITEM

Returns the vertical or horizontal position of a point on a chart item.
Use these position numbers with FORMAT.MOVE and FORMAT.SIZE to change
the position and size of chart items. Position is measured in points; a
point is 1/72nd of an inch.

**Syntax**

**GET.CHART.ITEM**(**x\_y\_index**, point\_index, item\_text)

X\_y\_index    is a number specifying which of the coordinates you want
returned.

|                 |                         |
| --------------- | ----------------------- |
| **X\_y\_index** | **Coordinate returned** |
| 1               | Horizontal coordinate   |
| 2               | Vertical coordinate     |

Point\_index    is a number specifying the point on the chart item.
These indexes are described below. If point\_index is omitted, it is
assumed to be 1.

  - > If the specified item is a point, point\_index must be 1.

  - > If the specified item is any line other than a data line, use the
    > following values for point\_index.

|                  |                         |
| ---------------- | ----------------------- |
| **Point\_index** | **Chart item position** |
| 1                | Lower or left           |
| 2                | Upper or right          |

 

  - > If the selected item is a legend, plot area, chart area, or an
    > area in an area chart, use the following values for point\_index.

|                  |                         |
| ---------------- | ----------------------- |
| **Point\_index** | **Chart item position** |
| 1                | Upper left              |
| 2                | Upper middle            |
| 3                | Upper right             |
| 4                | Right middle            |
| 5                | Lower right             |
| 6                | Lower middle            |
| 7                | Lower left              |
| 8                | Left middle             |

 

  - > If the selected item is an arrow in Microsoft Excel 4.0, use the
    > following values for point\_index. In Microsoft Excel version 5.0
    > or later, arrows are named lines, and the arrowhead position
    > returned is equivalent to the end of a line where the arrowhead
    > begins.

|                  |                         |
| ---------------- | ----------------------- |
| **Point\_index** | **Chart item position** |
| 1                | Arrow shaft             |
| 2                | Arrowhead               |

 

  - > If the selected item is a pie slice, use the following values for
    > point\_index.

|                  |                                              |
| ---------------- | -------------------------------------------- |
| **Point\_index** | **Chart item position**                      |
| 1                | Outermost counterclockwise point             |
| 2                | Outer center point                           |
| 3                | Outermost clockwise point                    |
| 4                | Midpoint of the most clockwise radius        |
| 5                | Center point                                 |
| 6                | Midpoint of the most counterclockwise radius |

Item\_text    is a selection code that specifies which item of a chart
to select. See the chart form of SELECT for the item\_text codes to use
for each item of a chart.

  - > If item\_text is omitted, it is assumed to be the currently
    > selected item.

  - > If item\_text is omitted and no item is selected, GET.CHART.ITEM
    > returns the \#VALUE\! error value.

>  

**Remarks**

If the specified item does not exist, or if a chart is not active when
the function is carried out, the \#VALUE\! error value is returned.

**Examples**

The following macro formulas return the horizontal and vertical
locations, respectively, of the top of the main-chart value axis:

GET.CHART.ITEM(1, 2, "Axis 1")

GET.CHART.ITEM(2, 2, "Axis 1")

You could then use FORMAT.MOVE to move a floating text item to the
position returned by these two formulas.

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

GET.FORMULA   Returns the contents of a cell

Return to [top](#E)

# GET.DEF

Returns the name, as text, that is defined for a particular area, value,
or formula in a workbook. Use GET.DEF to get the name corresponding to a
definition. To get the definition of a name, use GET.NAME.

**Syntax**

**GET.DEF**(**def\_text**, document\_text, type\_num)

Def\_text    can be anything you can define a name to refer to,
including a reference, a value, an object, or a formula.

  - > References must be given in R1C1 style, such as "R3C5".

  - > If def\_text is a value or formula, it is not necessary to include
    > the equal sign that is displayed in the Refers To box in the
    > Define Name dialog box, which appears when you choose the Name
    > command from the Define submenu on the Insert Menu.

  - > If there is more than one name for def\_text, GET.DEF returns the
    > first name. If no name matches def\_text, GET.DEF returns the
    > \#NAME? error value.

Document\_text    specifies the sheet or macro sheet that def\_text is
on. If document\_text is omitted, it is assumed to be the active macro
sheet.

Type\_num    is a number from 1 to 3 specifying which types of names are
returned.

|               |                   |
| ------------- | ----------------- |
| **Type\_num** | **Returns**       |
| 1 or omitted  | Normal names only |
| 2             | Hidden names only |
| 3             | All names         |

**Examples**

If the specified range in Sheet4 is named Sales, the following macro
formula returns "Sales":

GET.DEF("R2C2:R9C6", "Sheet4")

If the value 100 in Sheet4 is defined as Constant, the following macro
formula returns "Constant":

GET.DEF("100", "Sheet4")

If the specified formula in Sheet4 is named SumTotal, the following
macro formula returns "SumTotal":

GET.DEF("SUM(R1C1:R10C1)", "Sheet4")

If 3 is defined as the hidden name Counter on the active macro sheet,
the following macro formula returns "Counter":

GET.DEF("3", , 2)

**Related Functions**

GET.CELL   Returns information about the specified cell

GET.NAME   Returns the definition of a name

GET.NOTE   Returns characters from a note

NAMES   Returns the names defined on a workbook

Return to [top](#E)

# GET.DOCUMENT

Returns information about a sheet in a workbook.

**Syntax**

**GET.DOCUMENT**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of information you
want. The following lists show the possible values of type\_num and the
corresponding results.

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

**Note   **GET.DOCUMENT(62) and GET.DOCUMENT(63) are mutually exclusive.
If one returns a value, then the other returns the \#N/A error value.

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

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

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

GET.CELL   Returns information about the specified cell

GET.WINDOW   Returns information about a window

GET.WORKSPACE   Returns information about the workspace

Return to [top](#E)

# GET.FORMULA

Returns the contents of a cell as they would appear in the formula bar.
The contents are given as text, for example, "=2\*PI()/360". If the
formula contains references, they are returned as R1C1-style references,
such as "=RC\[1\]\*(1+R1C1)". Use GET.FORMULA to get a formula from a
cell in order to edit its arguments. Use GET.CELL(6) to get a formula in
either A1 or R1C1 format, depending on the workspace setting.

**Syntax**

**GET.FORMULA**(**reference**)

Reference    is a cell or range of cells on a sheet or macro sheet.

  - > If a range of cells is selected, GET.FORMULA returns the contents
    > of the upper-left cell in reference.

  - > Reference can be an external reference.

  - > Reference can be the object identifier of a picture created by the
    > camera tool.

  - > Reference can also be a reference to a chart series in the form
    > "Sn" where n is the number of the series. When a chart series is
    > specified, GET.FORMULA returns the series formula using R1C1-style
    > references.

**Tip**   If you want to get the formula in the active cell, use the
ACTIVE.CELL function as the reference argument.

**Examples**

If cell A3 on the active sheet contains the number 523, then:

GET.FORMULA(\!$A$3) equals "523"

If cell C2 on the active sheet contains the formula =B2\*(1+$A$1), then:

GET.FORMULA(\!$C$2) equals "=RC\[-1\]\*(1+R1C1)"

The following macro formula returns the contents of the active cell on
the active sheet:

GET.FORMULA(ACTIVE.CELL())

**Related Functions**

GET.CELL   Returns information about the specified cell

GET.DEF   Returns a name matching a definition

GET.NAME   Returns the definition of a name

GET.NOTE   Returns characters from a comment

Return to [top](#E)

# GET.LINK.INFO

Returns information about the specified link. Use GET.LINK.INFO to get
information about the update settings of a link.

**Syntax**

**GET.LINK.INFO**(**link\_text, type\_num**, type\_of\_link, reference)

Link\_text    is the path of the link as displayed in the Links dialog
box, which appears when you choose the Links command from the Edit menu.
The path to the file you wish to return DDE information on must be
surrounded by single quotes.

Type\_num    is a number that specifies what type of information about
the currently selected link to return. Type\_num 2 applies only to
publishers and subscribers in Microsoft Excel for the Macintosh.

|               |                                                                                                                |
| ------------- | -------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Returns**                                                                                                    |
| 1             | If the link is set to automatic update, returns 1; otherwise 2.                                                |
| 2             | Date of the latest edition as a serial number. Returns \#N/A if link\_text is not a publisher or a subscriber. |

Type\_of\_link    is a number from 1 to 6 that specifies what type of
link you want to get information about.

|                    |                              |
| ------------------ | ---------------------------- |
| **Type\_of\_link** | **Link document type**       |
| 1                  | Not applicable               |
| 2                  | DDE link (Microsoft Windows) |
| 3                  | Not applicable               |
| 4                  | Not applicable               |
| 5                  | Publisher (Macintosh)        |
| 6                  | Subscriber (Macintosh)       |

Reference    specifies the cell range in R1C1 format of the publisher or
subscriber that you want information about. Reference is required if you
have more than one publisher or subscriber of a single edition name on
the active workbook. Use reference to specify the location of the
subscriber you want to return information about. If the subscriber is a
picture, or if the publisher is an embedded chart, reference is the
number of the object as displayed in the Name box.

**Remarks**

  - > If Microsoft Excel cannot find link\_text, or if type\_of\_link
    > does not match the link specified by link\_text, GET.LINK.INFO
    > returns the \#VALUE\! error value.

  - > If you have more than one subscriber to the edition link\_text or
    > if the same area is published more than once, you must specify
    > reference.

>  

**Example**

In Microsoft Excel for Windows, the following macro formula returns
information about a DDE link to a Microsoft Word for Windows document.
The document is named NEWPROD.DOC.

GET.LINK.INFO("WinWord|'C:\\WINWORD\\NEWPROD.DOC'\!DDE\_LINK1", 1, 2)

In Microsoft Excel for the Macintosh, the following macro formula
returns information about a link to a publisher defined in cells A1:C3
on a workbook named New Products.

GET.LINK.INFO("A1:C3 New Products Edition \#1", 2, 5, "'New
Products'\!R1C1:R3C3")

**Related Functions**

CREATE.PUBLISHER   Creates a publisher from the selection

SUBSCRIBE.TO   Inserts contents of an edition into the active workbook

UPDATE.LINK   Updates a link to another workbook

Return to [top](#E)

# GET.NAME

Returns the definition of a name as it appears in the Refers To box of
the Define Name dialog box, which appears when you choose the Define
command from the Name submenu on the Insert menu. If the definition
contains references, they are given as R1C1-style references. Use
GET.NAME to check the value defined by a name. To get the name
corresponding to a definition, use GET.DEF.

**Syntax**

**GET.NAME**(**name\_text**, info\_type)

Name\_text    can be a name defined on the macro sheet; an external
reference to a name defined on the active workbook, for example,
"\!Sales"; or an external reference to a name defined on a particular
open workbook, for example, "\[Book1\]SHEET1\!Sales". Name\_text can
also be a hidden name.

Info\_type     specifies the type of information to return about the
name. If 1 or omitted, the definition is returned. If 2, returns TRUE if
the name is defined for just the sheet, FALSE if the name is defined for
the entire workbook.

**Remarks**

If the Contents check box has been selected in the Protect Sheet dialog
box to protect the workbook containing the name, GET.NAME returns the
\#N/A error value. To see the Protect Sheet dialog box, choose the
Protect Sheet command on the Protection submenu from the Tools menu.

**Examples**

If the name Sales on a macro sheet is defined as the number 523, then:

GET.NAME("Sales") equals "=523"

If the name Profit on the active sheet is defined as the formula
=Sales-Costs, then:

GET.NAME("\!Profit") equals "=Sales-Costs"

If the name Database on the active sheet is defined as the range
A1:F500, then:

GET.NAME("\!Database") equals "=R1C1:R500C6"

**Related Functions**

DEFINE.NAME   Defines a name on the active or macro sheet

GET.CELL   Returns information about the specified cell

GET.DEF   Returns a name matching a definition

NAMES   Returns the names defined in a workbook

SET.NAME   Defines a name as a value

Return to [top](#E)

# GET.NOTE

Returns characters from a comment.

**Syntax**

**GET.NOTE**(cell\_ref, start\_char, num\_chars)

Cell\_ref    is the cell to which the note is attached. If cell\_ref is
omitted, the comment attached to the active cell is returned.

Start\_char    is the number of the first character in the comment to
return. If start\_char is omitted, it is assumed to be 1, the first
character in the comment.

Num\_chars    is the number of characters to return. Num\_chars must be
less than or equal to 255. If num\_chars is omitted, it is assumed to be
the length of the comment attached to cell\_ref.

**Examples**

The following macro formula returns the first 200 characters in the
comment attached to cell A3 on the active sheet:

GET.NOTE(\!$A$3, 1, 200)

In Microsoft Excel for Windows, the following macro formula returns the
10th through the 39th characters of the comment attached to cell C2 on
SALES.XLS:

GET.NOTE("\[SALES.XLS\]Sheet1\!R2C3", 10, 30)

In Microsoft Excel for the Macintosh, the following macro formula
returns the 10th through the 39th characters of the comment attached to
cell C2 on SALES:

GET.NOTE("\[SALES\]Sheet1\!R2C3", 10, 30)

Use GET.NOTE with the NOTE function to move the contents of a comment to
a cell or text box or to another comment attached to a cell:

NOTE(GET.NOTE(\!$B$10),ACTIVE.CELL())

**Related Functions**

GET.CELL   Returns information about the specified cell

NOTE   Creates or changes a comment.

Return to [top](#E)

# GET.OBJECT

Returns information about the specified object. Use GET.OBJECT to return
information you can use in other macro formulas that manipulate objects.

**Syntax**

**GET.OBJECT**(**type\_num**, object\_id\_text, start\_num, count\_num,
item\_index)

Type\_num    is a number specifying the type of information you want
returned about an object. GET.OBJECT returns the \#VALUE\! error value
(and the macro is halted) if an object isn't specified or if more than
one object is selected.

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

Object\_id\_text    is the name and number, or number alone, of the
object you want information about. Object\_id\_text is the text
displayed in the reference area when the object is selected. If
object\_id\_text is omitted, it is assumed to be the selected object. If
object\_id\_text is omitted and no object is selected, GET.OBJECT
returns the \#REF\! error value and interrupts the macro.

Start\_num    is the number of the first character in the text box or
button or the first vertex in a polygon you want information about.
Start\_num is ignored unless a text box, button, or polygon is specified
by type\_num and object\_id\_text. If start\_num is omitted, it is
assumed to be 1.

Count\_num    is the number of characters in a text box or button, or
the number of vertices in a polygon, starting at start\_num, that you
want information about. Count\_num is ignored unless a text box, button,
or polygon is specified by type\_num and object\_id\_text. If count\_num
is omitted, it is assumed to be 255.

Item\_index    is the index number or position of the item in the list
box or drop-down box that you want information about, ranging from 1 to
the number of items in the list box or drop-down box.

**Tip   **Use GET.OBJECT(45) - GET.OBJECT(43) to determine the height of
an object and GET.OBJECT(44) - GET.OBJECT(42) to determine the width.

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

CREATE.OBJECT   Creates an object

FONT.PROPERTIES   Applies a font to the selection

OBJECT.PROTECTION   Controls how an object is protected

PLACEMENT   Determines an object's relationship to underlying cells

Return to [top](#E)

# GET.PIVOT.FIELD

Returns information about a field in a PivotTable report.

**Syntax**

**GET.PIVOT.FIELD**(type\_num, pivot\_field\_name, pivot\_table\_name)

Type\_num    is a value from 1 to 17 that returns the following types of
information:

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Value</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns an array of all the items which make up pivot_field_name. The array is made up of text constants, dates or numbers depending on the field.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns an array of all items which are set to show with the pivot_field_name. The array is made up of text constants, dates or numbers depending on the field. The array is returned in the order that the items are displayed in the PivotTable report. If pivot_field_name is a page field, then the array contains only one element, the value corresponding to the active page (this could be all if the All item is showing).</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns an array of all items which are hidden in the pivot_field_name. The array is made up of text constants, dates or numbers depending on the field. If pivot_field_name is a data field or the data header name, this function returns the #N/A! error value.</td>
</tr>
<tr class="odd">
<td>4</td>
<td><p>Returns an integer describing where the field is displayed in the active PivotTable report (either row or column):</p>
<p>0 = Hidden</p>
<p>1 = Row</p>
<p>2 = Col</p>
<p>3 = Page</p>
<p>4 = Data</p></td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns an array of all items in pivot_field_name that are group parents. The array is made up of text constants, dates or numbers depending on the field. The array is returned in the order which these items appear in the PivotTable report. Returns #N/A if there are no group parents and if the pivot_field_name is a data field or the data field header.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a number between 0 and 4095 which describes the subtotals attached to the field. The number is the sum of the values associated with each subtotal function. See PIVOT.FIELD.PROPERTIES for a list of all the values associated with subtotal calculations. If the field is showing as a data field or data field header, #N/A! is returned.</td>
</tr>
<tr class="even">
<td>7</td>
<td><p>Returns an integer describing the type of data contained in the field:</p>
<p>0 = Text</p>
<p>1 = Number</p>
<p>2 = Date</p></td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns an array five columns wide and one row high describing the summary function's custom calculation shown with the specified field (Data field) in the PivotTable report. The array will look as follows: {function, calculation, base field, base item, number format}. If pivot_field_name is not showing in the active PivotTable report as a data field, #N/A! is returned.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a reference to all of pivot_field_name's items currently showing in the active PivotTable report. If pivot_field_name is hidden, #N/A! is returned. If pivot_field_name is a page field, the reference to the currently showing page item is returned. If pivot_field_name is a data field, a reference to all the data for this field in the PivotTable report is returned. The references are returned as text.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a reference to the header cell for pivot_field_name. If pivot_field_name is a data field, a reference to all the headers in the data row or column is returned. If pivot_field_name is hidden, #N/A! is returned. The reference is returned as text.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the number of grouped fields in the grouped field set which includes pivot_field_name. If pivot_field_name is neither a parent field nor a child field, 1 is returned. If pivot_field_name is a data field or data header name, the function returns the #N/A! error value.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the level of pivot_field_name in the grouped field set which includes pivot_field_name. Returns 1 for the highest level parent field, 2 for its child field, and so on. If pivot_field_name is neither a parent field nor a child field, 1 is returned. If pivot_field_name is a data field or data header name, the function returns the #N/A! error value.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the name of the parent field for pivot_field_name as a text constant. If pivot_field_name is not a child field, #N/A! is returned.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the name of the child field for pivot_field_name as a text constant. If pivot_field_name is not a parent field, #N/A! is returned.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns a text constant representing the original name of the field in the data source.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns the position of the field among all the other fields in its orientation. For instance, a 1 would be returned if the field was the first row field.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns an array of all items in pivot_field_name that are group children. The array is made up of text constants, dates or numbers depending on the field. The array is returned in the order which these items appear in the PivotTable report. Returns #N/A if there are no group children, and if the pivot_field_name is a data field or the data field header.</td>
</tr>
</tbody>
</table>

Pivot\_field\_name    is the name of the field that you want information
about. If there is no field named pivot\_field\_name in the PivotTable
report, returns \#VALUE\!.

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, the PivotTable report
containing the active cell is used. If the active cell is not in a
PivotTable report, the \#VALUE\! error value is returned.

**Related Functions**

GET.PIVOT.ITEM   Returns information about an item in a PivotTable
report.

GET.PIVOT.TABLE   Returns information about a PivotTable report.

Return to [top](#E)

# GET.PIVOT.ITEM

Returns information about an item in a PivotTable report.

**Syntax**

**GET.PIVOT.ITEM**(**type\_num**, pivot\_item\_name, pivot\_field\_name,
pivot\_table\_name)

Type\_num    is a value from 1 to 9 the represents the type of
information you want about an item in a PivotTable report.

|               |                                                                                                                                                                                                                                                                                   |
| ------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Information**                                                                                                                                                                                                                                                                   |
| 1             | Returns the position of the item in its field. Returns \#N/A if pivot\_field\_name is a data field. Returns \#N/A\! if the item is hidden.                                                                                                                                        |
| 2             | Returns the reference to all the cells in the PivotTable header currently containing pivot\_item\_name. This reference is returned as text. If pivot\_item\_name is currently not showing in the PivotTable report, \#N/A\! is returned.                                          |
| 3             | Returns the reference to all the data in the PivotTable report which is qualified by pivot\_item\_name. This reference is returned as text. If pivot\_item\_name is currently not showing in the PivotTable report, \#N/A\! is returned.                                          |
| 4             | Returns an array of text constants representing the children of pivot\_item\_name if pivot\_item\_name is a parent. Otherwise the function returns \#N/A\!.                                                                                                                       |
| 5             | Returns a text constant representing the parent of pivot\_item\_name, if pivot\_item\_name exists as part of a group. Otherwise the function returns \#N/A\!.                                                                                                                     |
| 6             | Returns TRUE if pivot\_item\_name is a member of a group which is currently expanded to show detail. Returns FALSE if pivot\_item\_name is a member of a group currently collapsed to hide detail. If pivot\_item\_name is not a member of a group, the function returns \#N/A\!. |
| 7             | Returns TRUE if pivot\_item\_name is expanded to show detail. Returns FALSE if pivot\_item\_name is collapsed to hide detail.                                                                                                                                                     |
| 8             | Returns TRUE if the item pivot\_item\_name is currently visible, FALSE if it is hidden.                                                                                                                                                                                           |
| 9             | Returns the name of the item as it appeared in the original at a source. This will differ from the current item name only if the user changes the name of the item after creating the PivotTable report.                                                                          |

Pivot\_item\_name    is the name of the item that you want information
about. If there is no item named pivot\_item\_name in the PivotTable
report, returns \#VALUE\!.

Pivot\_field\_name    is the name of the field that you want information
about. If there is no field named pivot\_field\_name in the PivotTable
report, returns \#VALUE\!.

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, uses the PivotTable
report containing the active cell. If the active cell is not in a
PivotTable report, the \#VALUE\! error value is returned.

**Related Functions**

GET.PIVOT.FIELD   Returns information about an item in a PivotTable
report.

GET.PIVOT.TABLE   Returns information about a PivotTable report.

Return to [top](#E)

# GET.PIVOT.TABLE

Returns information about a PivotTable report.

**Syntax**

**GET.PIVOT.TABLE**(**type\_num**,pivot\_table\_name)

Type\_num is a value from 1 to 22 that represents a type of information
you want about a PivotTable report.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Information</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>Returns the name of the person who last updated the PivotTable report, as a text constant.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>Returns the date the PivotTable report was last updated, as a serial number.</td>
</tr>
<tr class="even">
<td>3</td>
<td>Returns a horizontal array of text constants representing all the fields in the PivotTable report.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>Returns an integer representing the number of fields in the PivotTable report.</td>
</tr>
<tr class="even">
<td>5</td>
<td>Returns a horizontal array of text constants representing all the visible fields in the PivotTable report (rows, columns, pages or data)</td>
</tr>
<tr class="odd">
<td>6</td>
<td>Returns a horizontal array of text constants representing all the hidden fields in the PivotTable report. Return #N/A if no hidden fields.</td>
</tr>
<tr class="even">
<td>7</td>
<td>Returns a horizontal array of text constants representing the names of all the fields currently showing in the PivotTable report as row fields. Returns #N/A if there are no row fields.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as column fields. Returns #N/A if no column fields exist.</td>
</tr>
<tr class="even">
<td>9</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as page fields. Return #N/A if no page fields exist.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>Returns a horizontal array of text constants representing all the fields currently showing in the PivotTable report as data fields. Returns #N/A if there are no data fields.</td>
</tr>
<tr class="even">
<td>11</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (not including the page header). This reference is returned as text.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>Returns the smallest rectangular reference which bounds the PivotTable report and all headers (including the page headers). This reference is returned as text.</td>
</tr>
<tr class="even">
<td>13</td>
<td>Returns the reference to the row header area as text. The row header area includes each row field header along with all the items in each row field. Returns #N/A if there are no row headers.</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Returns the reference to the column header area as text. The column header area includes each column field header along with all the items in each column field. Returns #N/A if there are no column headers.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Returns the reference to the data header area as text. The data header area includes the data field header along with all the headers in the data row/col. Returns #N/A if there is no data field.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Returns a reference to all the page headers as text.</td>
</tr>
<tr class="even">
<td>17</td>
<td>Returns the reference to the PivotTable report data area as text.</td>
</tr>
<tr class="odd">
<td>18</td>
<td>Returns TRUE if the PivotTable report is set to show row grand totals.</td>
</tr>
<tr class="even">
<td>19</td>
<td>Returns TRUE if the PivotTable report is set to show column grand totals.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Returns TRUE if the user is saving data with the PivotTable report.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Returns TRUE if the PivotTable report is set up to Autoformat on pivoting.</td>
</tr>
<tr class="odd">
<td>22</td>
<td><p>Returns the data source of the PivotTable report. The kind of information returned depends on the data source:</p>
<p>If the data source is a Microsoft Excel list or database, the cell reference is returned as text.</p>
<p>If the data source is an external data source, then an array is returned. Each row consists of a SQL connection string with the remaining elements as the query string broken down into 200 character segments.</p>
<p>If the data source is Multiple Consolidation ranges, then a two dimensional array is returned, each row of which consists of a reference and associated page field items.</p>
<p>If the data source is another PivotTable report, then one of the above three kinds of information is returned.</p></td>
</tr>
</tbody>
</table>

Pivot\_table\_name    is the name of a PivotTable report containing the
field that you want information about. If omitted, uses the PivotTable
report containing the active cell.

**Remarks**

Returns \#VALUE\! error value when pivot\_table\_name is not a valid
PivotTable name on the active sheet and the active cell is not within a
PivotTable report.

**Related Functions**

GET.PIVOT.FIELD   Returns information about an item in a PivotTable
report.

GET.PIVOT.ITEM   Returns information about a PivotTable report.

Return to [top](#E)

# GET.TOOL

Returns information about a button or buttons on a toolbar. Use GET.TOOL
to get information about a button to use with functions that add,
delete, or alter buttons.

**Syntax**

**GET.TOOL**(**type\_num, bar\_id, position**)

Type\_num    specifies what type of information you want GET.TOOL to
return.

|               |                                                                                                                          |
| ------------- | ------------------------------------------------------------------------------------------------------------------------ |
| **Type\_num** | **Returns**                                                                                                              |
| 1             | The button's ID number. Gaps are represented by zeros.                                                                   |
| 2             | The reference of the macro assigned to the button. If no macro is assigned, GET.TOOL returns the \#N/A error value.      |
| 3             | If the button is down, returns TRUE. If the button is up, returns FALSE.                                                 |
| 4             | If the button is enabled, returns TRUE. If the button is disabled, returns FALSE.                                        |
| 5             | A logical value indicating the type of the face on the button:                                                           |
|               | TRUE = bitmap                                                                                                            |
|               | FALSE = a default button face                                                                                            |
| 6             | The help\_text reference associated with the custom button. If the button is built-in, returns \#N/A.                    |
| 7             | The balloon\_text reference associated with the custom button. If the button is built-in, returns the \#N/A error value. |
| 8             | The Help context string associated with the custom button.                                                               |
| 9             | The Tip\_text associated with the custom button.                                                                         |

Bar\_id    specifies the number or name of the toolbar for which you
want information. For detailed information about bar\_id, see ADD.TOOL.

Position    specifies the position of the button on the toolbar.
Position starts with 1 at the left side (if horizontal) or at the top
(if vertical). A position can be occupied by a button or a gap.

**Example**

The following macro formula requests the help text associated with the
third button in Toolbar2:

GET.TOOL(6, "Toolbar2", 3)

**Related Functions**

ADD.TOOL   Adds one or more buttons to a toolbar

DELETE.TOOL   Deletes a button from a toolbar

ENABLE.TOOL   Enables or disables a button on a toolbar

GET.TOOLBAR   Retrieves information about a toolbar

Return to [top](#E)

# GET.TOOLBAR

Returns information about one toolbar or all toolbars. Use GET.TOOLBAR
to get information about a toolbar to use with functions that add,
delete, or alter toolbars.

**Syntax**

**GET.TOOLBAR**(**type\_num**, bar\_id)

Type\_num    specifies what type of information to return. If type\_num
is 8 or 9, GET.TOOLBAR returns an array of names or numbers of all
visible or hidden toolbars. Otherwise, bar\_id is required, and
GET.TOOLBAR returns the requested information about the specified
toolbar.

|               |                                                                                                                                                   |
| ------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Returns**                                                                                                                                       |
| 1             | A horizontal array of all tool IDs on the toolbar, ordered by position. Gaps are represented by zeros.                                            |
| 2             | Number indicating the horizontal position (x-coordinate) of the toolbar in the docked or floating region. For more information, see SHOW.TOOLBAR. |
| 3             | Number indicating the vertical position (y-coordinate) of the toolbar in the docked or floating region.                                           |
| 4             | Number indicating the width of the toolbar in points.                                                                                             |
| 5             | Number indicating the height of the toolbar in points.                                                                                            |
| 6             | Number indicating the toolbar location:                                                                                                           |
|               | 1 = Top dock in the workspace                                                                                                                     |
|               | 2 = Left dock in the workspace                                                                                                                    |
|               | 3 = Right dock in the workspace                                                                                                                   |
|               | 4 = Bottom dock in the workspace                                                                                                                  |
|               | 5 = Floating                                                                                                                                      |
| 7             | If the toolbar is visible, returns TRUE. If the toolbar is hidden, returns FALSE.                                                                 |
| 8             | An array of toolbar IDs (names or numbers in the bar\_id array) for all toolbars, visible and hidden.                                             |
| 9             | An array of toolbar IDs (names or numbers in the bar\_id array) for all visible toolbars.                                                         |
| 10            | If the toolbar is visible in full-screen mode, returns TRUE; otherwise, returns FALSE.                                                            |

Bar\_id    specifies the number or name of a toolbar for which you want
information. If type\_num is 8 or 9, Microsoft Excel ignores bar\_id.
For detailed information about bar\_id, see ADD.TOOL.

**Remarks**

If you request position information for a hidden toolbar, Microsoft
Excel returns the position where the toolbar would appear if shown.

**Examples**

The following macro formula returns information about the width of
Toolbar1:

GET.TOOLBAR(4, "Toolbar1")

When the following macro formula is entered as an array with
CTRL+SHIFT+ENTER, the IDs of all visible toolbars are returned, and the
array is named All\_Bar\_Ids:

SET.NAME("All\_Bar\_Ids", GET.TOOLBAR(9))

**Related Functions**

ADD.TOOL   Adds one or more buttons to a toolbar

ADD.TOOLBAR   Creates a new toolbar with the specified tools

DELETE.TOOLBAR   Deletes custom toolbars

GET.TOOL   Returns information about a tool or tools on a toolbar

SHOW.TOOLBAR   Hides or displays a toolbar

Return to [top](#E)

# GET.WINDOW

Returns information about a window. Use GET.WINDOW in a macro that
requires the status of a window, such as its name, size, position, and
display options.

**Syntax**

**GET.WINDOW**(**type\_num**, window\_text)

Type\_num    is a number that specifies what type of window information
you want. The following list shows the possible values of type\_num and
the corresponding results:

|               |                                                                                                                                                                                                                                                                                                                               |
| ------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Returns**                                                                                                                                                                                                                                                                                                                   |
| 1             | Name of the workbook and sheet in the window as text. For compatibility with Microsoft Excel version 4.0, if the window contains only a single sheet that has the same name as the workbook without its extension, returns only the name of the book. Otherwise, returns the name of the sheet in the form "\[Book1\]Sheet1". |
| 2             | Number of the window.                                                                                                                                                                                                                                                                                                         |
| 3             | X position, measured in points from the left edge of the workspace (in Microsoft Excel for Windows) or screen (in Microsoft Excel for the Macintosh) to the left edge of the window.                                                                                                                                          |
| 4             | Y position, measured in points from the bottom edge of the formula bar to the top edge of the window.                                                                                                                                                                                                                         |
| 5             | Width, measured in points.                                                                                                                                                                                                                                                                                                    |
| 6             | Height, measured in points.                                                                                                                                                                                                                                                                                                   |
| 7             | If window is hidden, returns TRUE; otherwise, returns FALSE.                                                                                                                                                                                                                                                                  |

The rest of the values for type\_num apply only to worksheets and macro
sheets, except where indicated:

|               |                                                                                                                                                                       |
| ------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Returns**                                                                                                                                                           |
| 8             | If formulas are displayed, returns TRUE; otherwise, returns FALSE.                                                                                                    |
| 9             | If gridlines are displayed, returns TRUE; otherwise, returns FALSE.                                                                                                   |
| 10            | If row and column headings are displayed, returns TRUE; otherwise, returns FALSE.                                                                                     |
| 11            | If zeros are displayed, returns TRUE; otherwise, returns FALSE.                                                                                                       |
| 12            | Gridline and heading color as a number in the range 1 to 56, corresponding to the colors in the View tab of the Options dialog box; if color is automatic, returns 0. |

Values 13 to 16 for type\_num return arrays that specify which rows or
columns are at the top and left edges of the panes in the window and the
widths and heights of those panes.

  - > The first number in the array corresponds to the first pane, the
    > second number to the second pane, and so on.

  - > If the edge of the pane occurs at the boundary between rows or
    > columns, the number returned is an integer.

  - > If the edge of the pane occurs within a row or column, the number
    > returned has a fractional part that represents the fraction of the
    > row or column visible within the pane.

  - > The numbers can be used as arguments to the SPLIT function to
    > split a window at specific locations.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>13</td>
<td>Leftmost column number of each pane, in a horizontal numeric array</td>
</tr>
<tr class="odd">
<td>14</td>
<td>Top row number of each pane, in a horizontal numeric array.</td>
</tr>
<tr class="even">
<td>15</td>
<td>Number of columns in each pane, in a horizontal numeric array.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Number of rows in each pane, in a horizontal numeric array.</td>
</tr>
<tr class="even">
<td>17</td>
<td><p>Number indicating the active pane:</p>
<p>1 = Upper, left, or upper-left</p>
<p>2 = Right or upper-right</p>
<p>3 = Lower or lower-left</p>
<p>4 = Lower-right</p></td>
</tr>
<tr class="odd">
<td>18</td>
<td>If window has a vertical split, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>19</td>
<td>If window has a horizontal split, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>20</td>
<td>If window is maximized, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>21</td>
<td>Reserved</td>
</tr>
<tr class="odd">
<td>22</td>
<td>If the Outline Symbols check box is selected in the View tab of the Options dialog box, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>23</td>
<td><p>Number indicating the size of the window (including charts):</p>
<p>1 = Restored</p>
<p>2 = Minimized (displayed as an icon)</p>
<p>3 = Maximized</p></td>
</tr>
<tr class="odd">
<td>24</td>
<td>If panes are frozen on the active window, returns TRUE; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>25</td>
<td>The numeric magnification of the active window (as a percentage of normal size) as set in the Zoom dialog box, or 100 if none is specified.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>Returns TRUE if horizontal scrollbars are displayed in the active window; otherwise, returns FALSE.</td>
</tr>
<tr class="even">
<td>27</td>
<td>Returns TRUE if vertical scrollbars are displayed in the active window; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>Returns the tab ratio of workbook tabs to horizontal scrollbar, from 0 to 1. The default is .6.</td>
</tr>
<tr class="even">
<td>29</td>
<td>Returns TRUE if workbook tabs are displayed in the active window; otherwise, returns FALSE.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>Returns the title of the active sheet in the window in the form "[Book1]Sheet1".</td>
</tr>
<tr class="even">
<td>31</td>
<td>Returns the name of a workbook only, without read/write indicated. For example, if Book1.xls is read only, then "Book.xls" will be returned without "[Read Only]" appended.</td>
</tr>
</tbody>
</table>

Window\_text    is the name that appears in the title bar of the window
that you want information about. If window\_text is omitted, it is
assumed to be the active window.

**Examples**

If the active window contains the workbook Book1, then:

GET.WINDOW(1) equals "Book1"

If the title of the active window is Macro1:3, then:

GET.WINDOW(2) equals 3

In Microsoft Excel for Windows, the following macro formula returns the
gridline and heading color of REPORT.XLS:

GET.WINDOW(12, "REPORT.XLS")

In Microsoft Excel for the Macintosh, the following macro formula
returns the gridline and heading color of REPORT MASTER:

GET.WINDOW(12, "REPORT MASTER")

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

GET.WORKSPACE   Returns information about the workspace

Return to [top](#E)

# GET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the Bcc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>The subject of the PowerTalk mailer, as text.</td>
</tr>
<tr class="even">
<td>29</td>
<td>The enclosures of the PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>TRUE, if the PowerTalk mailer has been received from another user (as opposed to just being added but not sent). FALSE, if the mailer has not been received from another user.</td>
</tr>
<tr class="even">
<td>31</td>
<td>The date and time the PowerTalk mailer was sent, as a serial number. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The sender name of the PowerTalk mailer, as text. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="even">
<td>33</td>
<td>The title of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>The subject of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>35</td>
<td>The author of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>The keywords for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>37</td>
<td>The comments for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>The name of the active sheet.</td>
</tr>
</tbody>
</table>

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

**Example**

The following macro formula returns the name of the active sheet in the
workbook named SALES.XLS:

GET.WORKBOOK(38, "SALES.XLS")

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

WORKBOOK.SELECT   Selects the specified documents in a workbook

Return to [top](#E)

# GET.WORKSPACE

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

Return to [top](#E)

# GOAL.SEEK

Equivalent to clicking the Goal Seek command on the Tools menu.
Calculates the values necessary to achieve a specific goal. If the goal
is an amount returned by a formula, the GOAL.SEEK function calculates
values that, when supplied to your formula, cause your formula to return
the amount you want.

**Syntax**

**GOAL.SEEK**(**target\_cell, target\_value, variable\_cell**)

**GOAL.SEEK**?(target\_cell, target\_value, variable\_cell)

Target\_cell    corresponds to the Set Cell box in the Goal Seek dialog
box and is a reference to the cell containing the formula. If
target\_cell does not contain a formula, Microsoft Excel displays an
error message.

Target\_value    corresponds to the To Value box in the Goal Seek dialog
box and is the value you want the formula in target\_cell to return.
This value is called a goal.

Variable\_cell    corresponds to the By Changing Cell box in the Goal
Seek dialog box and is the single cell that you want Microsoft Excel to
change so that the formula in target\_cell returns target\_value.
Target\_cell must depend on variable\_cell; if it does not, Microsoft
Excel will not be able to find a solution.

**Remarks**

The max\_num and max\_change values set with the CALCULATION function
can be used to change the solution process. Max\_num sets the number of
iterations; max\_change determines the precision of the solution.

**Tip**   You can also use Microsoft Excel Solver to help solve your
math equations for optimal values.

**Related Functions**

Related functions include the SOLVER functions, such as SOLVER.OPTIONS,
SOLVER.SOLVE, and so on.

Return to [top](#E)

# GOTO

Directs a macro to continue running at the upper-left cell of reference.
Use GOTO to direct macro execution to another cell or a named range.

**Syntax**

**GOTO**(**reference**)

Reference    is a cell reference or a name that is defined as a
reference. Reference can be an external reference to another macro
sheet. If that macro sheet is not open, GOTO displays a message.

**Tip   **It's often preferable to use IF, ELSE, ELSE.IF, and END.IF
instead of GOTO when you want to perform multiple actions based on a
condition because the IF method makes your macros more structured.

**Examples**

If A1 contains the \#N/A error value, then when the following formula is
calculated, the macro branches to C3:

IF(ISERROR($A$1), GOTO($C$3))

You can also use macro names with GOTO statements. The following macro
formula branches macro execution to a macro named Compile:

GOTO(Compile)

Because Compile is a named range, it should not be enclosed in quotation
marks.

**Related Function**

FORMULA.GOTO   Selects a named area or reference on any open workbook

Return to [top](#E)

# GRIDLINES

Allows you to turn chart gridlines on and off.

Arguments are logical values corresponding to the check boxes in the
Gridlines dialog box. If an argument is TRUE, Microsoft Excel selects
the check box; if FALSE, Microsoft Excel clears the check box. If
omitted, the setting is not changed. If a chart is not active, produces
a error and halts the macro.

**Syntax**

**GRIDLINES**(x\_major, x\_minor, y\_major, y\_minor, z\_major,
z\_minor, 2D\_effect)

**GRIDLINES**?(x\_major, x\_minor, y\_major, y\_minor, z\_major,
z\_minor, 2D\_effect)

X\_major    corresponds to the Category (X) Axis: Major Gridlines check
box.

X\_minor    corresponds to the Category (X) Axis: Minor Gridlines check
box.

Y\_major    corresponds to the Value (Y) Axis: Major Gridlines check
box. On 3-D charts, y\_major corresponds to the Series (Y) Axis: Major
Gridlines check box.

Y\_minor    corresponds to the Value (Y) Axis: Minor Gridlines check
box. On 3-D charts, y\_minor corresponds to the Series (Y) Axis: Minor
Gridlines check box.

Z\_major    corresponds to the Value (Z) Axis: Major Gridlines check box
(3-D only).

Z\_minor    corresponds to the Value (Z) Axis: Minor Gridlines check box
(3-D only).

2D\_effect    corresponds to the 2-D Walls and Gridlines check box (3-D
only).

Return to [top](#E)

# GROUP

Creates a single object from several selected objects and returns the
object identifier of the group (for example, "Group 5"). Use GROUP to
combine a number of objects so that you can move or resize them
together.

If no object is selected, only one object is selected, or a group is
already selected, GROUP returns the \#VALUE\! error value and interrupts
the macro.

**Syntax**

**GROUP**( )

**Related Function**

UNGROUP   Separates a grouped object

Return to [top](#E)
