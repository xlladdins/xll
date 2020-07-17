<span id="Q" class="anchor"></span>This document contains reference
information on the following Excel macro functions:

# Q

[QUERY.GET.DATA](#query.get.data), [QUERY.REFRESH](#query.refresh),
[QUIT](#quit)

# R

[RANDOM](#random), [RANKPERC](#rankperc), [REFTEXT](#reftext),
[REGISTER](#register), [REGRESS](#regress), [RELREF](#relref),
[REMOVE.LIST.ITEM](#remove.list.item),
[REMOVE.PAGE.BREAK](#remove.page.break),
[RENAME.COMMAND](#rename.command), [RENAME.OBJECT](#rename.object),
[REPLACE.FONT](#replace.font), [REPORT.DEFINE](#report.define),
[REPORT.DELETE](#report.delete), [REPORT.GET](#report.get),
[REPORT.PRINT](#report.print), [REQUEST](#request),
[RESET.TOOL](#reset.tool), [RESET.TOOLBAR](#reset.toolbar),
[RESTART](#restart), [RESULT](#result), [RESUME](#resume),
[RETURN](#return), [ROUTE.DOCUMENT](#route.document),
[ROUTING.SLIP](#routing.slip), [ROW.HEIGHT](#row.height), [RUN](#run)

# S

[SAMPLE](#sample), [SAVE](#save), [SAVE.AS](#save.as),
[SAVE.COPY.AS](#save.copy.as), [SAVE.DIALOG](#save.dialog),
[SAVE.TOOLBAR](#save.toolbar), [SAVE.WORKBOOK](#save.workbook),
[SAVE.WORKSPACE](#save.workspace), [SCALE](#scale), [SCALE Syntax
1](#scale-syntax-1), [SCALE Syntax 2](#scale-syntax-2), [SCALE Syntax
3](#scale-syntax-3), [SCALE Syntax 4](#scale-syntax-4), [SCALE Syntax
5](#scale-syntax-5), [SCENARIO.ADD](#scenario.add),
[SCENARIO.CELLS](#scenario.cells), [SCENARIO.DELETE](#scenario.delete),
[SCENARIO.EDIT](#scenario.edit), [SCENARIO.GET](#scenario.get),
[SCENARIO.MERGE](#scenario.merge), [SCENARIO.SHOW](#scenario.show),
[SCENARIO.SHOW.NEXT](#scenario.show.next),
[SCENARIO.SUMMARY](#scenario.summary),
[SCROLLBAR.PROPERTIES](#scrollbar.properties), [SELECT](#select),
[SELECT Syntax 1](#select-syntax-1), [SELECT Syntax
2](#select-syntax-2), [SELECT Syntax 3](#select-syntax-3),
[SELECT.ALL](#select.all), [SELECT.CHART](#select.chart),
[SELECT.END](#select.end), [SELECTION](#selection),
[SELECT.LAST.CELL](#select.last.cell),
[SELECT.LIST.ITEM](#select.list.item),
[SELECT.PLOT.AREA](#select.plot.area),
[SELECT.SPECIAL](#select.special), [SEND.KEYS](#send.keys),
[SEND.MAIL](#send.mail), [SEND.TO.BACK](#send.to.back),
[SERIES](#series), [SERIES.AXES](#series.axes),
[SERIES.ORDER](#series.order), [SERIES.X](#series.x),
[SERIES.Y](#series.y), [SET.CONTROL.VALUE](#set.control.value),
[SET.CRITERIA](#set.criteria), [SET.DATABASE](#set.database),
[SET.DIALOG.DEFAULT](#set.dialog.default),
[SET.DIALOG.FOCUS](#set.dialog.focus), [SET.EXTRACT](#set.extract),
[SET.LIST.ITEM](#set.list.item), [SET.NAME](#set.name),
[SET.PAGE.BREAK](#set.page.break), [SET.PREFERRED](#set.preferred),
[SET.PRINT.AREA](#set.print.area),
[SET.PRINT.TITLES](#set.print.titles),
[SET.UPDATE.STATUS](#set.update.status), [SET.VALUE](#set.value),
[SHORT.MENUS](#short.menus), [SHOW.ACTIVE.CELL](#show.active.cell),
[SHOW.BAR](#show.bar), [SHOW.CLIPBOARD](#show.clipboard),
[SHOW.DETAIL](#show.detail), [SHOW.DIALOG](#show.dialog),
[SHOW.INFO](#show.info), [SHOW.LEVELS](#show.levels),
[SHOW.TOOLBAR](#show.toolbar), [SIZE](#size),
[SLIDE.COPY.ROW](#slide.copy.row), [SLIDE.CUT.ROW](#slide.cut.row),
[SLIDE.DEFAULTS](#slide.defaults),
[SLIDE.DELETE.ROW](#slide.delete.row), [SLIDE.EDIT](#slide.edit),
[SLIDE.GET](#slide.get), [SLIDE.PASTE](#slide.paste),
[SLIDE.PASTE.ROW](#slide.paste.row), [SLIDE.SHOW](#slide.show),
[SOLVER.ADD](#solver.add), [SOLVER.CHANGE](#solver.change),
[SOLVER.DELETE](#solver.delete), [SOLVER.FINISH](#solver.finish),
[SOLVER.GET](#solver.get), [SOLVER.LOAD](#solver.load),
[SOLVER.OK](#solver.ok), [SOLVER.OPTIONS](#solver.options),
[SOLVER.RESET](#solver.reset), [SOLVER.SAVE](#solver.save),
[SOLVER.SOLVE](#solver.solve), [SORT](#sort), [SOUND.NOTE](#sound.note),
[SOUND.PLAY](#sound.play), [SPELLING](#spelling),
[SPELLING.CHECK](#spelling.check), [SPLIT](#split),
[SQL.BIND](#sql.bind), [SQL.CLOSE](#sql.close), [SQL.ERROR](#sql.error),
[SQL.EXEC.QUERY](#sql.exec.query), [SQL.GET.SCHEMA](#sql.get.schema),
[SQL.OPEN](#sql.open), [SQL.RETRIEVE](#sql.retrieve),
[SQL.RETRIEVE.TO.FILE](#sql.retrieve.to.file),
[STANDARD.FONT](#standard.font), [STANDARD.WIDTH](#standard.width),
[STEP](#step), [STYLE](#style), [SUBSCRIBE.TO](#subscribe.to),
[SUBTOTAL.CREATE](#subtotal.create),
[SUBTOTAL.REMOVE](#subtotal.remove), [SUMMARY.INFO](#summary.info)

# QUERY.GET.DATA

Builds a new query using the supplied information. The application
Microsoft Query nor any dialog boxes are displayed.

**Syntax**

**QUERY.GET.DATA**(**connection\_string**, **query\_text**,
keep\_query\_def, field\_names, row\_numbers, destination)

**QUERY.GET.DATA**?(connection\_string, query\_text, keep\_query\_def,
field\_names, row\_numbers, destination)

Connection\_string&nbsp;&nbsp;&nbsp;&nbsp;supplies information, such as
the data source name, user ID, and passwords, necessary to making a SQL
connection to an external data source. For example: "DSN=Myserver;
Server=server1; UID=dbayer; PWD=buyer1; Database=nwind".

&nbsp;

  - > You must define the data source name (DSN) used in
    > connection\_string before you try to connect to it.

  - > You can enter connection\_string as an array or a string. If
    > connection\_string exceeds 250 characters, you must enter it as an
    > array.

  - > If QUERY.GET.DATA is unable to access the data source using
    > connection\_string, it returns the \#N/A error value.

> &nbsp;

Query\_text&nbsp;&nbsp;&nbsp;&nbsp;is the SQL language query to be
executed on the data source.

Keep\_query\_def&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that, if TRUE
or omitted, preserves the query definition. If FALSE, the query
definition is lost and the data from the query no longer constitutes a
data range.

Field\_names&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that, if TRUE or
omitted, places field names from Microsoft Query into the first row of
the data range. If FALSE, the field names are discarded.

Row\_numbers&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that, if TRUE,
places row numbers from Microsoft Query into the first column in the
data range. If FALSE or omitted, the row numbers are discarded.

Destination&nbsp;&nbsp;&nbsp;&nbsp;is the location as a cell reference
where you want the data placed. If destination is in a current data
range then that data range is changed to reflect the new SQL statement.
The default destination is the currently selected cell or range.

**Remarks**

  - > If the information provided is not sufficient to create the query
    > then the error value \#REF\! is returned.

  - > If Microsoft Query is unavailable or can not be found, \#N/A is
    > returned.

  - > If connection string is longer than 255 characters, the string
    > will be truncated at the last semi-colon.

**Related Function**

QUERY.REFRESH&nbsp;&nbsp;&nbsp;Refreshes the data in a data range
returned by Microsoft Query

Return to [top](#Q)

# QUERY.REFRESH

Refreshes the data in a data range returned to a worksheet from
Microsoft Query. This function is equivalent to the Refresh button on
the External Data toolbar.

**Syntax**

**QUERY.REFRESH**(reference)

Reference&nbsp;&nbsp;&nbsp;&nbsp;is the reference to a single cell
inside a data range. If reference is not in a data range then the error
value \#REF\! is returned.

**Related Function**

QUERY.GET.DATA&nbsp;&nbsp;&nbsp;Builds a new query using the supplied
information

Return to [top](#Q)

# QUIT

Equivalent to clicking the Exit command on the File menu in Microsoft
Excel for Windows. Equivalent to clicking the Quit command on the File
menu in Microsoft Excel for the Macintosh. Quits Microsoft Excel and
closes any open workbooks. If open workbooks have unsaved changes,
Microsoft Excel displays a message asking if you want to save them. You
can use QUIT in an Auto\_Close macro to force Microsoft Excel to quit
when a particular workbook is closed.

**Syntax**

**QUIT**( )

**Caution&nbsp;&nbsp;&nbsp;**If you have cleared error-checking with an
ERROR(FALSE) function, QUIT will not ask whether you want to save
changes.

**Remarks**

When you use the QUIT function, Microsoft Excel does not run any
Auto\_Close macros before closing the workbook.

**Examples**

The following function displays a confirmation alert and quits Microsoft
Excel if the user clicks OK:

IF(ALERT("Are you sure you want to quit Microsoft Excel?",1), QUIT(),)

**Related Function**

FILE.CLOSE&nbsp;&nbsp;&nbsp;Closes the active workbook

Return to [top](#Q)

# RANDOM

Fills a range with independent random or patterned numbers drawn from
one of several distributions.

If this function is not available, you must install the Analysis ToolPak
add-in.

RANDOM provides six different random distributions and one patterned
data option. Because the distributions require different argument lists,
there are seven syntax forms for RANDOM.

**Syntax 1**

Uniform distribution

**RANDOM**(outrng, variables, points, **distribution**, seed, **from,
to**)

**RANDOM**?(outrng, variables, points, distribution, seed, from, to)

**Syntax 2**

Normal distribution

**RANDOM**(outrng, variables, points, **distribution**, seed, **mean,
standard\_dev**)

**RANDOM**?(outrng, variables, points, distribution, seed, mean,
standard\_dev)

**Syntax 3**

Bernoulli distribution

**RANDOM**(outrng, variables, points, **distribution**, seed,
**probability**)

**RANDOM**?(outrng, variables, points, distribution, seed, probability)

**Syntax 4**

Binomial distribution

**RANDOM**(outrng, variables, points, **distribution**, seed,
**probability, trials**)

**RANDOM**?(outrng, variables, points, distribution, seed, probability,
trials)

**Syntax 5**

Poisson distribution

**RANDOM**(outrng, variables, points, **distribution**, seed,
**lambda**)

**RANDOM**?(outrng, variables, points, distribution, seed, lambda)

**Syntax 6**

Patterned distribution

**RANDOM**(outrng, variables, points, **distribution**, seed, **from,
to, step, repeat\_num, repeat\_seq**)

**RANDOM**?(outrng, variables, points, distribution, seed, from, to,
step, repeat\_num, repeat\_seq)

**Syntax 7**

Discrete distribution

**RANDOM**(outrng, variables, points, **distribution**, seed,
**inprng**)

**RANDOM**?(outrng, variables, points, distribution, seed, inprng)

Outrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell) in
the output table or the name, as text, of a new sheet to contain the
output table. If FALSE, blank, or omitted, places the output table in a
new workbook.

Variables&nbsp;&nbsp;&nbsp;&nbsp;is the number of random number sets to
generate. RANDOM will generate variables columns of random numbers. If
omitted, variables is equal to the number of columns in the output
range.

Points&nbsp;&nbsp;&nbsp;&nbsp;is the number of data points per random
number set. RANDOM will generate points rows of random numbers for each
random number set. If omitted, points is equal to the number of rows in
the output range. Points is ignored when distribution is 6 (Patterned).

Distribution&nbsp;&nbsp;&nbsp;&nbsp;indicates the type of number
distribution.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Distribution</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Distribution type</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Uniform</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Normal</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Bernoulli</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Binomial</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Poisson</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Patterned</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Discrete</p>
</blockquote></td>
</tr>
</tbody>
</table>

Seed&nbsp;&nbsp;&nbsp;&nbsp;is an optional value with which to begin
random number generation. Seed is ignored when distribution is 6
(Patterned) or 7 (Discrete).

From&nbsp;&nbsp;&nbsp;&nbsp;is the lower bound.

To&nbsp;&nbsp;&nbsp;&nbsp;is the upper bound.

Mean&nbsp;&nbsp;&nbsp;&nbsp;is the mean.

Standard\_dev&nbsp;&nbsp;&nbsp;&nbsp;is the standard deviation.

Probability&nbsp;&nbsp;&nbsp;&nbsp;is the probability of success on each
trial.

Trials&nbsp;&nbsp;&nbsp;&nbsp;is the number of trials.

Lambda&nbsp;&nbsp;&nbsp;&nbsp;is the Poisson distribution parameter.

Step&nbsp;&nbsp;&nbsp;&nbsp;is the increment between from and to.

Repeat\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number of times to repeat each
value.

Repeat\_seq&nbsp;&nbsp;&nbsp;&nbsp;is the number of times to repeat each
sequence of values.

Inprng&nbsp;&nbsp;&nbsp;&nbsp;is a two-column range of values and their
probabilities.

Return to [top](#Q)

# RANKPERC

Returns a table that contains the ordinal and percent rank of each value
in a data set.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**RANKPERC**(**inprng**, outrng, grouped, labels)

**RANKPERC**?(inprng, outrng, grouped, labels)

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

> &nbsp;

Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that describes where
the labels are located in the input range, as shown in the following
table:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Labels</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Grouped</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Labels are in</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>TRUE</p>
</blockquote></td>
<td><blockquote>
<p>"C"</p>
</blockquote></td>
<td><blockquote>
<p>First row of the input range.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>TRUE</p>
</blockquote></td>
<td><blockquote>
<p>"R"</p>
</blockquote></td>
<td><blockquote>
<p>First column of the input range.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>FALSE or omitted</p>
</blockquote></td>
<td><blockquote>
<p>(ignored)</p>
</blockquote></td>
<td><blockquote>
<p>No labels. All cells in the input range are data.</p>
</blockquote></td>
</tr>
</tbody>
</table>

Return to [top](#Q)

# REFTEXT

Converts a reference to an absolute reference in the form of text. Use
REFTEXT when you need to manipulate references with text functions.
After manipulating the reference text, you can convert it back into a
normal reference by using TEXTREF.

**Syntax**

**REFTEXT**(**reference**, a1)

Reference&nbsp;&nbsp;&nbsp;&nbsp;is the reference you want to convert.

A1&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying A1-style or
R1C1-style references.

  - > If a1 is TRUE, REFTEXT returns an A1-style reference.

  - > If a1 is FALSE or omitted, REFTEXT returns an R1C1-style
    > reference.

> &nbsp;

**Examples**

REFTEXT(C3, TRUE) equals "$C$3"

REFTEXT(B2:F2) equals "R2C2:R2C6"

If the active cell is B9 on the active sheet named SHEET1, then:

REFTEXT(ACTIVE.CELL()) equals "\[Book1\]SHEET1\!R9C2"

REFTEXT(ACTIVE.CELL(), TRUE) equals "\[Book1\]SHEET1\!$B$9"

**Related Functions**

ABSREF&nbsp;&nbsp;&nbsp;Returns the absolute reference of a range of
cells to another range

DEREF&nbsp;&nbsp;&nbsp;Returns the values of cells in the reference

RELREF&nbsp;&nbsp;&nbsp;Returns a relative reference

TEXTREF&nbsp;&nbsp;&nbsp;Converts text to a reference

Return to [top](#Q)

# REGISTER

Registers the specified dynamic link library (DLL) or code resource and
returns the register ID. You can also specify a custom function name and
argument names that will appear in the Paste Function dialog box. If you
register a command (macro\_type = 2), you can also specify a shortcut
key. Because Microsoft Excel for Windows and Microsoft Excel for the
Macintosh use different types of code resources, REGISTER has a slightly
different syntax form when used in each operating environment.

**Important**&nbsp;&nbsp; This function is provided for advanced users
only. If you use the CALL function incorrectly, you could cause errors
that will require you to restart your computer.

**Syntax 1**

For Microsoft Excel for Windows

**REGISTER**(**module\_text**, procedure, type\_text, function\_text,
argument\_text, macro\_type, category, shortcut\_text, help\_topic,
function\_help, argument\_help1, argument\_help2,...)

**Syntax 2**

For Microsoft Excel for the Macintosh

**REGISTER**(**file\_text**, resource, type\_text, function\_text,
argument\_text, macro\_type, category, shortcut\_text, help\_topic,
function\_help, argument\_help1, argument\_help2,...)

Module\_text or file\_text&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the
name of the DLL that contains the function (in Microsoft Excel for
Windows) or the name of the file that contains the code resource (in
Microsoft Excel for the Macintosh).

Procedure or resource&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the name
of the function in the DLL (in Microsoft Excel for Windows) or the name
of the code resource (in Microsoft Excel for the Macintosh). In
Microsoft Excel for Windows, you can also use the ordinal value of the
function from the EXPORTS statement in the module-definition file
(.DEF). In Microsoft Excel for the Macintosh, you can also use the
resource ID number. The ordinal value or resource ID number should not
be in text form.

This argument may be omitted for stand-alone DLLs or code resources. In
this case, REGISTER will register all functions or code resources and
then return module\_text or file\_text.

Type\_text&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the data type of
the return value and the data types of all arguments to the DLL or code
resource. The first letter of type\_text specifies the return value.

Function\_text&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the name of the
function as you want it to appear in the Paste Function dialog box. If
you omit this argument, the function will not appear in the Paste
Function dialog box.

Argument\_text&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the names of
the arguments you want to appear in the Paste Function dialog box.
Argument names should be separated by commas.

Macro\_type&nbsp;&nbsp;&nbsp;&nbsp;specifies the macro type: 1 for a
function or 2 for a command. If macro\_type is omitted, it is assumed to
be 1 (function).

Category&nbsp;&nbsp;&nbsp;&nbsp;specifies the function category in the
Paste Function dialog box in which you want the registered function to
appear. You can use the category number or the category name for
category. If you use the category name, be sure to enclose it in double
quotation marks. If category is omitted, it is assumed to be 14 (User
Defined).

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Category number</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Category name</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Financial</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Date &amp; Time</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Math &amp; Trig</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Text</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Logical</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Lookup &amp; Matrix</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Database</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Statistical</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>Information</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>Commands (macro sheets only)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>Actions (macro sheets only)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>Customizing (macro sheets only)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>Macro Control (macro sheets only)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>User Defined</p>
</blockquote></td>
</tr>
</tbody>
</table>

Shortcut\_text&nbsp;&nbsp;&nbsp;&nbsp;is a character specifying the
shortcut key for the registered command. The shortcut key is
case-sensitive. This argument is used only if macro\_type = 2 (command).
If shortcut\_text is omitted, the command will not have a shortcut key.

Help\_topic&nbsp;&nbsp;&nbsp;&nbsp;is the reference (including path) to
the help file that you want displayed when the user clicks the Help
button when your custom function is displayed.

Function\_help&nbsp;&nbsp;&nbsp;&nbsp;is a text string describing your
custom function when it is selected in the Paste Function dialog box.
The maximum number of characters is 255.

Argument\_help1, argument\_help2&nbsp;&nbsp;&nbsp;&nbsp;are 1 to 21 text
strings that describes you custom function's arguments when the function
is selected in the Paste Function dialog box.

Example&nbsp;&nbsp;&nbsp;&nbsp;  
Syntax 1

In Microsoft Excel for Windows, the following macro formula registers
the GetTickCount function from Microsoft Windows. This function returns
the number of milliseconds that have elapsed since Microsoft Windows was
started.

REGISTER("User", "GetTickCount", "J")

Assuming that the REGISTER function is in cell A5, after your macro
registers GetTickCount, you can use the CALL function to return the
number of milliseconds that have elapsed:

CALL(A5)

Example&nbsp;&nbsp;&nbsp;&nbsp;  
Syntax 1 with optional function\_text

You can use the following macro formula to register the GetTickCount
function from Microsoft Windows and assign the custom name GetTicks to
it. To do this, include "GetTicks" as the optional function\_text
argument to the REGISTER function.

REGISTER("User", "GetTickCount", "J", "GetTicks", , 1, 9)

After the function is registered, the custom name GetTicks will appear
in the Information function category (category = 9) in the Paste
Function dialog box.

You can call the function from the same macro sheet on which it was
registered using the following formula:

GetTicks()

You can call the function from another sheet or macro sheet by including
the name of the original macro sheet in the formula. For example,
assuming the macro sheet on which GetTicks was registered is named
MACRO1.XLS, the following formula calls the function from another sheet:

MACRO1.XLS\!GetTicks()

**Tip&nbsp;&nbsp;&nbsp;**You can use functions in a DLL or code resource
directly on a sheet without first registering them from a macro sheet.
Use syntax 2a or 2b of the CALL function. For more information, see
CALL.

**Related Function**

UNREGISTER&nbsp;&nbsp;&nbsp;Removes a registered code resource from
memory

Return to [top](#Q)

# REGRESS

Performs multiple linear regression analysis.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**REGRESS**(**inpyrng, inpxrng**, constant, labels, confid, soutrng,
residuals, sresiduals, rplots, lplots, routrng, nplots, poutrng)

**REGRESS**?(inpyrng, inpxrng, constant, labels, confid, soutrng,
residuals, sresiduals, rplots, lplots, routrng, nplots, poutrng)

Inpyrng&nbsp;&nbsp;&nbsp;&nbsp;is the input range for the y-values
(dependent variable).

Inpxrng&nbsp;&nbsp;&nbsp;&nbsp;is the input range for the x-values
(independent variable).

Constant&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If constant is TRUE,
the y-intercept is assumed to be zero (the regression line passes
through the origin). If constant is FALSE or omitted, the y-intercept is
assumed to be a non-zero number.

Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value.

  - > If labels is TRUE, then the first row or column of the input
    > ranges contain labels.

  - > If labels is FALSE or omitted, all cells in inpyrng and inpxrng
    > are considered data. Microsoft Excel will then generate the
    > appropriate data labels for the output table.

> &nbsp;

Confid&nbsp;&nbsp;&nbsp;&nbsp;is an additional confidence level to apply
to the regression. If omitted, confid is 95%.

Soutrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell)
in the output table or the name, as text, of the new sheet to contain
the summary output table. If FALSE, blank, or omitted, places the
summary output table in a new workbook. Microsoft Excel version 5.0 uses
a single output table for REGRESS; Microsoft Excel version 4.0 used
three different output tables for summary, residual, and probability
data.

Residuals&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If residuals is
TRUE, REGRESS includes residuals in the output table. If residuals is
FALSE or omitted, residuals are not included.

Sresiduals&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If sresiduals is
TRUE, REGRESS includes standardized residuals in the output table. If
sresiduals is FALSE or omitted, standardized residuals are not included.

Rplots&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If rplots is TRUE,
REGRESS generates separate charts for each x versus the residual. If
rplots is FALSE or omitted, separate charts are not generated.

Lplots&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If lplots is TRUE,
REGRESS generates a chart showing the regression line fitted to the
observed values. If lplots is FALSE or omitted, the chart is not
generated.

Routrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell)
in the residuals output table or the name, as text, of the new sheet to
contain the residuals output table. If FALSE, blank, or omitted, places
the residuals output table in a new worksheet. This argument is for
compatibility with Microsoft Excel version 4.0 only and is ignored in
later versions.

Nplots&nbsp;&nbsp;&nbsp;&nbsp;is a logical value. If nplots is TRUE,
REGRESS generates a chart of normal probabilities. If nplots is FALSE or
omitted, the chart is not generated.

Poutrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell)
in the probability data output table or the name, as text, of the new
sheet to contain the probability output table. If FALSE, blank, or
omitted, places the probability output table in a new worksheet. This
argument is for compatibility with Microsoft Excel version 4.0 only and
is ignored in later versions.

Return to [top](#Q)

# RELREF

Returns the reference of a cell or cells relative to the upper-left cell
of rel\_to\_ref. The reference is given as an R1C1-style relative
reference in the form of text, such as "R\[1\]C\[1\]".

**Syntax**

**RELREF**(**reference, rel\_to\_ref**)

Reference&nbsp;&nbsp;&nbsp;&nbsp;is the cell or cells to which you want
to create a relative reference.

Rel\_to\_ref&nbsp;&nbsp;&nbsp;&nbsp;is the cell from which you want to
create the relative reference.

**Tip**&nbsp;&nbsp;&nbsp;If you know the absolute reference of a cell
that you want to include in a formula, but your formula requires a
relative reference, use RELREF to generate the relative reference. This
is especially useful with the FORMULA function, since its formula\_text
argument requires R1C1-style references, and RELREF returns relative
R1C1-style references. You can also use the FORMULA.CONVERT function to
convert absolute references to relative references.

**Examples**

RELREF($A$1, $C$3) equals "R\[-2\]C\[-2\]"

RELREF($A$1:$E$5, $C$3:$G$7) equals "R\[-2\]C\[-2\]:R\[2\]C\[2\]"

RELREF($A$1:$E$5, $C$3) equals "R\[-2\]C\[-2\]:R\[2\]C\[2\]"

**Related Functions**

ABSREF&nbsp;&nbsp;&nbsp;Returns the absolute reference of a range of
cells to another range

DEREF&nbsp;&nbsp;&nbsp;Returns the value of the cells in the reference

FORMULA&nbsp;&nbsp;&nbsp;Enters values into a cell or range or onto a
chart

FORMULA.CONVERT&nbsp;&nbsp;&nbsp;Changes the reference style and type

Return to [top](#Q)

# REMOVE.LIST.ITEM

Removes an item in a list box or drop-down box.

**Syntax**

**REMOVE.LIST.ITEM**(**index\_num**, count\_num)

Index\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies the index of the item to
remove, from 1 to the number of items in the list. Specify zero to
remove all items in the list.

Count\_num&nbsp;&nbsp; Specifies the number of items to delete starting
from index\_num. If omitted, only one item is removed.

**Remarks**

If count\_num + index\_num is greater than the number of items in the
list, all items starting with index\_num to the end of the list are
removed.

**Examples**

REMOVE.LIST.ITEM(3,2) removes two items starting with the third item

REMOVE.LIST.ITEM(3) removes only the third item

**Related Function**

LISTBOX.PROPERTIES&nbsp;&nbsp;&nbsp;Sets the properties of a list box
and drop-down controls on worksheet and dialog sheets

Return to [top](#Q)

# REMOVE.PAGE.BREAK

Equivalent to clicking the Remove Page Break command on the Insert menu.
Removes manual page breaks that you set with the SET.PAGE.BREAK function
or the Page Break command on the Insert menu. If the active cell is not
below or to the right of a manual page break, REMOVE.PAGE.BREAK takes no
action. If the entire sheet is selected, REMOVE.PAGE.BREAK removes all
manual page breaks. REMOVE.PAGE.BREAK does not remove automatic page
breaks.

**Syntax**

**REMOVE.PAGE.BREAK**( )

**Related Function**

SET.PAGE.BREAK&nbsp;&nbsp;&nbsp;Sets manual page breaks

Return to [top](#Q)

# RENAME.COMMAND

Changes the name of a built-in or custom menu command or the name of a
menu. Use RENAME.COMMAND to change the name of a command on a menu, for
example, when you create two custom commands that toggle on the menu.
Examples of two built-in commands that toggle are the Page Break and
Remove Page Break commands on the Insert menu.

**Syntax**

**RENAME.COMMAND**(**bar\_num, menu, command, name\_text**, position)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;can be either the number of one of the
Microsoft Excel built-in menu bars or the number returned by a
previously run ADD.BAR function. See ADD.COMMAND for a list of ID
numbers for built-in menu bars.

Menu&nbsp;&nbsp;&nbsp;&nbsp;can be either the name of a menu as text or
the number of a menu. Menus are numbered starting with 1 from the left
of the screen.

Command&nbsp;&nbsp;&nbsp;&nbsp;can be either the name of the command as
text or the number of the command to be renamed (the first command on a
menu is command 1). If command is 0, RENAME.COMMAND renames the menu
instead of the command. Because other macros can change the position of
custom menu commands, you should use the name of the command rather than
a number whenever possible.

If the specified menu bar, menu, or command does not exist,
RENAME.COMMAND returns the \#VALUE\! error value and interrupts the
macro.

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the new name for the command.

Position&nbsp;&nbsp;&nbsp;&nbsp;is the name of the command on a submenu
that you want to rename. If you use position, you must use command as
the name of the submenu.

**Tip**&nbsp;&nbsp;&nbsp;To specify an access key for the new name,
precede the character you want to use with an ampersand (&). The access
key is indicated by an underline under one letter of a menu or command
name. In Microsoft Excel for the Macintosh, you can use the General tab
in the Options dialog box to turn command underlining on or off. To see
the Options dialog box, click Options on the Tools menu.

**Example**

To rename the Save All command as Global Save, and to make the letter
"G" in Global Save an access key, use the following macro formula:

RENAME.COMMAND(10, "File", "Save All", "\&Global Save")

**Related Functions**

ADD.BAR&nbsp;&nbsp;&nbsp;Adds a menu bar

ADD.COMMAND&nbsp;&nbsp;&nbsp;Adds a command to a menu

CHECK.COMMAND&nbsp;&nbsp;&nbsp;Adds or deletes a check mark to or from a
command

DELETE.COMMAND&nbsp;&nbsp;&nbsp;Deletes a command from a menu

ENABLE.COMMAND&nbsp;&nbsp;&nbsp;Enables or disables a menu or custom
command

Return to [top](#Q)

# RENAME.OBJECT

Renames the selected object or group. This is useful for giving objects
names more relevant to their usage. This is also useful if it is
uncertain how the object may have been named.

**Syntax**

**RENAME.OBJECT**(new\_name)

New\_name&nbsp;&nbsp;&nbsp;&nbsp;is the new name to be given to the
selected object.

**Related Functions**

GET.OBJECT&nbsp;&nbsp;&nbsp;Returns information about a specified object

INSERT.OBJECT&nbsp;&nbsp;&nbsp;Equivalent to clicking the Object command
on the Insert menu

SELECT Syntax 2&nbsp;&nbsp;&nbsp;Selects objects on worksheets

Return to [top](#Q)

# REPLACE.FONT

Replaces one of the four built-in fonts in Microsoft Excel for Windows
version 2.1 or earlier with a new font and style. This function is
included only for macro compatibility. To change the font of the
selected cell or range as part of a macro, use the FONT.PROPERTIES
function.

**Syntax**

**REPLACE.FONT**(font\_num, name\_text, size\_num, bold, italic,
underline, strike, color, outline, shadow)

**Related Function**

FONT.PROPERTIES&nbsp;&nbsp;&nbsp;Sets various font attributes

Return to [top](#Q)

# REPORT.DEFINE

Equivalent to clicking the Report Manager command on the View menu and
then clicking the Add option in the Report Manager dialog box. Creates
or replaces a report definition. If this function is not available, you
must install the Report Manager add-in.

**Syntax**

**REPORT.DEFINE**(**report\_name, sections\_array**, pages\_logical)

Report\_name&nbsp;&nbsp;&nbsp;&nbsp;specifies the name of the report. If
the workbook already contains a report with report\_name, the new report
replaces the existing one.

Sections\_array&nbsp;&nbsp;&nbsp;&nbsp;is an array that contains one or
more rows of view, scenario, and sheet name that define the report. The
sheet name is the sheet on which the view and scenario are defined. If
the sheet name is not specified, the current sheet is used when
REPORT.DEFINE is run.

Pages\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that, if TRUE
or omitted, specifies continuous page numbers for multiple sections or,
if FALSE, resets page numbers to 1 for each new section.

**Remarks**

  - > REPORT.DEFINE returns the \#VALUE error value if report\_name is
    > invalid or if the workbook is protected.

  - > If there are no reports defined, this function will bring up the
    > Add Report dialog box.

**Related Functions**

REPORT.DELETE&nbsp;&nbsp;&nbsp;Removes a report from the active workbook

REPORT.PRINT&nbsp;&nbsp;&nbsp;Prints a report

REPORT.GET&nbsp;&nbsp;&nbsp;Returns information about reports defined
for the active workbook

Return to [top](#Q)

# REPORT.DELETE

Equivalent to clicking the Report Manager command on the View menu and
then selecting a report in the Report Manager dialog box and clicking
the Delete button. Removes a report definition from the active workbook.

If this function is not available, you must install the Report Manager
add-in.

**Syntax**

**REPORT.DELETE**(**report\_name**)

Report\_name&nbsp;&nbsp;&nbsp;&nbsp;specifies the name of the report to
be removed. Report\_name can be any text that does not contain quotation
marks.

**Remarks**

REPORT.DELETE returns the \#VALUE error value if report\_name is invalid
or if the workbook is protected.

**Related Functions**

REPORT.DEFINE&nbsp;&nbsp;&nbsp;Creates a report

REPORT.PRINT&nbsp;&nbsp;&nbsp;Prints a report

REPORT.GET&nbsp;&nbsp;&nbsp;Returns information about reports defined
for the active workbook

Return to [top](#Q)

# REPORT.GET

Returns information about reports defined for the active workbook. Use
REPORT.GET to return information you can use in other macro commands
that manipulate reports.

If this function is not available, you must install the Report Manager
add-in.

**Syntax**

**REPORT.GET**(**type\_num**, report\_name)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 specifying the
type of information to return, as shown in the following table.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>An array of reports from all sheets in the active workbook or the #N/A error value if none are specified</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>An array of views, scenarios, and sheet names for the specified report in the active workbook. REPORT.GET returns the #N/A error value if the scenario check box is not selected. Returns the #VALUE! error value if name is invalid or the workbook is protected.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>If continuous page numbers are used, returns TRUE. If page numbers start at 1 for each section, returns FALSE. Returns the #VALUE! error value if report_name is invalid or the workbook is protected.</p>
</blockquote></td>
</tr>
</tbody>
</table>

Report\_name&nbsp;&nbsp;&nbsp;&nbsp;specifies the name of a report in
the active workbook.

**Remarks**

Report\_name is required if type\_num is 2 or 3.

**Example**

The following macro formula returns an array of reports from the active
workbook.

REPORT.GET(1)

**Related Functions**

REPORT.DEFINE&nbsp;&nbsp;&nbsp;Creates a report

REPORT.DELETE&nbsp;&nbsp;&nbsp;Removes a report from the active workbook

REPORT.PRINT&nbsp;&nbsp;&nbsp;Prints a report

Return to [top](#Q)

# REPORT.PRINT

Equivalent to clicking the Print button in the Report Manager dialog
box. Prints a report.

If this function is not available, you must install the Report Manager
add-in.

**Syntax**

**REPORT.PRINT**(**report\_name**, copies\_num,
show\_print\_dlg\_logical)

**REPORT.PRINT**?(report\_name, copies\_num)

Report\_name&nbsp;&nbsp;&nbsp;&nbsp;specifies the name of a report in
the active workbook.

Copies\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number of copies you want to
print. If omitted, the default is 1.

Show\_print\_dlg\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value
that, if TRUE, displays a dialog box asking how many copies to print,
or, if FALSE or omitted, prints the report immediately using existing
print settings.

**Remarks**

REPORT.PRINT returns the \#VALUE\! error value if report\_name is
invalid or if the workbook is protected.

**Related Functions**

REPORT.DEFINE&nbsp;&nbsp;&nbsp;Creates a report

REPORT.DELETE&nbsp;&nbsp;&nbsp;Removes a report from the active workbook

Return to [top](#Q)

# REQUEST

Requests an array of a specific type of information from an application
with which you have a dynamic data exchange (DDE) link. Use REQUEST with
other Microsoft Excel DDE functions to move information from another
application into Microsoft Excel.

**Syntax**

**REQUEST**(**channel\_num, item\_text**)

**Important**&nbsp;&nbsp;&nbsp;Microsoft Excel for the Macintosh
requires system software version 7.0 or later for this function.

Channel\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number returned by a previously
run INITIATE function. Channel\_num refers to a channel through which
Microsoft Excel communicates with another program.

Item\_text&nbsp;&nbsp;&nbsp;&nbsp;is a code indicating the type of
information you want to request from another application. The form of
item\_text depends on the application connected to channel\_num.

REQUEST returns the data as an array. For example, suppose the remote
data to be returned came from a sheet that looked like the following
illustration.

![](./media/image1.png)

REQUEST would return that data as the following array:

{1, 2, 3;4, 5, 6}

If REQUEST is not successful, it returns the following error values.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Value returned</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Situation</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>#VALUE!</p>
</blockquote></td>
<td><blockquote>
<p>Channel_num is not a valid channel number.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>#N/A</p>
</blockquote></td>
<td><blockquote>
<p>The application you are accessing is busy doing something else.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>#DIV/0!</p>
</blockquote></td>
<td><blockquote>
<p>The application you are accessing does not respond after a certain length of time, or you have pressed ESC or COMMAND+PERIOD to cancel.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>#REF!</p>
</blockquote></td>
<td><blockquote>
<p>The request is refused.</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Tip&nbsp;&nbsp;&nbsp;**Use the ERROR.TYPE function to distinguish
between the different error values.

**Example**

Suppose you had opened a DDE channel to Microsoft Word for Windows.
WChan contains the number of the open channel. In Microsoft Excel for
Windows, the following function returns the text specified by the
bookmark named BMK1.

\=REQUEST(WChan, "BMK1")

**Related Functions**

EXECUTE&nbsp;&nbsp;&nbsp;Carries out a command in another application

INITIATE&nbsp;&nbsp;&nbsp;Opens a channel to another application

POKE&nbsp;&nbsp;&nbsp;Sends data to another application

SEND.KEYS&nbsp;&nbsp;&nbsp;Sends a key sequence to another application

TERMINATE&nbsp;&nbsp;&nbsp;Closes a dynamic data exchange (DDE) channel
previously opened with the INITIATE function

Return to [top](#Q)

# RESET.TOOL

Resets a button to its original button face.

**Syntax**

**RESET.TOOL**(**bar\_id, position**)

Bar\_id&nbsp;&nbsp;&nbsp;&nbsp;is the number or name of the toolbar
containing the button you want to reset. For detailed information about
bar\_id, see ADD.TOOL.

Position&nbsp;&nbsp;&nbsp;&nbsp;specifies the position of the button
within the toolbar. Position starts with 1 at the left side (if
horizontal) or at the top (if vertical).

**Related Functions**

ADD.TOOL&nbsp;&nbsp;&nbsp;Adds one or more buttons to a toolbar

DELETE.TOOL&nbsp;&nbsp;&nbsp;Deletes a button from a toolbar

RESET.TOOLBAR&nbsp;&nbsp;&nbsp;Resets a button to its original button
face

Return to [top](#Q)

# RESET.TOOLBAR

Resets built-in toolbars to the default Microsoft Excel set.

**Syntax**

**RESET.TOOLBAR**(bar\_id)

Bar\_id&nbsp;&nbsp;&nbsp;&nbsp;specifies the number or name of the
toolbar that you want to reset. For detailed information about bar\_id,
see ADD.TOOL.

**Remarks**

If RESET.TOOLBAR successfully resets the toolbar, it returns TRUE. If
you try to reset a custom toolbar, RESET.TOOLBAR returns \#VALUE\! and
takes no other action.

**Related Functions**

ADD.TOOL&nbsp;&nbsp;&nbsp;Adds one or more tools to a toolbar

DELETE.TOOLBAR&nbsp;&nbsp;&nbsp;Deletes custom toolbars

Return to [top](#Q)

# RESTART

Removes a number of RETURN statements from the stack. When one macro
calls another, the RETURN statement at the end of the second macro
returns control to the calling macro. You can use the RESTART function
to determine which macro regains control.

**Syntax**

**RESTART**(level\_num)

Level\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the number of
previous RETURN statements you want to be ignored. If level\_num is
omitted, the next RETURN statement will halt macro execution.

For example, if the currently running macro has two "ancestors" (the
current macro was called by one macro that, in turn, was called by
another macro), using RESTART(1) in the third macro returns control to
the first calling macro when the RETURN statement is encountered. The
RESTART(1) formula removes one level of RETURN statements from Microsoft
Excel's memory so that the second macro is skipped.

**Remarks**

RESTART is particularly useful if you frequently use macros to call
other macros that in turn call other macros. Use RESTART in combination
with IF statements to prevent macro execution from returning to macros
that called, either directly or indirectly, the currently running macro.

**Related Functions**

HALT&nbsp;&nbsp;&nbsp;Stops all macros from running

RETURN&nbsp;&nbsp;&nbsp;Ends the currently running macro

Return to [top](#Q)

# RESULT

Specifies the type of data a macro or custom function returns. Use
RESULT to make sure your macros, custom functions, or subroutines return
values of the correct data type.

**Syntax**

**RESULT**(type\_num)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the data type.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Type of returned data</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Number</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Text</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Logical</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Reference</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>Error</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>64</p>
</blockquote></td>
<td><blockquote>
<p>Array</p>
</blockquote></td>
</tr>
</tbody>
</table>

&nbsp;

  - > Type\_num can be the sum of the numbers in the preceding table to
    > allow for more than one possible result type. For example, if
    > type\_num is 12, which equals 4 + 8, the result can be a logical
    > or a reference value.

  - > If you omit type\_num, it is assumed to be 7. Since 7 equals 1 + 2
    > + 4, the value returned can be a number (1), text (2), or logical
    > value (4).

> &nbsp;

**Examples**

The following function specifies that a custom function's return value
can be a number or a logical value (4+1=5):

RESULT(5)

**Related Functions**

ARGUMENT&nbsp;&nbsp;&nbsp;Passes an argument to a macro

RETURN&nbsp;&nbsp;&nbsp;Ends the currently running macro

Return to [top](#Q)

# RESUME

Equivalent to choosing the Resume button on the toolbar. Resumes a
paused macro. Returns TRUE if successful or the \#VALUE\! error value if
no macro is paused. A macro can be paused by using the PAUSE function or
choosing Pause from the Single Step dialog box, which appears when you
choose the Step Into button from the Macro dialog box.

**Syntax**

**RESUME**(type\_num)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 specifying how
to resume.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>How Microsoft Excel resumes</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1 or omitted</p>
</blockquote></td>
<td><blockquote>
<p>If paused by a PAUSE function, continues running the macro. If paused from the Single Step dialog box, returns to that dialog box.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Halts the paused macro</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Continues running the macro</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Opens the Single Step dialog box</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Tip**&nbsp;&nbsp;&nbsp;You can use Microsoft Excel's ON functions to
resume based on an event. For an example, see ENTER.DATA.

**Remarks**

  - > If one macro runs a second macro that pauses, and you need to halt
    > only the paused macro, use RESUME(2) instead of HALT. HALT halts
    > all macros and prevents resuming or returning to any macro.

  - > If the macro was paused from the Single Step dialog box, RESUME
    > returns to the Single Step dialog box.

> &nbsp;

**Related Functions**

HALT&nbsp;&nbsp;&nbsp;Stops all macros from running

PAUSE&nbsp;&nbsp;&nbsp;Pauses a macro

RETURN&nbsp;&nbsp;&nbsp;Ends the currently running macro

Return to [top](#Q)

# RETURN

Ends the currently running macro. If the currently running macro is a
subroutine macro that was called by another macro, control is returned
to the calling macro. If the currently running macro is a custom
function, control is returned to the formula that called the custom
function. If the currently running macro is a command macro started by
the user with the Run button in the Macro dialog box or a shortcut key
or by clicking an object, control is returned to the user.

**Syntax**

**RETURN**(value)

Value&nbsp;&nbsp;&nbsp;&nbsp;specifies what to return.

  - > If the macro is a custom function or a subroutine, value specifies
    > what value to return. However, not all subroutines return values;
    > the last line in macros that do not return values is =RETURN().

  - > If the macro is a command macro run by the user, value should be
    > omitted.

> &nbsp;

**Remarks**

RETURN signals the end of a macro. Every macro must end with a RETURN or
HALT function, but not every macro returns values.

**Example**

The following function returns the sum of the range B1:B10:

RETURN(SUM(B1:B10))

**Related Functions**

BREAK&nbsp;&nbsp;&nbsp;Interrupts a FOR-NEXT, FOR.CELL-NEXT, or
WHILE-NEXT loop

HALT&nbsp;&nbsp;&nbsp;Stops all macros from running

RESULT&nbsp;&nbsp;&nbsp;Specifies the data type a custom function
returns

Return to [top](#Q)

# ROUTE.DOCUMENT

Routes the workbook using the defined routing slip information.

**Syntax**

**ROUTE.DOCUMENT**()

**Remarks**

If there is no routing slip, returns \#N/A. If an error occurs or
routing is not enabled for the system, returns \#VALUE\!.

**Related Functions**

SEND.MAIL&nbsp;&nbsp;&nbsp;Sends the active workbook using email

ROUTING.SLIP&nbsp;&nbsp;&nbsp;Adds or Edits the routing slip attached to
the current workbook

Return to [top](#Q)

# ROUTING.SLIP

Equivalent to clicking the Add Routing Slip command on the File menu.
Adds or Edits the routing slip attached to the current workbook.

**Syntax**

**ROUTING.SLIP**(recipients,subject, message, route\_num,
return\_logical, status\_logical)

**ROUTING.SLIP**?(recipients,subject, message, route\_num,
return\_logical, status\_logical)

Recipients&nbsp;&nbsp;&nbsp;&nbsp;is the name of the person to whom you
want to send the mail. The name should be given as text.

  - > To specify more than one name, give the list of names as an array.
    > For example, ROUTING.SLIP({"John", "Paul", "George", "Ringo"})
    > would send the active workbook to the four names in the array. You
    > can also refer to a range on a sheet or macro sheet that contains
    > a list of names to whom you want the mail to be sent.

  - > Specifying recipients while a routing is in progress only modifies
    > the non-grayed recipients (that is, those recipients who have not
    > received the message yet). Recipients who have already received,
    > reviewed and forwarded the routed workbook cannot be modified.

Subject&nbsp;&nbsp;&nbsp;&nbsp;is a text string containing the subject
text used for the mail messages used to route the workbook. If omitted,
the default subject line is "Routing: name", where name is the file name
or title as displayed in the Summary Info dialog box, if available.

Message&nbsp;&nbsp;&nbsp;&nbsp;is a text string containing the body text
used for the mail messages used to route the workbook.

Route\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number indicating the type of
routing method.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Route_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Method</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1 or omitted</p>
</blockquote></td>
<td><blockquote>
<p>One after another routing</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>All at once routing</p>
</blockquote></td>
</tr>
</tbody>
</table>

Return\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value which, if TRUE
or omitted, indicates that the routing should be returned to the
originator when the routing is complete. If FALSE, the routing will end
with the last recipient in the To list box in the Routing Slip Dialog
box.

Status\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding
to the Track Status check box in the Routing Slip dialog box. If TRUE or
omitted, status tracking messages for the routing are sent. FALSE means
that no status tracking is performed.

**Remarks**

  - > If this function is used on a workbook that is already being
    > routed, the route\_num, status\_logical and return\_logical
    > arguments are ignored (they cannot be changed).

  - > When arguments are omitted and a routing slip already exists, the
    > omitted arguments are replaced by the current values of the
    > routing slip.

**Related Functions**

ROUTE.DOCUMENT&nbsp;&nbsp;&nbsp;Routes the workbook using the defined
routing slip information

SEND.MAIL&nbsp;&nbsp;&nbsp;Sends the active workbook using email

Return to [top](#Q)

# ROW.HEIGHT

Equivalent to choosing the Height command on the Row submenu of the
Format menu. Changes the height of the rows in a reference.

**Syntax**

**ROW.HEIGHT**(height\_num, reference, standard\_height, type\_num)

**ROW.HEIGHT**?(height\_num, reference, standard\_height, type\_num)

Height\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies how high you want the rows
to be in points. If standard\_height is TRUE, height\_num is ignored.

Reference&nbsp;&nbsp;&nbsp;&nbsp;specifies the rows for which you want
to change the height.

  - > If reference is omitted, the reference is assumed to be the
    > current selection.

  - > If reference is specified, it must be either an external reference
    > to the active worksheet, such as \!$2:$4 or \!Database, or an
    > R1C1-style reference in the form of text or a name, such as
    > "R1:R3", "R\[-4\]:R\[-2\]", or Database.

  - > If reference is a relative R1C1-style reference in the form of
    > text, it is assumed to be relative to the active cell.

> &nbsp;

Standard\_height&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that sets the
row height as determined by the font in each row.

  - > If standard\_height is TRUE, Microsoft Excel sets the row height
    > to a standard height that may vary from row to row depending on
    > the fonts used in each row, ignoring height\_num.

  - > If standard\_height is FALSE or omitted, Microsoft Excel sets the
    > row height according to height\_num.

> &nbsp;

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 3 corresponding
to selecting the Hide, Unhide, or AutoFit commands from the Row submenu.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Action taken</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Hides the row selection by setting the row height to 0</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Unhides the row selection by setting the row height to the value set before the selection was hidden</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Sets the row selection to an AutoFit height, which varies from row to row depending on how large the font is in any cell in each row or on how many lines of text are wrapped</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Remarks**

  - > If any of the argument settings conflict, such as when
    > standard\_height is TRUE and type\_num is 3, Microsoft Excel uses
    > the type\_num argument and ignores any arguments that conflict
    > with type\_num.

  - > If you are recording a macro while using a mouse, and you change
    > row heights by dragging the row border, Microsoft Excel records
    > the reference of the rows using R1C1-style references in the form
    > of text. If Uses Relative References is selected, Microsoft Excel
    > uses R1C1-style relative references. If Uses Relative References
    > is not selected, Microsoft Excel uses R1C1-style absolute
    > references.

> &nbsp;

**Related Function**

COLUMN.WIDTH&nbsp;&nbsp;&nbsp;Sets the widths of the specified columns

Return to [top](#Q)

# RUN

Equivalent to choosing the Run button in the Macro dialog box, which
appears when you choose the Macros command on the Macro submenu of the
Tools menu. Runs a macro.

**Syntax**

**RUN**(reference, step)

**RUN**?(reference, step)

Reference&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the macro you want to
run or a number from 1 to 4 specifying an Auto macro to run.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>If reference is</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Specifies</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>All Auto_Open macros on the active workbook</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>All Auto_Close macros</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>All Auto_Activate macros</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>All Auto_Deactivate macros</p>
</blockquote></td>
</tr>
</tbody>
</table>

&nbsp;

  - > If reference is a range of cells, RUN begins with the macro
    > function in the upper-left cell of reference.

  - > If the macro sheet containing the macro is not in the active
    > workbook, reference can be an external reference to the name of
    > the macro, such as RUN(\[BOOK1\]Macro\!Months) or an external
    > R1C1-style reference to the location of the macro, such as
    > RUN("\[Book1\]Macro\!R2C3"). The reference must be in text form.

  - > If reference is omitted, the macro function in the active cell is
    > carried out, and macro execution continues down that column.

> &nbsp;

Step&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying that the macro
is to be run in single-step mode. If step is TRUE, Microsoft Excel runs
the macro in single-step mode; if FALSE or omitted, Microsoft Excel runs
the macro normally.

**Remarks**

  - > RUN is recorded when you choose the Run button the Macro dialog
    > box while recording a macro. The reference you enter in the Run
    > dialog box is recorded as text, with A1-style references converted
    > to R1C1-style references.

  - > To run a macro from a macro sheet, you could alternatively enter
    > the name of the macro as a formula, followed by a set of
    > parentheses. For example, enter =\[Book1\]Macro\!Months() instead
    > of =RUN(\[Book1\]Macro\!Months).

> &nbsp;

**Related Function**

GOTO&nbsp;&nbsp;&nbsp;Directs macro execution to another cell

Return to [top](#Q)

# SAMPLE

Samples data.

If this function is not available, you must install the Analysis ToolPak
add-in.

**Syntax**

**SAMPLE**(**inprng**, outrng, **method, rate**, labels)

**SAMPLE**?(inprng, outrng, method, rate, labels)

Inprng&nbsp;&nbsp;&nbsp;&nbsp;is the input range.

Outrng&nbsp;&nbsp;&nbsp;&nbsp;is the first cell (the upper-left cell) in
the output column or the name, as text, of a new sheet to contain the
output column. If FALSE, blank, or omitted, places the output table in a
new workbook.

Method&nbsp;&nbsp;&nbsp;&nbsp;is a text character that indicates the
type of sampling.

  - > If method is "P", then periodic sampling is used. The input range
    > is sampled every nth cell, where n = rate.

  - > If method is "R", then random sampling is used. The output column
    > will contain rate samples.

> &nbsp;

Rate&nbsp;&nbsp;&nbsp;&nbsp;is the sampling rate, if method is "P"
(periodic sampling). Rate is the number of samples to take if method is
"R" (random sampling).

Labels&nbsp;&nbsp;&nbsp;&nbsp;is a logical value.

  - > If labels is TRUE, then the first row or column of inprng contains
    > labels.

  - > If labels is FALSE or omitted, all cells in inprng are considered
    > data. Microsoft Excel generates appropriate data labels for the
    > output table.

Return to [top](#Q)

# SAVE

Equivalent to choosing the Save command from the File menu. Saves the
active workbook.

**Syntax**

**SAVE**( )

**Remarks**

Use the SAVE.AS function instead of SAVE when you want to change the
filename or file type, specify a password, create a backup file, or save
a file to a different directory or folder.

**Related Functions**

SAVE.AS&nbsp;&nbsp;&nbsp;Saves a workbook and allows you to specify the
name, file type, password, backup file, and location of the workbook

SAVE.WORKBOOK&nbsp;&nbsp;&nbsp;Saves a workbook

Return to [top](#Q)

# SAVE.AS

Equivalent to clicking the Save As command on the File menu. Use SAVE.AS
to specify a new filename, file type, protection password, or
write-reservation password, or to create a backup file.

**Syntax**

**SAVE.AS**(document\_text, type\_num, prot\_pwd, backup,
write\_res\_pwd, read\_only\_rec)

**SAVE.AS**?(document\_text, type\_num, prot\_pwd, backup,
write\_res\_pwd, read\_only\_rec)

Document\_text&nbsp;&nbsp;&nbsp;&nbsp;specifies the name of a workbook
to save, such as SALES.XLS (in Microsoft Excel for Windows) or SALES (in
Microsoft Excel for the Macintosh). You can include a full path in
document\_text, such as C:\\EXCEL\\ANALYZE.XLS (in Microsoft Excel for
Windows) or HARDDISK:FINANCIALS:ANALYZE (in Microsoft Excel for the
Macintosh).

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the file format
in which to save the workbook.

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

Prot\_pwd&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Protection Password
box in the Save Options dialog box in Microsoft Excel 95 or earlier
versions, or the Password To Open box in Microsoft Excel 97 or later.

  - > Prot\_pwd is a password given as text or as a reference to a cell
    > containing text. Prot\_pwd should be no more than 15 characters.

  - > If a file is saved with a password, the password must be supplied
    > for the file to be opened.

> &nbsp;

Backup&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
Always Create Backup check box in the Save Options dialog box and
specifies whether to make a backup workbook. If backup is TRUE,
Microsoft Excel creates a backup file; if FALSE, no backup file is
created; if omitted, the status is unchanged.

Write\_res\_pwd&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Write
Reservation Password box in the Save Options dialog box in Microsoft
Excel 95 or earlier versions, or the Password To Modify box in Microsoft
Excel 97 or later. Allows the user to write to a file. If a file is
saved with a password and the password is not supplied when the file is
opened, the file is opened read-only.

Read\_only\_rec&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding
to the Read-Only Recommended check box in the Save Options dialog box.

  - > If read\_only\_rec is TRUE, Microsoft Excel saves the workbook as
    > a read-only recommended workbook; if FALSE, Microsoft Excel saves
    > the workbook normally; if omitted, Microsoft Excel uses the
    > current settings.

  - > When you open a workbook that was saved as read-only recommended,
    > Microsoft Excel displays a message recommending that you open the
    > workbook as read-only.

> &nbsp;

**Related Functions**

CLOSE&nbsp;&nbsp;&nbsp;Closes the active window

GET.DOCUMENT&nbsp;&nbsp;&nbsp;Returns information about a workbook

SAVE&nbsp;&nbsp;&nbsp;Saves the active workbook

SAVE.WORKBOOK&nbsp;&nbsp;&nbsp;Saves a workbook

Return to [top](#Q)

# SAVE.COPY.AS

Saves a copy of the current workbook using a different name but all the
current workbook settings, such as passwords and file protection. Does
not affect the current workbook. Use this command if you need a
temporary copy of the current workbook; for example, to include in an
electronic mail message.

**Syntax**

**SAVE.COPY.AS**(document\_text)

Document\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name you want to give the
copy of the workbook.

**Example**

Suppose that you are creating a macro that makes changes to a file
called BUDGET95.XLS. Use the following function to save a copy of this
file called TEMP.XLS without affecting BUDGET95.XLS:

SAVE.COPY.AS("temp.xls")

Return to [top](#Q)

# SAVE.DIALOG

Displays the standard Microsoft Excel File Save As dialog box and gets a
file name from the user. This function returns the path and file name of
the file that has been saved. Use SAVE.AS to automatically save a file
with a particular format and other properties.

**Syntax**

**SAVE.DIALOG**(init\_filename, title, button\_text, file\_filter,
filter\_index)

Init\_filename&nbsp;&nbsp; Specifies the suggested filename for saving.
If omitted, the active workbook's name is used, as returned by the
GET.DOCUMENT(1) function.

Title&nbsp;&nbsp;&nbsp;&nbsp;Specifies the default window title on
Microsoft Excel for Windows. For Microsoft Excel for the Macintosh,
title specifies the prompt string. If omitted, "File Save As" will be
used for Microsoft Excel for Windows, and "Save As:" For Microsoft Excel
for the Macintosh.

Button\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text used for the save button
in the dialog. If omitted, "Save" will be used as the default. This
argument is ignored on the Microsoft Excel for Windows.

File\_filter&nbsp;&nbsp;&nbsp;&nbsp;is the file filtering criteria to
use, as text. For Microsoft Excel for Windows, file\_filter consists of
two parts, a descriptive phrase denoting the file type followed by a
comma and then the MS-DOS wildcard file filter specification, as in
"Text Files (\*.TXT), \*.TXT, Add-in Files (\*.XLA), \*.XLA". Groups of
filter specifications are also separated by commas. Each separate pair
is listed in the file type drop-down list box. File\_filter can include
an asterisk (\*) to represent any sequence of characters and a question
mark (?) to represent any single character. For Microsoft Excel for the
Macintosh, file\_filter consists of file type codes separated by commas,
as in "TEXT,XLA,XLS4". Spaces are significant and should not be inserted
before or after the comma separators unless they are part of the file
type code.

Filter\_index&nbsp;&nbsp;&nbsp;&nbsp;specifies the index number of the
default file filtering criteria from 1 to the number of filters
specified in file\_filter. If omitted or greater than the number of
filters present, 1 will be used as the starting index number. The
argument is ignored on Microsoft Excel for the Macintosh.

**Remarks**

  - > To use multiple MS-DOS wildcard expressions within file\_filter
    > for a single file filter type, separate the wildcard expressions
    > with semicolons, as in "VB Files (\*.bas; \*.txt), \*.bas;\*.txt".

  - > If file\_filter is omitted, "ALL Files (\*.\*), \*.\*" will be
    > used as the default in Microsoft Excel for Windows. The default
    > for Microsoft Excel for the Macintosh is all file types.

  - > If the user cancels the dialog box, FALSE is returned.

**Examples**

SAVE.DIALOG("TRAVEL.XLS","How do you want to save this file?",,  
"Text Files (\*.TXT), \*.TXT, Add-in Files (\*.XLA), \*.XLA, ALL FILES
(\*.\*), \*.\*") opens a File Save As dialog box titled "How do you want
to save this file?", with "TRAVEL.XLS" as the suggested file name, and
with three file filter criteria in the drop-down list box.

**Related Function**

OPEN.DIALOG&nbsp;&nbsp;&nbsp;Displays the standard Microsoft Excel File
Open dialog box with the specified file filters

Return to [top](#Q)

# SAVE.TOOLBAR

Saves one or more toolbar definitions to a specified file.

**Syntax**

**SAVE.TOOLBAR**(bar\_id, filename)

Bar\_id&nbsp;&nbsp;&nbsp;&nbsp;is either the name or number of a toolbar
whose definition you want to save or an array of toolbar names or
numbers whose definitions you want to save. Use an array to save several
toolbar definitions at the same time. For detailed information about
bar\_id, see ADD.TOOL. If bar\_id is omitted, all toolbar definitions
are saved.

Filename&nbsp;&nbsp;&nbsp;&nbsp;is text specifying the name of the
destination file. If filename does not exist, Microsoft Excel creates a
new file. If filename exists, Microsoft Excel overwrites the file. If
filename is omitted, Microsoft Excel saves the toolbar or toolbars in
Username8.xlb, where "username" is your Windows or network logon name.
With Microsoft Windows, Username8.xlb is stored in the directory where
Windows is installed; with Apple Macintosh, EXCEL TOOLBARS is stored in
the System:Preferences folder

**Examples**

In Microsoft Excel for Windows, the following macro formula saves
Toolbar6 as \\EXCDT\\TOOLFILE.XLB.

SAVE.TOOLBAR("Toolbar6", "\\EXCDT\\TOOLFILE.XLB")

In Microsoft Excel for the Macintosh, the following macro formula saves
Toolbar6 as TOOLFILE.

SAVE.TOOLBAR("Toolbar6", "TOOLFILE")

**Related Functions**

ADD.TOOL&nbsp;&nbsp;&nbsp;Adds one or more tools to a toolbar

ADD.TOOLBAR&nbsp;&nbsp;&nbsp;Creates a new toolbar with the specified
tools

OPEN&nbsp;&nbsp;&nbsp;Opens a workbook

Return to [top](#Q)

# SAVE.WORKBOOK

Equivalent to clicking the Save Workbook command on the File menu in
Microsoft Excel version 4.0. Provided for compatibility with Microsoft
Excel version 4.0. Saves the workbook to which the active sheet belongs.
To save Microsoft Excel version 5.0 or later workbooks, use SAVE.AS.

**Syntax**

**SAVE.WORKBOOK**(document\_text, type\_num, prot\_pwd, backup,
write\_res\_pwd, read\_only\_rec)

**SAVE.WORKBOOK**?(document\_text, type\_num, prot\_pwd, backup,
write\_res\_pwd, read\_only\_rec)

For a description of the arguments, see SAVE.AS.

**Related Functions**

CLOSE&nbsp;&nbsp;&nbsp;Closes the active window

GET.DOCUMENT&nbsp;&nbsp;&nbsp;Returns information about a workbook

SAVE&nbsp;&nbsp;&nbsp;Saves the active workbook

SAVE.AS&nbsp;&nbsp;&nbsp;Saves a workbook and allows you to specify the
name, file type, password, backup file, and location of the workbook

Return to [top](#Q)

# SAVE.WORKSPACE

Equivalent to clicking the Save Workspace command on the File menu.
Saves the currently opened workbook or workbooks as a workspace.

**Syntax**

**SAVE.WORKSPACE**(name\_text)

**SAVE.WORKSPACE**?(name\_text)

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of the workspace to save.

**Related Function**

SAVE.AS&nbsp;&nbsp;&nbsp;Specifies a new filename.

Return to [top](#Q)

# SCALE

Changes the position, formatting, and scaling of axes in a chart. There
are five syntax forms of this function.

Syntax 1&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the category axis in 2-D charts

Syntax 2&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the value axis in 2-D charts

Syntax 3&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the category axis in 3-D charts

Syntax 4&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the series axis in 3-D charts

Syntax 5&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the value axis in 3-D charts

Return to [top](#Q)

# SCALE Syntax 1

Equivalent to clicking the Selected Axis command on the Format menu when
a chart's category (x) axis is selected, and then clicking the Scale
tab. There are five syntax forms of this function. Syntax 1 of SCALE
applies if the selected axis is a category (x) axis on a 2-D chart and
the chart is not an xy (scatter) chart. Use this syntax of SCALE to
change the position, formatting, and scaling of the category axis.

**Syntax 1**

**SCALE**(cross, cat\_labels, cat\_marks, between, max, reverse)

**SCALE**?(cross, cat\_labels, cat\_marks, between, max, reverse)

Arguments correspond to text boxes and check boxes in the Scale tab on
the Format Axis dialog box. Arguments corresponding to check boxes are
logical values. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Cross&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the Value (Y)
Axis Crosses At Category number text box. The default is 1. Cross is
ignored if max is set to TRUE.

Cat\_labels&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the
Number Of Categories Between Tick Mark Labels text box. The default is
1.

Cat\_marks&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the
Number Of Categories Between Tick Marks text box. The default is 1.

Between&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Value (Y) Axis Crosses
Between Categories check box. This argument only applies if cat\_labels
is set to a number other than 1.

Max&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Value (Y) Axis Crosses At
Maximum Category check box. If max is TRUE, it overrides any setting for
cross.

Reverse&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Categories In Reverse
Order check box.

**Related Functions**

AXES&nbsp;&nbsp;&nbsp;Controls whether axes on a chart are visible

GRIDLINES&nbsp;&nbsp;&nbsp;Controls whether chart gridlines are visible

Syntax 2&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the value axis in 2-D charts

Syntax 3&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the category axis in 3-D charts

Syntax 4&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the series axis in 3-D charts

Syntax 5&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the value axis in 3-D charts

Return to [top](#Q)

# SCALE Syntax 2

Equivalent to clicking the Selected Axes command on the Format menu when
a chart's value (y) axis is selected, and then clicking the Scale tab.
There are five syntax forms of this function. Syntax 2 of SCALE applies
if the selected axis is a value (y) axis on a 2-D chart, or either axis
on an xy (scatter) chart. Use this syntax of SCALE to change the
position, formatting, and scaling of the value axis.

**Syntax 2**

**SCALE**(min\_num, max\_num, major, minor, cross, logarithmic, reverse,
max)

**SCALE**?(min\_num, max\_num, major, minor, cross, logarithmic,
reverse, max)

The first five arguments correspond to the five range variables on the
Scale tab. Each argument can be either the logical value TRUE or a
number:

  - > If an argument is TRUE, Microsoft Excel selects the Auto check
    > box.

  - > If an argument is a number, that number is used for the variable.

> &nbsp;

Min\_num&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Minimum check box and
is the minimum value for the value axis.

Max\_num&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Maximum check box and
is the maximum value for the value axis.

Major&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Major Unit check box and
is the major unit of measure.

Minor&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Minor Unit check box and
is the minor unit of measure.

Cross&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Category (X) Axis
Crosses At text box for the value (y) axis of a 2-D chart or the Value
(Y) Axis Crosses At text box for the category (x) axis of an xy
(scatter) chart.

The last three arguments are logical values corresponding to check boxes
on the Scale tab . If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Logarithmic&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Logarithmic Scale
check box.

Reverse&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Values In Reverse
Order check box.

Max&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Category (X) Axis Crosses
At Maximum Value check box.

**Related Functions**

Syntax 1&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the category axis in 2-D charts

Syntax 3&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the category axis in 3-D charts

Syntax 4&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the series axis in 3-D charts

Syntax 5&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the value axis in 3-D charts

Return to [top](#Q)

# SCALE Syntax 3

Equivalent to clicking the Selected Axes command on the Format menu when
a chart's category (x) axis is selected, and then click the Scale tab.
There are five syntax forms of this function. Syntax 3 of SCALE applies
if the selected axis is a category (x) axis on a 3-D chart. Use this
syntax of SCALE to change the position, formatting, and scaling of the
category axis.

**Syntax 3**

**SCALE**(cat\_labels, cat\_marks, reverse, between)

**SCALE**?(cat\_labels, cat\_marks, reverse, between)

Cat\_labels&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the
Number Of Categories Between Tick-Mark Labels box. The default is 1.
Cat\_labels can also be a logical value. If TRUE, an automatic setting
will be used. If FALSE, or omitted, the number will be used.

Cat\_marks&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the
Number Of Categories Between Tick Marks text box. The default is 1.
Cat\_marks can also be a logical value. If TRUE, an automatic setting
will be used. If FALSE, or omitted, the number will be used.

Reverse&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Categories In Reverse
Order check box. If reverse is TRUE, Microsoft Excel selects the check
box; if FALSE, Microsoft Excel clears the check box.

Between&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Value (Z) Axis Crosses
Between Categories check box. If between is TRUE, Microsoft Excel
selects the check box and the data points appear between categories. If
between is FALSE or omitted, Microsoft Excel clears the check box.

**Related Functions**

Syntax 1&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the category axis in 2-D charts

Syntax 2&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the value axis in 2-D charts

Syntax 4&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the series axis in 3-D charts

Syntax 5&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the value axis in 3-D charts

Return to [top](#Q)

# SCALE Syntax 4

Equivalent to clicking the Selected Axes command on the Format menu when
a chart's value (y) axis is selected, and then clicking the Scale tab.
There are five syntax forms of this function. Syntax 4 of SCALE applies
if the selected axis is a series (y) axis on a 3-D chart. Use this
syntax of SCALE to change the position, formatting, and scaling of the
series axis.

**Syntax 4**

Series (y) axis, 3-D chart

**SCALE**(series\_labels, series\_marks, reverse)

**SCALE**?(series\_labels, series\_marks, reverse)

Series\_labels&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the
Number Of Series Between Tick Labels text box. The default is 1.
Series\_labels can also be a logical value. If TRUE, and automatic
setting will be used. If FALSE, or omitted, the number will be used.

Series\_marks&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the
Number Of Series Between Tick Marks text box. The default is 1.
Series\_marks can also be a logical value. If TRUE, and automatic
setting will be used. If FALSE, or omitted, the number will be used.

Reverse&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to
the Series In Reverse Order check box on the Scale tab. If reverse is
TRUE, Microsoft Excel displays the series in reverse order; if FALSE or
omitted, Microsoft Excel displays the series normally.

**Related Functions**

Syntax 1&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the category axis in 2-D charts

Syntax 2&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the value axis in 2-D charts

Syntax 3&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the category axis in 3-D charts

Syntax 5&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the value axis in 3-D charts

Return to [top](#Q)

# SCALE Syntax 5

Equivalent to clicking the Selected Axes command on the Format menu when
a chart's value (z) axis is selected, and then clicking the Scale tab.
There are five syntax forms of this function. Syntax 5 of SCALE applies
if the selected axis is a value (z) axis on a 3-D chart. Use this syntax
of SCALE to change the position, formatting, and scaling of the value
axis.

**Syntax 5**

**SCALE**(min\_num, max\_num, major, minor, cross, logarithmic, reverse,
min)

**SCALE**?(min\_num, max\_num, major, minor, cross, logarithmic,
reverse, min)

The first five arguments correspond to the five range variables in the
Format Axis dialog box, as shown in the following list. Each argument
can be either the logical value TRUE or a number.

  - > If TRUE or omitted, the Auto check box is selected.

  - > If a number, that number is used.

Min\_num&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Minimum check box and
is the minimum value for the value axis.

Max\_num&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Maximum check box and
is the maximum value for the value axis.

Major&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Major Unit check box and
is the major unit of measure.

Minor&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Minor Unit check box and
is the minor unit of measure.

Cross&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Floor (XY Plane) Crosses
At check box.

The last three arguments are logical values corresponding to check boxes
on the Scale tab. If an argument is TRUE, Microsoft Excel selects the
check box; if FALSE, Microsoft Excel clears the check box.

Logarithmic&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Logarithmic Scale
check box.

Reverse&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Values In Reverse
Order check box.

Min&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Floor (XY Plane) Crosses
At Minimum Value check box.

**Related Functions**

Syntax 1&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the category axis in 2-D charts

Syntax 2&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the value axis in 2-D charts

Syntax 3&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the category axis in 3-D charts

Syntax 4&nbsp;&nbsp;&nbsp;Changes the position, formatting, and scaling
of the series axis in 3-D charts

Return to [top](#Q)

# SCENARIO.ADD

Equivalent to clicking the Scenarios command on the Tools menu and then
clicking the Add button. Defines the specified values as a scenario. A
scenario is a set of values to be used as input for a model on your
worksheet.

**Syntax**

**SCENARIO.ADD**(**scen\_name**, value\_array, changing\_ref,
scen\_comment, locked, hidden)

Scen\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the scenario you want
to define.

Value\_array&nbsp;&nbsp;&nbsp;&nbsp;is a horizontal array of values you
want to use as input for the model on your worksheet.

  - > Any entry that would be valid for a cell in your model can be a
    > value in value\_array.

  - > The values must be arranged in the same order as the model's
    > changing cells. The changing cells are listed in the Changing
    > Cells box in the Scenario Manager dialog box.

  - > If value\_array is omitted, it is assumed to contain the current
    > values of the changing cells.

> &nbsp;

Changing\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to cells you want to
define as changing cells for a scenario.

  - > If omitted, uses the changing cells for the last scenario defined
    > for the sheet.

  - > If changing\_ref contains nonadjacent references, you must
    > separate the reference areas by commas (or other list separator).
    > If you are using A1-style references, then you must enclose
    > reference in an extra set of parentheses.

> &nbsp;

Scen\_comment&nbsp;&nbsp;&nbsp;&nbsp;is text specifying a descriptive
comment for the scenario defined by scen\_name.

Locked&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to the
Prevent Changes check box in the Add or Edit Scenario dialogs boxes. If
TRUE or omitted , prevents users from changing values in a scenario. If
FALSE, users are allowed to make changes to the scenario. The locking
will not become enabled until the sheet is protected with the Protect
Sheet command from the Protection submenu on the Tools menu.

Hidden&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to the
Hide check box in the Add or Edit Scenario dialog boxes. If TRUE, the
scenario will be hidden from view from the users and will not appear in
the Scenario Manager dialog box. If FALSE or omitted, the scenario will
remain unhidden. The scenario will not become hidden until the sheet is
protected with the Protect Sheet command from the Protection submenu on
the Tools menu.

**Related Functions**

REPORT.DEFINE&nbsp;&nbsp;&nbsp;Creates a report

SCENARIO.GET&nbsp;&nbsp;&nbsp;Returns the specified information about
the scenarios defined on your worksheet

Return to [top](#Q)

# SCENARIO.CELLS

Equivalent to clicking the Scenarios command on the Tools menu and then
editing the Changing Cells box. Defines the changing cells for a model
on your worksheet. Changing cells are the cells into which values will
be entered when you display a scenario. If you have only one set of
changing cells on your sheet, SCENARIO.CELLS will change the changing
cells for all scenarios. If your sheet has scenarios defined with
multiple sets of changing cell, this function returns an error and the
macro is halted. This function is provided for compatibility with
Microsoft Excel version 4.0. Use SCENARIO.EDIT with the changing\_ref
argument instead of SCENARIO.CELLS if you want to change the changing
cells of a scenario.

**Syntax**

**SCENARIO.CELLS**(**changing\_ref**)

**SCENARIO.CELLS**? (changing\_ref)

Changing\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cells you
want to define as changing cells for the model. If changing\_ref
contains nonadjacent references, you must separate the reference areas
by commas and enclose changing\_ref in an extra set of parentheses.

**Related Function**

SCENARIO.EDIT&nbsp;&nbsp;&nbsp;Equivalent to clicking the Scenarios
command on the Tools menu and then clicking the Edit button

Return to [top](#Q)

# SCENARIO.DELETE

Equivalent to clicking the Scenarios command on the Tools menu, clicking
a scenario, and then clicking the Delete button. Deletes the specified
scenario.

**Syntax**

**SCENARIO.DELETE**(**scen\_name**)

Scen\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the scenario you want
to delete.

**Related Functions**

SCENARIO.GET&nbsp;&nbsp;&nbsp;Returns the specified information about
the scenarios defined on your worksheet

SCENARIO.ADD&nbsp;&nbsp;&nbsp;Equivalent to clicking the Scenario
Manager command on the Tools menu and then clicking the Add button

SCENARIO.EDIT&nbsp;&nbsp;&nbsp;Equivalent to clicking the Scenario
Manager command on the Tools menu and then clicking the Edit button

Return to [top](#Q)

# SCENARIO.EDIT

Equivalent to clicking the Scenarios command from the Tools menu and
then clicking the Edit button.

**Syntax**

**SCENARIO.EDIT**(**scen\_name**, new\_scenname, value\_array,
changing\_ref, scen\_comment, locked, hidden)

**SCENARIO.EDIT**?(scen\_name, new\_scenname, value\_array,
changing\_ref, scen\_comment, locked, hidden)

Scen\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the scenario that you
want to edit.

New\_scenname&nbsp;&nbsp;&nbsp;&nbsp;is the new name you want to give to
the scenario.

Value\_array&nbsp;&nbsp;&nbsp;&nbsp;is a horizontal array of values that
you want to use for the scenario.

  - > If value\_array is omitted but changing\_ref is specified,
    > Scenario Manager uses the values in changing\_ref as value\_array.

  - > Value\_array must match the dimensions of changing\_ref for the
    > scenario being edit.

Changing\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to cells you want to
define as changing cells for a scenario.

Scen\_comment&nbsp;&nbsp;&nbsp;&nbsp;is text specifying a descriptive
comment for the scenario you want to edit.

Locked&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to the
Prevent Changes check box in the Add or Edit Scenario dialogs boxes. If
TRUE or omitted , prevents users from changing values in a scenario. If
FALSE, users are allowed to make changes to the scenario. The locking
will not become enabled until the sheet is protected with the Protect
Sheet command from the Protection submenu on the Tools menu.

Hidden&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to the
Hide check box in the Add or Edit Scenario dialog boxes. If TRUE, the
scenario will be hidden from view from the users. If FALSE or omitted,
the scenario will remain unhidden. The scenario will not become hidden
until the sheet is hidden with the Hide command from the Window menu.

**Related Functions**

SCENARIO.GET&nbsp;&nbsp;&nbsp;Returns the specified information about
the scenarios defined on your worksheet

SCENARIO.ADD&nbsp;&nbsp;&nbsp;Equivalent to clicking the Scenario
Manager command on the Tools menu and then clicking the Add button

SCENARIO.DELETE&nbsp;&nbsp;&nbsp;Equivalent to clicking the Scenario
Manager command on the Tools menu and then selecting a scenario and
clicking the Delete button

Return to [top](#Q)

# SCENARIO.GET

Returns the specified information about the scenarios defined on your
worksheet.

**Syntax**

**SCENARIO.GET**(**type\_num**, scen\_name)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 8 specifying the
type of information you want.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Information returned</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>A horizontal array of all scenario names in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A reference to the set of changing cells of scen_name (specified in the Changing Cells box of the Scenario Manager dialog box). If scen_name is omitted, the first scenario is used.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>A reference to the result cells (specified in the Result Cells box in the Scenario Summary dialog box)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>An array of scenario values for the scenario scen_name . Each scenario is in a separate row. If scen_name is omitted, the first scenario is used.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Comment, as text, for the scenario</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Returns TRUE if the specified scenario is locked to prevent changes; FALSE, if unlocked. Scen_name is required.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Returns TRUE if the specified scenario is hidden; FALSE, if visible to the user. Scen_name is required.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Returns the user name of the person who last modified the scenario by either adding or editing a scenario. Scen_name is required.</p>
</blockquote></td>
</tr>
</tbody>
</table>

Scen\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the scenario that you
want information about. Ignored if type\_num equals 1 or 3.

**Remarks**

In the returned array of scenario values, the number of rows is the
number of scenarios, and the number of columns is the number of changing
cells.

Return to [top](#Q)

# SCENARIO.MERGE

Equivalent to choosing the Scenarios command from the Tools menu and
then selecting Merge. This function merges scenarios from other sheets
onto the active sheet. A scenario is a set of values to be used as input
for a model on your worksheet.

**Syntax**

**SCENARIO.MERGE**(source\_file)

**SCENARIO.MERGE**?(source\_file)

Source\_file&nbsp;&nbsp;&nbsp;&nbsp;is the name of the book and sheet
from which you want to merge scenarios onto the active sheet.

**Related Function**

SCENARIO.GET&nbsp;&nbsp;&nbsp;Returns the specified information about
the scenarios defined on your worksheet

Return to [top](#Q)

# SCENARIO.SHOW

Equivalent to clicking the Scenarios command on the Tools menu and then
selecting a scenario and clicking the Show button. Recalculates a model
using the specified scenario and displays the result.

**Syntax**

**SCENARIO.SHOW**(**scen\_name**)

Scen\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of the previously defined
scenario whose values you want to switch to.

Return to [top](#Q)

# SCENARIO.SHOW.NEXT

Equivalent to clicking the Scenarios command on the Tools menu,
selecting the next scenario from the Scenarios list, and clicking the
Show button. Recalculates a model using the next scenario and displays
the result.

**Syntax**

**SCENARIO.SHOW.NEXT**( )

**Remarks**

After displaying the last scenario, running SCENARIO.SHOW.NEXT again
displays the first scenario.

Return to [top](#Q)

# SCENARIO.SUMMARY

Equivalent to clicking the Scenarios command on the Tools menu and then
clicking the Summary button. Generates a table summarizing the results
of all the scenarios for the model on your worksheet.

**Syntax**

**SCENARIO.SUMMARY**(result\_ref, report\_type)

**SCENARIO.SUMMARY**?(result\_ref, report\_type)

Result\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the result cells
you want to include in the summary report. Normally, result\_ref refers
to one or more cells containing the formulas that depend on the changing
cell values for your model&mdash;that is, the cells that show the
results of a particular scenario.

  - > If result\_ref is omitted, no result cells are included in the
    > report.

  - > If result\_ref contains nonadjacent references, you must separate
    > the reference areas by commas and enclose result\_ref in an extra
    > set of parentheses.

Report\_type&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the type of
report desired.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Report_type</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Type of Report</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1 or omitted</p>
</blockquote></td>
<td><blockquote>
<p>A scenario summary report (Microsoft Excel version 4.0)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A scenario PivotTable report. Requires result_ref.</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Remarks**

  - > SCENARIO.SUMMARY generates a summary table of the changing cell
    > and result cell values for each scenario.

  - > The table is generated on a new sheet in the current workbook. The
    > sheet becomes active after SCENARIO.SUMMARY runs.

Return to [top](#Q)

# SCROLLBAR.PROPERTIES

Sets the properties of the scroll bar and spinner button on a worksheet
or dialog sheet.

**Syntax**

**SCROLLBAR.PROPERTIES**(value, min, max, inc, page, link, 3d\_shading)

**SCROLLBAR.PROPERTIES**?(value, min, max, inc, page, link, 3d\_shading)

Value&nbsp;&nbsp;&nbsp;&nbsp;is the value of the control, and can range
from min to max, inclusive. It designates where the scroll bar button is
positioned along the scroll bar.

Min&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the minimum value that
the scroll bar can have. This number ranges from 0 to 30,000, but cannot
be greater than the maximum value given in max.

Max&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the maximum value that
the scroll bar can have. This number ranges from 0 to 30,000.

Inc&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the increment that the
value is adjusted by when the scrollbar arrow is clicked.

Page&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the increment that
the value is adjusted by when the page scroll region of a scroll bar is
clicked.

Link&nbsp;&nbsp;&nbsp;&nbsp;is the cell on the macro sheet to which the
scroll bar value is linked. Whenever the scroll bar control is changed,
the value of the control is entered into the cell. Similarly, whenever
the value in the cell is changed, the setting for the scroll bar is also
changed. To clear the link, set this value to an empty string.

3d\_shading&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether the scroll bar or spinner button appears as 3-D. If TRUE, the
scroll bar or spinner button will appear as 3-D. If FALSE or omitted,
the scroll bar or spinner button will not be 3-D. The argument is
available for only worksheets.

**Related Functions**

PUSHBUTTON.PROPERTIES&nbsp;&nbsp;&nbsp;Sets the properties of the push
button control

EDITBOX.PROPERTIES&nbsp;&nbsp;&nbsp;Sets the properties of an edit box
on a worksheet or dialog sheet

Return to [top](#Q)

# SELECT

Equivalent to selecting cells or changing the active cell. There are
three syntax forms of SELECT. Use syntax 1 to select a cell on a
worksheet or macro sheet; use one of the other syntax forms to select
worksheet or macro sheet objects or chart items.

Syntax 1&nbsp;&nbsp;&nbsp;Selects cells

Syntax 2&nbsp;&nbsp;&nbsp;Selects objects on worksheets

Syntax 3&nbsp;&nbsp;&nbsp;Selects chart objects

Return to [top](#Q)

# SELECT Syntax 1

Equivalent to selecting cells or changing the active cell. There are
three syntax forms of SELECT. Use syntax 1 to select a cell on a
worksheet or macro sheet; use one of the other syntax forms to select
worksheet or macro sheet objects or chart items.

**Syntax**

**SELECT**(selection, active\_cell)

Selection&nbsp;&nbsp;&nbsp;&nbsp;is the cell or range of cells you want
to select. Selection can be a reference to the active worksheet, such as
\!$A$1:$A$3 or \!Sales, or an R1C1-style reference to a cell or range
relative to the active cell in the current selection, such as
"R\[-1\]C\[-1\]:R\[1\]C\[1\]". The reference must be in text form. If
selection is omitted, the current selection is used.

Active\_cell&nbsp;&nbsp;&nbsp;&nbsp;is the cell in selection you want to
make the active cell. Active\_cell can be a reference to a single cell
on the active worksheet, such as \!$A$1, or an R1C1-style reference
relative to the active cell, such as "R\[-1\]C\[-1\]". The reference
must be in text form. If active\_cell is omitted, SELECT makes the cell
in the upper-left corner of selection the active cell.

**Remarks**

  - > Active\_cell must be within selection. If it is not, an error
    > message is displayed and SELECT returns the \#VALUE\! error value.

  - > If you are recording a macro using relative references, Microsoft
    > Excel records the action using R1C1-style relative references in
    > the form of text.

  - > If you are recording using absolute references, Microsoft Excel
    > records the action using R1C1-style absolute references in the
    > form of text.

  - > You cannot give an external reference to a specific sheet as the
    > selection argument. The sheet on which you want to make a
    > selection must be active when you use SELECT. Use FORMULA.GOTO to
    > make a selection on another sheet in the same workbook or in
    > another workbook.

> &nbsp;

**Tip**&nbsp;&nbsp;&nbsp;You can enter data in a cell without selecting
the cell by using the reference arguments to the CUT, COPY, or FORMULA
functions.

**Examples**

The following macro formula selects cells C3:E5 on the active worksheet
and makes C5 the active cell:

SELECT(\!$C$3:$E$5, \!$C$5)

If the active cell is C3, the following macro formula selects cells
E5:G7 and makes cell F6 the active cell in the selection:

SELECT("R\[2\]C\[2\]:R\[4\]C\[4\]", "R\[1\]C\[1\]")

You can also make multiple nonadjacent selections with SELECT. The
following macro formula selects a number of nonadjacent ranges:

SELECT("R1C1, R3C2:R4C3, R8C4:R10C5")

The following sequence of macro formulas moves the active cell right,
left, down, and up within the selection, just as TAB, SHIFT+TAB, ENTER,
and SHIFT+ENTER do:

SELECT(, "RC\[1\]")

SELECT(, "RC\[-1\]")

SELECT(, "R\[1\]C")

SELECT(, "R\[-1\]C")

Use SELECT with the OFFSET function to select a new range a specified
distance away from the current range. For example, the following macro
formula selects a range that is the same size as the current range, one
column over:

SELECT(OFFSET(SELECTION(), 0, 1))

**Related Functions**

ACTIVE.CELL&nbsp;&nbsp;&nbsp;Returns the reference of the active cell

SELECT.SPECIAL&nbsp;&nbsp;&nbsp;Selects a group of cells belonging to a
category

SELECTION&nbsp;&nbsp;&nbsp;Returns the reference of the selection

SELECT Syntax 2&nbsp;&nbsp;&nbsp;Selects objects on worksheets

SELECT Syntax 3&nbsp;&nbsp;&nbsp;Selects chart objects

Return to [top](#Q)

# SELECT Syntax 2

Equivalent to selecting objects on a chart, worksheet, or macro sheet.
There are three syntax forms of SELECT. Use syntax 2 to select an object
on which to perform an action; use one of the other syntax forms to
select cells on a worksheet or macro sheet or items on a chart.

**Syntax**

**SELECT**(object\_id\_text, replace)

Object\_id\_text&nbsp;&nbsp;&nbsp;&nbsp;is text that identifies the
object to select. Object\_id\_text can be the name of more than one
object. To give the name of more than one object, use the following
format:

SELECT("Oval 3, Arc 2, Line 4")

The last item in the object\_id\_text list will be the active object.
The active object is important when moving and sizing a group of
objects. A multiple selection of objects is moved and sized relative to
the upper-left corner of the active object.

Replace&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies whether
previously selected objects are included in the selection. If replace is
TRUE or omitted, Microsoft Excel only selects the objects specified by
object\_id\_text; if FALSE, it includes any objects that were previously
selected. For example, if a button is selected and a SELECT formula
selects an arc and an oval, TRUE leaves only the arc and oval selected,
and FALSE includes the button with the arc and oval.

**Remarks**

Objects can be identified by their object type and number as described
in CREATE.OBJECT, or by the unique number that specifies the order of
their creation. For example, if the third object you create is an oval,
you could use either "oval 3" or "3" as object\_id\_text.

**Examples**

The following macro formulas each select a number of objects and specify
Arc 2 as the active object:

SELECT("Oval 3, Arc 1, Line 4, Arc 2")

SELECT("3, 1, 4, 2")

**Related Functions**

FORMAT.MOVE&nbsp;&nbsp;&nbsp;Moves the selected object

FORMAT.SIZE&nbsp;&nbsp;&nbsp;Changes the size of the selected objects

GET.OBJECT&nbsp;&nbsp;&nbsp;Returns information about an object

SELECTION&nbsp;&nbsp;&nbsp;Returns the reference of the selection

SELECT Syntax 1&nbsp;&nbsp;&nbsp;Selects cells

SELECT Syntax 3&nbsp;&nbsp;&nbsp;Selects chart objects

Return to [top](#Q)

# SELECT Syntax 3

Selects a chart object as specified by the selection code item\_text.
There are three syntax forms of SELECT. Use syntax 3 to select a chart
item to which you want to apply formatting; use one of the other syntax
forms to select cells or objects on a worksheet or macro sheet.

**Syntax**

**SELECT**(item\_text, single\_point)

Item\_text&nbsp;&nbsp;&nbsp;&nbsp;is a selection code from the following
table which specifies which chart object to select.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>To select</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Item_text</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Entire chart</p>
</blockquote></td>
<td><blockquote>
<p>"Chart"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Plot area</p>
</blockquote></td>
<td><blockquote>
<p>"Plot"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Legend</p>
</blockquote></td>
<td><blockquote>
<p>"Legend"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Primary chart value axis</p>
</blockquote></td>
<td><blockquote>
<p>"Axis 1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Primary chart category axis</p>
</blockquote></td>
<td><blockquote>
<p>"Axis 2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Secondary chart value axis or 3-D series axis</p>
</blockquote></td>
<td><blockquote>
<p>"Axis 3"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Secondary chart category axis</p>
</blockquote></td>
<td><blockquote>
<p>"Axis 4"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Chart title</p>
</blockquote></td>
<td><blockquote>
<p>"Title"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Label for the primary chart value axis</p>
</blockquote></td>
<td><blockquote>
<p>"Text Axis 1"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Label for the primary chart category axis</p>
</blockquote></td>
<td><blockquote>
<p>"Text Axis 2"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Label for the primary chart series axis</p>
</blockquote></td>
<td><blockquote>
<p>"Text Axis 3"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>nth floating text item</p>
</blockquote></td>
<td><blockquote>
<p>"Text n"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>nth arrow</p>
</blockquote></td>
<td><blockquote>
<p>"Arrow n"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Major gridlines of value axis</p>
</blockquote></td>
<td><blockquote>
<p>"Gridline 1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Minor gridlines of value axis</p>
</blockquote></td>
<td><blockquote>
<p>"Gridline 2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Major gridlines of category axis</p>
</blockquote></td>
<td><blockquote>
<p>"Gridline 3"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Minor gridlines of category axis</p>
</blockquote></td>
<td><blockquote>
<p>"Gridline 4"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Major gridlines of series axis</p>
</blockquote></td>
<td><blockquote>
<p>"Gridline 5"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Minor gridlines of series axis</p>
</blockquote></td>
<td><blockquote>
<p>"Gridline 6"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Primary chart droplines</p>
</blockquote></td>
<td><blockquote>
<p>"Dropline 1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Secondary chart droplines</p>
</blockquote></td>
<td><blockquote>
<p>"Dropline 2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Primary chart hi-lo lines</p>
</blockquote></td>
<td><blockquote>
<p>"Hiloline 1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Secondary chart hi-lo lines</p>
</blockquote></td>
<td><blockquote>
<p>"Hiloline 2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Primary chart up bar</p>
</blockquote></td>
<td><blockquote>
<p>"UpBar1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Secondary chart up bar</p>
</blockquote></td>
<td><blockquote>
<p>"UpBar2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Primary chart down bar</p>
</blockquote></td>
<td><blockquote>
<p>"DownBar1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Secondary chart down bar</p>
</blockquote></td>
<td><blockquote>
<p>"DownBar2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Primary chart series line</p>
</blockquote></td>
<td><blockquote>
<p>"Seriesline1"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Secondary chart series line</p>
</blockquote></td>
<td><blockquote>
<p>"Seriesline2"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Entire series</p>
</blockquote></td>
<td><blockquote>
<p>"Sn"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Data associated with point m in series n if single_point is TRUE</p>
</blockquote></td>
<td><blockquote>
<p>"SnPm"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Text attached to point m of series n</p>
</blockquote></td>
<td><blockquote>
<p>"Text SnPm"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Series title text of series n of an area chart</p>
</blockquote></td>
<td><blockquote>
<p>"Text Sn"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Base of a 3-D chart</p>
</blockquote></td>
<td><blockquote>
<p>"Floor"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Back of a 3-D chart</p>
</blockquote></td>
<td><blockquote>
<p>"Walls"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Corners of a 3-D chart</p>
</blockquote></td>
<td><blockquote>
<p>"Corners"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Trend line</p>
</blockquote></td>
<td><blockquote>
<p>"SnTm"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Error bars</p>
</blockquote></td>
<td><blockquote>
<p>"SnEm"</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Legend Marker</p>
</blockquote></td>
<td><blockquote>
<p>"Legend Marker n"</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Legend Entry</p>
</blockquote></td>
<td><blockquote>
<p>"Legend Entry n"</p>
</blockquote></td>
</tr>
</tbody>
</table>

For trend lines and error bars, the value m can be X or Y, depending on
which point you want to select. If m is blank, selects both.

Single\_point&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that determines
whether to select a single point. Single\_point is available only when
item\_text is "SnPm".

  - > If single\_point is TRUE, Microsoft Excel selects a single point.

  - > If single\_point is FALSE or omitted, Microsoft Excel selects a
    > single point if there is only one series in the chart or selects
    > the entire series if there is more than one series in the chart.

  - > If you specify single\_point when item\_text is any value other
    > than "SnPm", SELECT returns an error value.

> &nbsp;

**Examples**

SELECT("Chart") selects the entire chart.

SELECT("Dropline 2") selects the droplines of an overlay chart.

SELECT("S1P3", TRUE) selects the third point in the first series.

SELECT("Text S1") selects the series title text of the first series in
an area chart.

**Related Functions**

SELECTION&nbsp;&nbsp;&nbsp;Returns the reference of the selection

SELECT Syntax 1&nbsp;&nbsp;&nbsp;Selects cells

SELECT Syntax 2&nbsp;&nbsp;&nbsp;Selects objects on worksheets

Return to [top](#Q)

# SELECT.ALL

Equivalent to selecting all the sheets in a workbook.

**Syntax**

**SELECT.ALL**( )

Return to [top](#Q)

# SELECT.CHART

Equivalent to the Select Chart command on the Chart menu in Microsoft
Excel version 4.0. This function is equivalent to using the third form
of SELECT with "Chart" as the item\_text argument.

**Syntax**

**SELECT.CHART**( )

**Remarks**

This function is included for compatibility with macros written with
Microsoft Excel for the Macintosh version 1.5 or earlier.

**Related Function**

SELECT&nbsp;&nbsp;&nbsp;Selects a chart object

Return to [top](#Q)

# SELECT.END

Selects the cell at the edge of the range or the first cell of the next
range in the direction specified. Equivalent to pressing CTRL+ARROW in
Microsoft Excel for Windows or COMMAND+ARROW in Microsoft Excel for the
Macintosh.

**Syntax**

**SELECT.END**(**direction\_num**)

Direction\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 indicating
the direction in which to move.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Direction_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Direction</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Left (equivalent to CTRL+LEFT ARROW or COMMAND+LEFT ARROW)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Right (equivalent to CTRL+RIGHT ARROW or COMMAND+RIGHT ARROW)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Up (equivalent to CTRL+UP ARROW or COMMAND+UP ARROW)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Down (equivalent to CTRL+DOWN ARROW or COMMAND+DOWN ARROW)</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Related Function**

SELECT.LAST.CELL&nbsp;&nbsp;&nbsp;Selects the last cell on a worksheet
or macro sheet that contains a formula, value, or format or that is
referred to in a formula or name

Return to [top](#Q)

# SELECTION

Returns the reference or object identifier of the selection as an
external reference. Use SELECTION to return information about the
current selection for use in other macro formulas.

**Syntax**

**SELECTION**( )

If a cell or range of cells is selected, Microsoft Excel returns the
corresponding external reference. If an object is selected, Microsoft
Excel returns the object identifier listed in the following table.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Item selected</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Identifier returned</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Imported graphic</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Linked graphic</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Chart picture</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Linked chart</p>
</blockquote></td>
<td><blockquote>
<p>Chart n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Range</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Linked range</p>
</blockquote></td>
<td><blockquote>
<p>Picture n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Text box</p>
</blockquote></td>
<td><blockquote>
<p>Text n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Button</p>
</blockquote></td>
<td><blockquote>
<p>Button n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Rectangle</p>
</blockquote></td>
<td><blockquote>
<p>Rectangle n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Oval</p>
</blockquote></td>
<td><blockquote>
<p>Oval n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Line</p>
</blockquote></td>
<td><blockquote>
<p>Line n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Arc</p>
</blockquote></td>
<td><blockquote>
<p>Arc n</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Group</p>
</blockquote></td>
<td><blockquote>
<p>Group n</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Freehand drawing or polygon</p>
</blockquote></td>
<td><blockquote>
<p>Drawing n</p>
</blockquote></td>
</tr>
</tbody>
</table>

SELECTION also returns the identifiers of chart items. The identifiers
returned are the same as the identifiers you specify when you use the
SELECT function. For a list of these identifiers, see the description of
item\_text in SELECT.

If you select cells and use the value returned by SELECTION in a
function or operation, you usually get the value contained in the
selection instead of its reference. References are automatically
converted to the contents of the reference. If you want to work with the
actual reference, use SET.NAME to assign a name to it, even if the
reference refers to objects. See the last example following. You can
also use the REFTEXT function to convert the reference to text, which
you can then store or manipulate.

**Remarks**

  - > If an object is selected, SELECTION returns the identifier of the
    > object. If multiple objects are selected, it returns the
    > identifiers of all the selected objects, as a string separated by
    > commas.

  - > If more than 1024 characters would be returned, SELECTION returns
    > the \#VALUE\! error value.

> &nbsp;

**Examples**

If the sheet in the active window is named SHEET1 in the workbook BOOK1,
and if A1:A3 is the selection, then:

SELECTION() equals \[BOOK1\]SHEET1\!A1:A3

The following macro formula moves the current selection one row down:

SELECT(OFFSET(SELECTION(), 1, 0))

The above formula is particularly useful for moving incrementally
through a database to add or modify records.

The following macro formula defines the name "EntryRange" on the active
sheet to refer to one row below the current selection on the active
sheet:

DEFINE.NAME("EntryRange", OFFSET(SELECTION(), 1, 0))

The following macro formula defines the name "Objects" on your macro
sheet to refer to the object names in the current multiple selection:

SET.NAME("Objects", SELECTION())

**Related Functions**

ACTIVE.CELL&nbsp;&nbsp;&nbsp;Returns the reference of the active cell

SELECT&nbsp;&nbsp;&nbsp;Selects a cell, graphic object, or chart

Return to [top](#Q)

# SELECT.LAST.CELL

Equivalent to choosing the Special button from the Go To dialog box and
selecting the Last Cell option. The Go To dialog box appears when you
choose the Go To command from the Edit menu. Selects the cell at the
intersection of the last row and column that contains a formula, value,
or format, or that is referred to in a formula or name.

**Syntax**

**SELECT.LAST.CELL**( )

**Related Function**

SELECT.END&nbsp;&nbsp;&nbsp;Selects the last cell in a range

Return to [top](#Q)

# SELECT.LIST.ITEM

Selects an item in a list box or in a group box.

**Syntax**

**SELECT.LIST.ITEM**(**index\_num**, selected\_logical)

Index\_num&nbsp;&nbsp;&nbsp;&nbsp;is the index number of the item to
select. Using zero will deselect all items. Adding 1 to the number of
items in the list will select all the items specified.

Selected\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies the
selection mode of the list box. Zero is single selection. 1 is simple
multi-select. 2 is extended multi-select.

**Related Functions**

ADD.LIST.ITEM&nbsp;&nbsp;&nbsp;Adds an item in a list box or drop-down
control on a worksheet or dialog sheet control

REMOVE.LIST.ITEM&nbsp;&nbsp;&nbsp;Removes an item in a list box or
drop-down box

Return to [top](#Q)

# SELECT.PLOT.AREA

Equivalent to clicking the Select Plot Area command on the Chart menu in
Microsoft Excel version 4.0. Selects the plot area of the active chart.

**Syntax**

**SELECT.PLOT.AREA**( )

**Remarks**

SELECT.PLOT.AREA is included only for compatibility with previous
versions of Microsoft Excel for the Macintosh. SELECT.PLOT.AREA is the
same as the SELECT("Plot") function.

**Related Function**

SELECT&nbsp;&nbsp;&nbsp;Selects a cell, graphic object, or chart

Return to [top](#Q)

# SELECT.SPECIAL

Equivalent to clicking the Go To command on the Edit menu and then
selecting the Special button. Use SELECT.SPECIAL to select groups of
similar cells in one of a variety of categories.

**Syntax**

**SELECT.SPECIAL**(**type\_num**, value\_type, levels)

**SELECT.SPECIAL**?(type\_num, value\_type, levels)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 13 corresponding
to options in the Go To Special dialog box and describes what to select.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Description</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Notes/comments</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Constants</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Formulas</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Blanks</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Current region</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Current array</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Row differences</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Column differences</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>Precedents</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>Dependents</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>Last cell</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>Visible cells only (outlining)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>All objects</p>
</blockquote></td>
</tr>
</tbody>
</table>

Value\_type&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying which types of
constants or formulas you want to select. Value\_type is available only
when type\_num is 2 or 3.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Value_type</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Selects</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Numbers</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Text</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Logical values</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>Error values</p>
</blockquote></td>
</tr>
</tbody>
</table>

These values can be added to select more than one type. The default for
value\_type is 23, which select all value types.

Levels&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying how precedents and
dependents are selected. Levels is available only when type\_num is 9 or
10. The default is 1.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Levels</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Selects</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Direct only</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>All levels</p>
</blockquote></td>
</tr>
</tbody>
</table>

Return to [top](#Q)

# SEND.KEYS

Sends keystrokes to the active application just as if they were typed at
the keyboard. Use SEND.KEYS to send keystrokes that perform actions and
execute commands to applications you are running with Microsoft Excel's
other dynamic data exchange (DDE) functions.

**Syntax**

**SEND.KEYS**(**key\_text**, wait\_logical)

**Note&nbsp;&nbsp;&nbsp;**This function is available only in Microsoft
Excel for Windows.

Key\_text&nbsp;&nbsp;&nbsp;&nbsp;is the key or key combination you want
to send to another application. The format for key\_text is described in
the ON.KEY function.

Wait\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that determines
whether the macro continues before the actions caused by key\_text are
carried out.

  - > If wait\_logical is TRUE, Microsoft Excel waits for the keys to be
    > processed before returning control to the macro.

  - > If wait\_logical is FALSE or omitted, the macro continues running
    > without waiting for the keys to be processed.

> &nbsp;

**Remarks**

If Microsoft Excel is the active application, wait\_logical is assumed
to be FALSE, even if you enter wait\_logical as TRUE. This is because if
wait\_logical is TRUE, Microsoft Excel waits for the keys to be
processed in the other application before returning control to the
macro. Microsoft Excel doesn't process keys while a macro is running.

**Example**

The following macro uses the Calculator application in Microsoft Excel
for Windows to multiply some numbers, and then cuts the result and
pastes it into Microsoft Excel.

\=EXEC("CALC.EXE", 1)

\=SEND.KEYS("10\*30", TRUE)

\=SEND.KEYS("\~", TRUE)

\=SEND.KEYS("%ec", TRUE)

\=APP.ACTIVATE(, FALSE)

\=SELECT(\!B1)

\=PASTE()

\=RETURN()

**Related Functions**

APP.ACTIVATE&nbsp;&nbsp;&nbsp;Switches to an application

EXECUTE&nbsp;&nbsp;&nbsp;Carries out a command in another application

ON.KEY&nbsp;&nbsp;&nbsp;Runs a macro when a specified key is pressed

Return to [top](#Q)

# SEND.MAIL

Equivalent to clicking the Send Mail command on the File menu. Sends the
active workbook using email.

**Syntax**

**SEND.MAIL**(**recipients**, subject, return\_receipt)

**SEND.MAIL**?(recipients, subject, return\_receipt)

**Important**&nbsp;&nbsp;&nbsp;To use SEND.MAIL in Microsoft Excel for
Windows, you must be using a mail client that supports the Messaging
Applications Programming Interface (MAPI) or Vendor-Independent
Messaging (VIM). To use SEND.MAIL in Microsoft Excel for the Macintosh,
you must be using Microsoft Mail version 2.0 or later.

Recipients&nbsp;&nbsp;&nbsp;&nbsp;is the name of the person to whom you
want to send the mail. The name should be given as text.

  - > To specify more than one name, give the list of names as an array.
    > For example, SEND.MAIL({"John", "Paul", "George", "Ringo"}) would
    > send the active workbook to the four names in the array. You can
    > also refer to a range on a sheet or macro sheet that contains a
    > list of names to whom you want the mail to be sent.

  - > To send mail to users on different Microsoft Mail for the
    > Macintosh servers, specify the server name along with the user
    > name. The following text, as the recipients argument, sends mail
    > to wandagr on server2, gregpr on the current server, and victorge
    > on server7:

> {"wandagr@server2", "gregpr", "victorge@server7"}
> 
> &nbsp;

Subject&nbsp;&nbsp;&nbsp;&nbsp;is a text string that specifies the
subject of the message. If subject is omitted, the name of the active
workbook is used as the subject.

Return\_receipt&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that
corresponds to the Return Receipt check box. If return\_receipt is TRUE,
Microsoft Excel selects the check box and sends a return receipt; if
FALSE or omitted, Microsoft Excel clears the check box.

**Related Function**

OPEN.MAIL&nbsp;&nbsp;&nbsp;Opens files sent via Microsoft Mail that
Microsoft Excel can open

Return to [top](#Q)

# SEND.TO.BACK

Sends the selected object or objects to the back. Use SEND.TO.BACK to
position selected objects behind other objects.

If the selection is not an object or a group of objects, SEND.TO.BACK
returns the \#VALUE\! error value and interrupts the macro.

**Syntax**

**SEND.TO.BACK**( )

**Related Function**

BRING.TO.FRONT&nbsp;&nbsp;&nbsp;Brings selected objects to the front

Return to [top](#Q)

# SERIES

Charts Only

Represents a data series in the active chart. SERIES is used only in
charts; you cannot enter it on a sheet or macro sheet. You normally
create or change data series by using the Chart Wizard or EDIT.SERIES
macro function, which is equivalent to the Edit Series command on the
Chart menu in Microsoft Excel version 4.0. However, you can edit a data
series manually by selecting it, switching to the formula bar, and
typing the changes.

**Syntax**

**SERIES**(name\_ref, categories, **values, plot\_order**)

Name\_ref&nbsp;&nbsp;&nbsp;&nbsp;is the name of the data series. It can
be an external reference to a single cell or a name defined as a single
cell. Name\_ref can also be text enclosed in quotation marks (for
example, "Projected Sales").

Categories&nbsp;&nbsp;&nbsp;&nbsp;is an external reference to the name
of the workbook and to the cells that contain one of the following sets
of data:

  - > Category labels for all charts except xy (scatter) charts

  - > X-coordinate data for xy (scatter) charts

> &nbsp;

Values&nbsp;&nbsp;&nbsp;&nbsp;is an external reference to the name of
the workbook and to the cells that contain values (or y-coordinate data
in scatter charts).

Plot\_order&nbsp;&nbsp;&nbsp;&nbsp;is an integer specifying whether the
series is plotted first, second, or third, and so on, in the chart. No
two series can have the same plot\_order.

**Remarks**

  - > Categories and values can be arrays or references to a multiple
    > selection, although they cannot be names that refer to a multiple
    > selection. If you specify a multiple selection for any of these
    > arguments, make sure you include the necessary sets of parentheses
    > so that Microsoft Excel does not treat the components of the
    > references as separate arguments.

  - > If either categories or values is a multiple selection, then all
    > areas in that selection must be either vertical (more rows than
    > columns) or horizontal (more columns than rows).

> &nbsp;

**Related Functions**

CHART.WIZARD&nbsp;&nbsp;&nbsp;Creates and formats a chart

EDIT.SERIES&nbsp;&nbsp;&nbsp;Creates or changes a chart series

Return to [top](#Q)

# SERIES.AXES

Equivalent to the Axis Tab in the Format Data Series dialog box. Changes
the axis on which a series is plotted. This function is for
compatibility with Microsoft Excel versions earlier than Microsoft Excel
97.

**Syntax**

**SERIES.AXES**(axis)

Axis&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying on which axis to plot
the data series: use 1 for primary axis, 2 for secondary axis.

Return to [top](#Q)

# SERIES.ORDER

Changes the order of series in a chart.

**Syntax**

**SERIES.ORDER**(chart\_num, old\_series\_num, new\_series\_num)

Chart\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number of the group containing
the series you want to change

Old\_series\_num&nbsp;&nbsp;&nbsp;&nbsp;is the current number of the
series in the group.

New\_series\_num&nbsp;&nbsp;&nbsp;&nbsp;is the new number you want for
the series in the group.

Return to [top](#Q)

# SERIES.X

Equivalent to the X Values tab in the Format Data Series dialog box.
Specifies the category labels (x values) for a data series. This
function is for compatibility with Microsoft Excel versions earlier than
Microsoft Excel 97.

**Syntax**

**SERIES.X**(x\_ref)  
X-ref&nbsp;&nbsp;&nbsp;&nbsp;is an external reference in the form of
text specifying the range containing the category labels (or x values
for a scatter (xy) chart) you want to use.

**Related Function**

SERIES.Y&nbsp;&nbsp;&nbsp;Specifies the name and values for a data
series

Return to [top](#Q)

# SERIES.Y

Equivalent to the Name and Values tab in the Format Data Series dialog
box. Specifies the name and values for a data series. This function is
for compatibility with Microsoft Excel versions earlier than Microsoft
Excel 97.

**Syntax**

**SERIES.Y**(name\_ref, y\_ref)

Name\_ref&nbsp;&nbsp;&nbsp;&nbsp;is text or an external reference in the
form of text specifying the name for the data series that appears in the
legend for the chart.

Y\_ref&nbsp;&nbsp;&nbsp;&nbsp;is an external reference in the form of
text specifying the range containing the values for the data series.

**Related Function**

SERIES.X&nbsp;&nbsp;&nbsp;Specifies the category labels (x values) for a
data series

Return to [top](#Q)

# SET.CONTROL.VALUE

Changes the value for the active control, such as a list box, drop-down
box, check box, option button, scroll bar, and spinner button.

**Syntax**

**SET.CONTROL.VALUE**(value)

Value&nbsp;&nbsp;&nbsp;&nbsp;is the value you want to change. The
control interprets this value as follows:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Control</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Value is</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>List box</p>
</blockquote></td>
<td><blockquote>
<p>The index of the selected item. If zero, then no item is selected.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Drop-down box</p>
</blockquote></td>
<td><blockquote>
<p>The index of the selected item. If zero, then no item is selected.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Check box</p>
</blockquote></td>
<td><blockquote>
<p>0 = Off<br />
1 = On<br />
2 = Mixed</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Option button</p>
</blockquote></td>
<td><blockquote>
<p>0= Off<br />
1 = On</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Scroll bar</p>
</blockquote></td>
<td><blockquote>
<p>The numeric value of the control, between the maximum and minimum values</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>Spinner button</p>
</blockquote></td>
<td><blockquote>
<p>The numeric value of the control, between the maximum and minimum values</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Related Functions**

ADD.LIST.ITEM&nbsp;&nbsp;&nbsp;Adds an item in a list box or drop-down
control on a worksheet or dialog sheet control

REMOVE.LIST.ITEM&nbsp;&nbsp;&nbsp;Removes an item in a list box or
drop-down box

SELECT.LIST.ITEM&nbsp;&nbsp;&nbsp;Selects an item in a list box or in a
group box

CHECKBOX.PROPERTIES&nbsp;&nbsp;&nbsp;Sets various properties of check
box and option box controls

SCROLLBAR.PROPERTIES&nbsp;&nbsp;&nbsp;Sets the properties of the scroll
bar and spinner controls

Return to [top](#Q)

# SET.CRITERIA

Equivalent to clicking the Set Criteria command on the Data menu in
Microsoft Excel version 4.0. Defines the name Criteria for the selected
range on a sheet or macro sheet.

**Syntax**

**SET.CRITERIA**( )

**Related Functions**

SET.DATABASE&nbsp;&nbsp;&nbsp;Equivalent to clicking the Set Database
command on the Data menu in Microsoft Excel version 4.0

SET.EXTRACT&nbsp;&nbsp;&nbsp;Equivalent to clicking the Set Extract
command on the Data menu in Microsoft Excel version 4.0

Return to [top](#Q)

# SET.DATABASE

Equivalent to clicking the Set Database command on the Data menu in
Microsoft Excel version 4.0. Defines the name Database for the selected
range on a sheet or macro sheet.

**Syntax**

**SET.DATABASE**( )

**Related Functions**

SET.CRITERIA&nbsp;&nbsp;&nbsp;Equivalent to clicking the Set Criteria
command on the Data menu in Microsoft Excel version 4.0

SET.EXTRACT&nbsp;&nbsp;&nbsp;Equivalent to clicking the Set Extract
command on the Data menu in Microsoft Excel version 4.0

Return to [top](#Q)

# SET.DIALOG.DEFAULT

Sets which button is automatically pressed (the default button) when the
user presses ENTER. While running, this default button is visually
recognized by its thick border. This function is used only with a dialog
sheet active.

**Syntax**

**SET.DIALOG.DEFAULT**(**object\_id\_text**)

Object\_id\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of the button
control to set as the default button, as in "Button 5".

**Related Function**

SET.DIALOG.FOCUS&nbsp;&nbsp;&nbsp;Sets the focus of a dialog box

Return to [top](#Q)

# SET.DIALOG.FOCUS

Sets the focus of a dialog box. This function is used only with a dialog
sheet active.

**Syntax**

**SET.DIALOG.FOCUS**(**object\_id\_text**)

Object\_id\_text&nbsp;&nbsp;&nbsp;&nbsp;the name of the control or
object as text to give the focus to, as in "Check box 4".

**Related Function**

SET.DIALOG.DEFAULT&nbsp;&nbsp;&nbsp;Sets which button is automatically
pressed (the default button) when the user presses ENTER

Return to [top](#Q)

# SET.EXTRACT

Equivalent to clicking the Set Extract command on the Data menu in
Microsoft Excel version 4.0. Defines the name Extract for the selected
range on the active sheet.

**Syntax**

**SET.EXTRACT**( )

**Related Functions**

SET.DATABASE&nbsp;&nbsp;&nbsp;Equivalent to clicking the Set Database
command on the Data menu in Microsoft Excel version 4.0

SET.CRITERIA&nbsp;&nbsp;&nbsp;Equivalent to clicking the Set Criteria
command on the Data menu in Microsoft Excel version 4.0

Return to [top](#Q)

# SET.LIST.ITEM

Sets the text of an item in a list box or drop-down box control.

**Syntax**

**SET.LIST.ITEM**(**text**, **index\_num**)

Text&nbsp;&nbsp;&nbsp;&nbsp;specifies the text of the item to be added.
Instead of text, an empty string may be inserted.

Index\_num&nbsp;&nbsp;&nbsp;&nbsp;is the list index of the item to be
changed, from 1 to the number of items in the list.

**Remarks**

If the list box or drop-down box was already filled using the
LISTBOX.PROPERTIES function, then changing an item with SET.LIST.ITEM
causes the fillrange contents to be discarded, leaving a list with one
non-blank element and index\_num entries.

**Related Functions**

REMOVE.LIST.ITEM&nbsp;&nbsp;&nbsp;Removes an item in a list box or
drop-down box

SELECT.LIST.ITEM&nbsp;&nbsp;&nbsp;Selects an item in a list box or in a
group box

Return to [top](#Q)

# SET.NAME

Defines a name on a macro sheet to refer to a value. The defined name
exists only on the macro sheet's list of names and does not appear in
the global list of names for the workbook. The SET.NAME function is
useful for storing values while the macro is calculating.

**Syntax**

**SET.NAME**(**name\_text**, value)

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name in the form of text that
refers to value.

Value&nbsp;&nbsp;&nbsp;&nbsp;is the value you want to store in
name\_text.

  - > If value is omitted, the name name\_text is deleted.

  - > If value is a reference, name\_text is defined to refer to that
    > reference.

> &nbsp;

**Remarks**

  - > If you want to define a name as a constant value, you can use the
    > following syntax instead of SET.NAME:

> name\_text=value
> 
> See the first two examples following.

  - > SET.NAME defines names as absolute references, even if a relative
    > reference is specified. See the third and fourth examples
    > following.

  - > If you want name\_text to refer permanently to the value of a
    > referenced cell rather than to the reference itself, you must use
    > the DEREF function. Use of DEREF prevents name\_text from
    > referring to a new value every time the contents of the referenced
    > cell changes. See the last example following.

> &nbsp;

**Tips**

  - > If you need to return an array to a macro sheet (for example, if
    > the macro needs a list of all open windows), assign a name to the
    > array instead of placing the array information in a range of
    > cells. For example:

> SET.NAME("OpenDocuments", WINDOWS()) or  
> SET.NAME("OpenDocuments", {"WORKSHEET1", "WORKSHEET2"})

  - > You can then use the INDEX function with the name you have defined
    > to access items in the array stored in the name.

  - > When you're debugging a macro and want to know the current value
    > assigned to a name created by SET.NAME, you can halt the macro,
    > click Define on the Name submenu of the Insert menu, and select
    > the name from the Define Name dialog box.

> &nbsp;

**Examples**

Each of these formulas defines the name Counter to refer to the constant
number 1 on the macro sheet:

SET.NAME("Counter", 1)

Counter=1

Each of these formulas redefines Counter to refer to the current value
of Counter plus 1:

SET.NAME("Counter", Counter+1)

Counter=Counter+1

The following macro formula defines the name Reference to refer to cell
$A$1:

SET.NAME("Reference", A1)

The following macro formula defines the name Results to refer to the
cells $A$1:$C$3:

SET.NAME("Results", A1:C3)

The following macro formula defines the name Range as the current
selection:

SET.NAME("Range", SELECTION())

If $A$1 contains the value 2, the following macro formula defines the
name Index to refer to the constant value 2:

SET.NAME("Index", DEREF(A1))

**Related Functions**

DEFINE.NAME&nbsp;&nbsp;&nbsp;Defines a name on the active worksheet or
macro sheet

SET.VALUE&nbsp;&nbsp;&nbsp;Sets the value of a cell on a macro sheet

Return to [top](#Q)

# SET.PAGE.BREAK

Equivalent to clicking the Page Break command on the Insert menu. Sets
manual page breaks. Use SET.PAGE.BREAK to override the automatic page
breaks. Setting a manual page break changes the automatic page breaks
that follow it.

The page break occurs above and to the left of the active cell and
appears as dotted lines if you have set up a printer. If the active cell
is in column A, a manual page break is added only above the cell. If the
active cell is in row 1, a manual page break is added only at the left
edge of the cell. If the row or column next to the active cell already
has a page break, SET.PAGE.BREAK takes no action.

**Syntax**

**SET.PAGE.BREAK**( )

**Related Functions**

PRINT.PREVIEW&nbsp;&nbsp;&nbsp;Previews pages and page breaks before
printing

REMOVE.PAGE.BREAK&nbsp;&nbsp;&nbsp;Removes manual page breaks

Return to [top](#Q)

# SET.PREFERRED

Changes the default format that Microsoft Excel uses when you create a
new chart or when you format a chart PREFERRED macro function. When you
use the SET.PREFERRED function, the format of the active chart becomes
the preferred format.

**Syntax**

**SET.PREFERRED**(format)

Format&nbsp;&nbsp;&nbsp;&nbsp;is the name of the format that you want as
the default format for charts. If omitted, the format of the currently
active chart is used. If format is "Built\_in", then Microsoft Excel
will use the standard, built-in chart as the default. If the chart was
created in Microsoft Excel version 4.0 and if format is "PREFERRED",
then the preferred chart format used in Microsoft Excel version 4.0 will
be used. Format is case sensitive.

**Related Function**

PREFERRED&nbsp;&nbsp;&nbsp;Changes the format of the active chart to the
preferred format

Return to [top](#Q)

# SET.PRINT.AREA

Defines the print area for the workbook&mdash;the area that prints when
you click the Print command on the File menu. Equivalent to entering a
range in the Print Area edit box on the Sheet tab in the Page Setup
dialog box, which appears when you click the Page Setup command on the
File menu.

**Syntax**

**SET.PRINT.AREA**(range)

Range&nbsp;&nbsp;&nbsp;&nbsp;is the reference to the range that you want
to be printed. If you specify no range by using a set of empty quotation
marks (""), deletes the print area.

**Remarks**

  - > If you use SET.PRINT.AREA with a multiple selection and then use
    > the PRINT function, the individual selections are printed one
    > after the other in the order they were selected.

  - > To resume printing the entire worksheet, click the Page Setup
    > command on the File menu and click the Sheet tab. Then delete the
    > range in the Print Area edit box.

> &nbsp;

**Related Functions**

PRINT&nbsp;&nbsp;&nbsp;Prints the active sheet

SET.PRINT.TITLES&nbsp;&nbsp;&nbsp;Identifies text to print as titles

Return to [top](#Q)

# SET.PRINT.TITLES

Defines the print titles for the sheet. Use SET.PRINT.TITLES if you want
Microsoft Excel to print the titles whenever it prints any cells in a
row or column that intersect the print titles area; a cell need only
share the row or column with a print title for the title to be printed
above or to the left of that cell.

**Syntax**

**SET.PRINT.TITLES**(titles\_for\_cols\_ref, titles\_for\_rows\_ref)

**SET.PRINT.TITLES**?(titles\_for\_cols\_ref, titles\_for\_rows\_ref)

Titles\_for\_cols\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the row
to be used as a title for columns.

  - > If you specify part of a row, Microsoft Excel expands the title to
    > a full row.

  - > If you omit titles\_for\_cols\_ref, Microsoft Excel uses the
    > existing row of column titles, if any.

  - > If you specify empty text (""), Microsoft Excel removes the row
    > from the print titles definition.

> &nbsp;

Titles\_for\_rows\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the
column to be used as a title for rows.

  - > If you specify part of a column, Microsoft Excel expands the title
    > to a full column.

  - > If you omit titles\_for\_rows\_ref, Microsoft Excel uses the
    > existing column of row titles, if any.

  - > If you specify empty text (""), Microsoft Excel removes the column
    > from the print titles definition.

> &nbsp;

**Remarks**

  - > SET.PRINT.TITLES operates on the current sheet. If you specify a
    > range that is invalid for the current sheet, Microsoft Excel
    > returns the \#VALUE error value.

  - > The print titles selection can be a multiple selection. Microsoft
    > Excel names this selection Print\_Titles when SET.PRINT.TITLES is
    > run.

> &nbsp;

**Related Functions**

DEFINE.NAME&nbsp;&nbsp;&nbsp;Defines a name on the active worksheet or
macro sheet

PRINT&nbsp;&nbsp;&nbsp;Prints the active sheet

SET.PRINT.AREA&nbsp;&nbsp;&nbsp;Defines the print area

Return to [top](#Q)

# SET.UPDATE.STATUS

Sets the update status of a link to automatic or manual. Use
SET.UPDATE.STATUS to change the way a link is updated.

**Syntax**

**SET.UPDATE.STATUS**(**link\_text, status**, type\_of\_link)

Link\_text&nbsp;&nbsp;&nbsp;&nbsp;is the path of the linked file for
which you want to change the update status.

Status&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and describes how you
want the link to be updated.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Status</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Update method</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Automatic</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Manual</p>
</blockquote></td>
</tr>
</tbody>
</table>

Type\_of\_link&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 4 that
specifies what type of link you want to get information about.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_of_link</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Link document type</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Not available</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>DDE/OLE link</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Not available</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Not available</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Example**

In Microsoft Excel for Windows, the following macro formula sets the
update status of the DDE link to Microsoft Word for Windows to manual:

SET.UPDATE.STATUS("WordDocument|'C:\\MEMO.DOC'\!DDE.LINK1", 2, 2)

**Related Functions**

GET.LINK.INFO&nbsp;&nbsp;&nbsp;Returns information about a link

UPDATE.LINK&nbsp;&nbsp;&nbsp;Updates a link to another document

Return to [top](#Q)

# SET.VALUE

Changes the value of a cell or cells on the macro sheet (not the
worksheet) without changing any formulas entered in those cells. Use
SET.VALUE to assign initial values and to store values during the
calculation of a macro. SET.VALUE is especially useful for initializing
a dialog box and the conditional test in a WHILE loop. SET.VALUE assigns
values to a specific reference or to the name of a reference that has
already been defined. For information about creating a new name or
entering data on a worksheet, see "Remarks" later in this topic.

**Syntax**

**SET.VALUE**(**reference**, values)

Reference&nbsp;&nbsp;&nbsp;&nbsp;specifies the cell or cells on the
macro sheet to which you want to assign a new value or values. If the
cell is empty, enters the value in the cell.

  - > If a cell in reference previously contained a formula, the formula
    > is not changed, but the value of the cell might change. See the
    > second example following.

  - > If reference is a reference to a range of cells, rather than to a
    > single cell, then values should be an array of the same size. If
    > not, Microsoft Excel expands it into multiple values using the
    > normal rules for expanding arrays. See the third example
    > following.

> &nbsp;

Values&nbsp;&nbsp;&nbsp;&nbsp;is the value or set of values to which you
want to assign the cell or cells in reference.

**Remarks**

Consider the following guidelines as you choose a function to set values
on a worksheet or macro sheet:

  - > Use SET.VALUE to assign initial values to a reference (including
    > names that have already been defined) on a macro sheet, and to
    > store values during the calculation of a macro.

  - > Use FORMULA to enter values in a worksheet cell.

  - > Use SET.NAME to change the value of a name on a macro sheet (the
    > name is created if it does not already exist). For more
    > information, see SET.NAME.

  - > Use DEFINE.NAME to create or change the value of a name on a
    > worksheet.

> &nbsp;

**Examples**

The following macro formula changes the value of cell A1 on the macro
sheet to 1:

SET.VALUE($A$1, 1)

Suppose the name TempAverage refers to a cell containing the formula
AVERAGE(Temp1, Temp2, Temp3). The following formula assigns the value 99
to this cell, even if the average of the arguments is not 99, without
changing the formula in TempAverage:

SET.VALUE(TempAverage, 99)

The preceding formula is useful if a WHILE loop or some other
conditional test depends on TempAverage and you want to force the
conditional test to have a particular result. Of course, TempAverage is
restored to its correct value as soon as it is recalculated. (Recall
that unlike formulas in a worksheet, formulas in a macro sheet are not
recalculated until the macro actually uses them.)

The following macro formula stores the values 1, 2, 3, and 4 in cells
A1:B2:

SET.VALUE($A$1:$B$2, {1, 2;3, 4})

**Related Functions**

DEFINE.NAME&nbsp;&nbsp;&nbsp;Defines a name on the active worksheet or
macro sheet

FORMULA&nbsp;&nbsp;&nbsp;Enters values into a cell or range or onto a
chart

SET.NAME&nbsp;&nbsp;&nbsp;Defines a name as a value

Return to [top](#Q)

# SHORT.MENUS

Equivalent to clicking the Short Menus command on the Options menu or
the Chart menu in Microsoft Excel version 3.0 or earlier.

**Syntax**

**SHORT.MENUS**(logical)

Return to [top](#Q)

# SHOW.ACTIVE.CELL

Scrolls the active window so the active cell becomes visible. If an
object is selected, SHOW.ACTIVE.CELL returns the \#VALUE\! error value
and halts the macro.

**Syntax**

**SHOW.ACTIVE.CELL**( )

**Related Functions**

ACTIVE.CELL&nbsp;&nbsp;&nbsp;Returns the reference of the active cell

FORMULA.GOTO&nbsp;&nbsp;&nbsp;Selects a named area or reference on any
open workbook

Return to [top](#Q)

# SHOW.BAR

Displays the specified menu bar. Use SHOW.BAR to display a menu bar you
have created with the ADD.BAR function or to display a built-in
Microsoft Excel 95 or earlier version menu bar.

**Syntax**

**SHOW.BAR**(bar\_num)

Bar\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number of the menu bar you want
to display. It can be the number of one of the Microsoft Excel built-in
menu bars, the number returned by a previously executed ADD.BAR
function, or a reference to a cell containing a previously executed
ADD.BAR function.

If bar\_num is omitted, Microsoft Excel displays the appropriate menu
bar for the active workbook as shown in the following table.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Bar_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Bar displayed</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>A sheet or macro sheet (Microsoft Excel version 4.0)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A chart (Microsoft Excel version 4.0)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>No active window</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The Info window (Microsoft Excel 95 or earlier versions)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>A sheet or macro sheet (short menus)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>A chart (short menus)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Shortcut menus 1 (for Cells, Workbook tabs, Toolbars, VB Windows)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Shortcut menus 2 (for objects)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>Shortcut menus 3 (for chart elements)</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>A sheet or macro sheet</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>A chart</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>A Visual Basic module</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>13-35</p>
</blockquote></td>
<td><blockquote>
<p>Reserved for use by shortcut menus. These numbers will return an error if a macro tries to do anything with them.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>37-51</p>
</blockquote></td>
<td><blockquote>
<p>Custom menu bar for macro use</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Remarks**

  - > When displaying a built-in menu bar, you can display only bars 1
    > or 5 if a sheet or macro sheet is active, bars 2 or 6 if a chart
    > is active, and so on. If you try to display a chart menu bar while
    > a sheet or macro sheet is active, SHOW.BAR returns an error and
    > interrupts the current macro.

  - > Displaying a custom menu bar disables automatic menu-bar switching
    > when different types of sheets are selected. For example, if a
    > custom menu bar is displayed and you switch to a chart, neither of
    > the two chart menus is automatically displayed as it would be when
    > you are using the built-in menu bars. Automatic menu-bar switching
    > is reenabled when a built-in bar is displayed using SHOW.BAR.

> &nbsp;

**Example**

The following macro formula displays short menus on a worksheet or macro
sheet:

SHOW.BAR(5)

**Related Functions**

ADD.BAR&nbsp;&nbsp;&nbsp;Adds a menu bar

DELETE.BAR&nbsp;&nbsp;&nbsp;Deletes a menu bar

SHOW.TOOLBAR&nbsp;&nbsp;&nbsp;Hides or displays a toolbar

Return to [top](#Q)

# SHOW.CLIPBOARD

Displays the contents of the Clipboard in a new window.

**Syntax**

**SHOW.CLIPBOARD**( )

**Remarks**

  - > In Microsoft Excel for Windows, the Clipboard must already be
    > running if you want to display its contents in a new window. If it
    > is not already running, you must run the SHOW.CLIPBOARD function
    > twice, once to start the Clipboard application and again to
    > display it in a new window.

  - > If the Clipboard contains cells, the window shows the size of the
    > Clipboard contents in rows and columns. If the Clipboard contains
    > text cut from the formula bar, the window displays the text.

> &nbsp;

Return to [top](#Q)

# SHOW.DETAIL

Expands or collapses the detail under the specified expand or collapse
button.

**Syntax**

**SHOW.DETAIL**(**rowcol, rowcol\_num**, expand, show\_field)

Rowcol&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies whether to
operate on rows or columns of data.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Rowcol</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Operates on</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Rows</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Columns</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>The current cell's row or column. The second argument, rowcol_num, is then ignored.</p>
</blockquote></td>
</tr>
</tbody>
</table>

Rowcol\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies the row or
column to expand or collapse. If you are in A1 mode, you must still give
the column as a number. If rowcol\_num is not a summary row or column,
SHOW.DETAIL returns the \#VALUE\! error value and interrupts the macro.

Expand&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies whether
to expand or collapse the detail under the row or column. If expand is
TRUE, Microsoft Excel expands the detail under the row or column; if
FALSE, it collapses the detail under the row or column. If expand is
omitted, the detail is expanded if it is currently collapsed and
collapsed if it is currently expanded.

Show\_Field&nbsp;&nbsp;&nbsp;&nbsp;is a string specifying the name of
the field to add to a PivotTable report, if the selection is inside a
PivotTable report. The new field is added as the new innermost field.
Available for only innermost row or column fields.

**Related Function**

SHOW.LEVELS&nbsp;&nbsp;&nbsp;Displays a specific number of levels of an
outline

Return to [top](#Q)

# SHOW.DIALOG

Runs a dialog on a dialog sheet.

**Syntax**

**SHOW.DIALOG**(dialog\_sheet)

Dialog\_sheet&nbsp;&nbsp;&nbsp;&nbsp;is the name of the dialog sheet to
run. If omitted, the active sheet will be the sheet that is run. If this
function is run on a sheet other than a dialog sheet, this function
returns the \#VALUE error value.

**Remarks**

Returns TRUE if the dialog box is closed by the user choosing an OK
button. Returns FALSE if the dialog box is cancelled by choosing the
Cancel button or the ESC key, or in Microsoft Excel for the Macintosh by
pressing COMMAND+. (period).

**Related Function**

HIDE.DIALOG&nbsp;&nbsp;&nbsp;Closes the dialog box that has the current
focus

Return to [top](#Q)

# SHOW.INFO

This function should not be used. The Info Window has been removed from
Microsoft Excel 97 or later.

**Related Functions**

FORMULA.GOTO&nbsp;&nbsp;&nbsp;Selects a named area or reference on any
open workbook

GET.CELL&nbsp;&nbsp;&nbsp;Returns information about the specified cell

SELECT&nbsp;&nbsp;&nbsp;Selects a cell, worksheet object, or chart item

Return to [top](#Q)

# SHOW.LEVELS

Displays the specified number of row and column levels of an outline.

**Syntax**

**SHOW.LEVELS**(row\_level, col\_level)

Row\_level&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of row levels of
an outline to display. If the outline has fewer levels than specified by
row\_level, Microsoft Excel shows all levels. If row\_level is zero or
omitted, no action is taken on rows.

Col\_level&nbsp;&nbsp;&nbsp;&nbsp;specifies the number of column levels
of an outline to display. If the outline has fewer levels than specified
by col\_level, Microsoft Excel shows all levels. If col\_level is zero
or omitted, no action is taken on columns.

**Remarks**

If you omit both arguments, SHOW.LEVELS returns the \#VALUE\! error
value.

**Related Function**

SHOW.DETAIL&nbsp;&nbsp;&nbsp;Expands or collapses a portion of an
outline

Return to [top](#Q)

# SHOW.TOOLBAR

Equivalent to selecting the check box corresponding to a toolbar on the
Toolbars tab in the Customize dialog box, which appears when you select
the Customize command (View menu, Toolbars submenu). Hides or displays a
toolbar. Use SHOW.TOOLBAR to display or hide a menu bar you have created
with the ADD.BAR function or to display a built-in Microsoft Excel 95 or
earlier version toolbar.

**Syntax**

**SHOW.TOOLBAR**(**bar\_id, visible**, dock, x\_pos, y\_pos, width,
protect, tool\_tips, large\_buttons, color\_buttons)

Bar\_id&nbsp;&nbsp;&nbsp;&nbsp;is a number or name of a toolbar
corresponding to the toolbars you want to display. For detailed
information about bar\_id, see ADD.TOOL.

Visible&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that, if TRUE,
specifies that the toolbar is visible or, if FALSE, specifies that the
toolbar is hidden.

Dock&nbsp;&nbsp;&nbsp;&nbsp;specifies the docking location of the
toolbar.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Dock</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Position of toolbar</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Top of workspace</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Left edge of workspace</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Right edge of workspace</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Bottom of workspace</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Floating (not docked)</p>
</blockquote></td>
</tr>
</tbody>
</table>

X\_pos&nbsp;&nbsp;&nbsp;&nbsp;specifies the horizontal position of the
toolbar.

  - > If the toolbar is docked (not floating), x\_pos is measured
    > horizontally from the left edge of the toolbar to the left edge of
    > the toolbar's docking area.

  - > If the toolbar is floating, x\_pos is measured horizontally from
    > the left edge of the toolbar to the right edge of the rightmost
    > toolbar in the left docking area.

  - > X\_pos is measured in points. A point is 1/72nd of an inch.

Y\_pos&nbsp;&nbsp;&nbsp;&nbsp;specifies the vertical position of the
toolbar.

  - > If the toolbar is docked, y\_pos is measured vertically from the
    > top edge of the toolbar to the top edge of the toolbar's docking
    > area.

  - > If the toolbar is floating, y\_pos is measured vertically from the
    > top edge of the toolbar to the top edge of the Microsoft Excel
    > workspace.

  - > Y\_pos is measured in points.

> &nbsp;

Width&nbsp;&nbsp;&nbsp;&nbsp;specifies the width of the toolbar and is
measured in points. If you omit width, Microsoft Excel uses the existing
width setting.

Protect&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the degree to
which you can modify a toolbar and its buttons. Each succeeding protect
number retains the protection status of its previous numbers. For
example, a protect status of 3 (a toolbar cannot become docked if it is
floating) assumes the protection status of 0, 1, and 2 as well.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Protect</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Description</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>0</p>
</blockquote></td>
<td><blockquote>
<p>Default. Toolbars can be re-shaped, docked, and floating. Toolbar buttons can be removed from and moved to the toolbar.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Toolbars can be re-shaped, docked, and floating. Toolbar buttons can not be removed from nor moved to the toolbar.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A floating toolbar cannot be re-shaped. It can be docked.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>A floating toolbar cannot be docked. If it is already docked, it cannot become floating.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The toolbar cannot be moved at all. If it is already floating, it cannot be re-shaped or moved. If it is docked, it cannot become un-docked.</p>
</blockquote></td>
</tr>
</tbody>
</table>

Tool\_tips&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to
the Show Screentips On Toolbars check box on the Options tab. If TRUE,
ScreenTips will be displayed. If FALSE, ScreenTips will not be
displayed.

Large Buttons&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds
to the Large Icons check box on the Options tab. If TRUE, large icons
will be displayed. If FALSE, large icons will not be displayed.

Color\_buttons&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that
corresponds to the Color Toolbars check box. If TRUE, the toolbar
buttons will be displayed in color. If FALSE, the toolbar buttons will
not be displayed in color. This argument is for compatibility with
Microsoft Excel version 5.0.

**Related Functions**

ADD.BAR&nbsp;&nbsp;&nbsp;Adds a menu bar

ADD.TOOLBAR&nbsp;&nbsp;&nbsp;Creates a new toolbar with the specified
tools

Return to [top](#Q)

# SIZE

Equivalent to clicking the Size command on the Control menu in Microsoft
Excel for Windows version 3.0 or earlier or to changing the size of a
window by dragging its border. In Microsoft Excel for the Macintosh
version 3.0 or earlier, equivalent to changing the size of a window by
dragging its size box. This function is included only for macro
compatibility and will be converted to WINDOW.SIZE when you open older
macro sheets. For more information, see WINDOW.SIZE.

**Syntax**

**SIZE**(**width,height**,window\_text)

**SIZE**?(width,height,window\_text)

**Related Function**

WINDOW.SIZE&nbsp;&nbsp;&nbsp;Changes the size of the active window

Return to [top](#Q)

# SLIDE.COPY.ROW

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Copy Row button on a slide show sheet. Copies
the selected slides, each of which is defined on a single row, to the
Clipboard.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.COPY.ROW**( )

**Remarks**

  - > SLIDE.COPY.ROW, SLIDE.CUT.ROW, SLIDE.DELETE.ROW, and
    > SLIDE.PASTE.ROW return TRUE if successful, or FALSE if not
    > successful. If the active sheet is not a slide show or is
    > protected, these functions return the \#N/A error value. If the
    > current selection is not valid, these functions return the
    > \#VALUE\! error value.

> &nbsp;

**Related Functions**

SLIDE.CUT.ROW&nbsp;&nbsp;&nbsp;Cuts the selected slides and pastes them
onto the Clipboard

SLIDE.DEFAULTS&nbsp;&nbsp;&nbsp;Specifies default values for the active
slide show sheet

SLIDE.DELETE.ROW&nbsp;&nbsp;&nbsp;Deletes the selected slides

SLIDE.EDIT&nbsp;&nbsp;&nbsp;Changes the attributes of the selected slide

SLIDE.GET&nbsp;&nbsp;&nbsp;Returns information about a slide or slide
show

SLIDE.PASTE&nbsp;&nbsp;&nbsp;Pastes the contents of the Clipboard onto a
slide

SLIDE.PASTE.ROW&nbsp;&nbsp;&nbsp;Pastes previously cut or copied slides
onto the current selection

SLIDE.SHOW&nbsp;&nbsp;&nbsp;Starts a slide show in the active sheet

Return to [top](#Q)

# SLIDE.CUT.ROW

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Cut Row button on a slide show sheet. Cuts
the selected slides, each of which is defined on a single row, and
pastes them onto the Clipboard. For more information, see
SLIDE.COPY.ROW.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.CUT.ROW**( )

**Related Function**

SLIDE.COPY.ROW&nbsp;&nbsp;&nbsp;Copies the selected slides and pastes
them onto the Clipboard

Return to [top](#Q)

# SLIDE.DEFAULTS

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Set Defaults button on a slide show sheet.
Specifies the default values for the transition effect, speed, advance
rate, and sound on the active slide show sheet.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.DEFAULTS**(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

**SLIDE.DEFAULTS**?(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

For a description of the arguments, see SLIDE.PASTE. If an argument is
omitted, its default value is not changed.

**Remarks**

  - > SLIDE.DEFAULTS returns TRUE if it successfully changes the default
    > values, or FALSE if you click the Cancel button when using the
    > dialog-box form. If the active sheet is not a slide show or is
    > protected, SLIDE.DEFAULTS returns the \#N/A error value.

Return to [top](#Q)

# SLIDE.DELETE.ROW

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Delete Row button on a slide show sheet.
Deletes the selected slides, each of which is defined on a single row.
For more information, see SLIDE.COPY.ROW.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.DELETE.ROW**( )

**Related Function**

SLIDE.COPY.ROW&nbsp;&nbsp;&nbsp;Copies the selected slides and pastes
them onto the Clipboard

Return to [top](#Q)

# SLIDE.EDIT

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Edit button in a slide show sheet. Gives the
currently selected slide the attributes you specify.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.EDIT**(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

**SLIDE.EDIT**?(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

For a description of the arguments, see SLIDE.PASTE.

**Remarks**

  - > SLIDE.EDIT returns TRUE if it successfully edits the slide, or
    > FALSE if you click the Cancel button when using the dialog-box
    > form. If the active sheet is not a slide show or is protected,
    > SLIDE.EDIT returns the \#N/A error value. If the current selection
    > is not a valid slide, SLIDE.EDIT returns the \#VALUE error value.

> &nbsp;

**Related Function**

SLIDE.PASTE&nbsp;&nbsp;&nbsp;Pastes the contents of the Clipboard onto a
slide

Return to [top](#Q)

# SLIDE.GET

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Returns the specified information about a slide show or a specific slide
in the slide show.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.GET**(**type\_num**, name\_text, slide\_num)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the type of
information you want.

These values of type\_num return information about a slide show.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Type of information</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Number of slides in the slide show</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A two-item horizontal array containing the numbers of the first and last slides in the current selection, or the #VALUE error value if the selection is nonadjacent</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Version number of the Slide Show add-in that created the slide show sheet</p>
</blockquote></td>
</tr>
</tbody>
</table>

These values of type\_num return information about a specific slide in
the slide show.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Type of information</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Transition effect number</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Transition effect name</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Transition effect speed</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Number of seconds the slide is displayed before advancing</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Name of the sound file associated with the slide, or empty text ("") if none is specified (in Microsoft Excel for the Macintosh, this includes the number or name of the sound resource within the sound file)</p>
</blockquote></td>
</tr>
</tbody>
</table>

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of an open slide show
sheet for which you want information. If name\_text is omitted, it is
assumed to be the active sheet.

Slide\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number of the slide about which
you want information.

  - > If slide\_num is omitted, it is assumed to be the slide associated
    > with the active cell on the sheet specified by name\_text.

  - > If type\_num is 1 through 3, slide\_num is ignored.

Return to [top](#Q)

# SLIDE.PASTE

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Paste button on a slide show sheet. Pastes
the contents of the Clipboard as the next available slide of the active
slide show sheet, and gives the slide the attributes you specify.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.PASTE**(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

**SLIDE.PASTE**?(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

Effect\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the transition
effect you want to use when displaying the slide.

  - > The numbers correspond to the effects in the Effect list in the
    > Edit Slide dialog box. The first effect in the list is 1 (None).

  - > If effect\_num is omitted, the default setting is used.

> &nbsp;

Speed\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 10 specifying
the speed of the transition effect.

  - > If speed\_num is omitted, the default setting is used.

  - > If speed\_num is greater than 10, Microsoft Excel uses the value
    > 10 anyway.

  - > If effect\_num is 1 (none), speed\_num is ignored.

> &nbsp;

Advance\_rate\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying how
long (in seconds) the slide is displayed before advancing to the next
one.

  - > If advance\_rate\_num is omitted, the default setting is used.

  - > If advance\_rate\_num is 0, you must press a key or click with the
    > mouse to advance to the next slide.

> &nbsp;

Soundfile\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a file enclosed in
quotation marks and specifies sound that will be played when the slide
is displayed.

  - > If soundfile\_text is omitted, Microsoft Excel plays the default
    > sound defined for the slide show sheet, if any.

  - > If soundfile\_text is empty text (""), no sound is played.

  - > In Microsoft Excel for the Macintosh, soundfile\_text also
    > includes the number or name of the sound resource to play in the
    > file.

> &nbsp;

Resource&nbsp;&nbsp;&nbsp;&nbsp;is the number or name of a sound
resource in soundfile\_text.

  - > This argument applies only to Microsoft Excel for the Macintosh.

  - > If resource is omitted, Microsoft Excel uses the first resource in
    > the file.

  - > If the file does not contain a sound resource with the specified
    > name or number, Microsoft Excel halts the macro and displays an
    > error message.

> &nbsp;

**Remarks**

  - > SLIDE.PASTE returns TRUE if it successfully pastes the slide, or
    > FALSE if you click the Cancel button when using the dialog-box
    > form. If the active sheet is not a slide show or is protected,
    > SLIDE.PASTE returns the \#N/A error value. If the Clipboard format
    > is not compatible with the slide show sheet's format, SLIDE.PASTE
    > returns the \#VALUE error value.

> &nbsp;

**Examples**

In Microsoft Excel for Windows, the following macro formula pastes the
contents of the Clipboard into the active slide show sheet. The slide's
transition effect is fade, at a speed of 8; it is displayed for five
seconds; and Microsoft Excel plays the specified sound file:

SLIDE.PASTE(3, 8, 5, "C:\\SLIDES\\SOUND\\MACHINES.WAV")

In Microsoft Excel for the Macintosh, the formula is:

SLIDE.PASTE(3, 8, 5, "HARD DISK:SLIDES:SOUND:MACHINE SOUNDS")

Return to [top](#Q)

# SLIDE.PASTE.ROW

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Paste Row button on a slide show sheet.
Pastes previously cut or copied slides onto the current selection. For
more information, see SLIDE.COPY.ROW.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.PASTE.ROW**( )

**Related Function**

SLIDE.COPY.ROW&nbsp;&nbsp;&nbsp;Copies the selected slides and pastes
them onto the Clipboard

Return to [top](#Q)

# SLIDE.SHOW

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Start Show button on a slide show sheet.
Starts the slide show in the active sheet.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.SHOW**(initialslide\_num, repeat\_logical, dialogtitle\_text,
allownav\_logical, allowcontrol\_logical)

**SLIDE.SHOW**?(initialslide\_num, repeat\_logical, dialogtitle\_text,
allownav\_logical, allowcontrol\_logical)

All arguments except dialogtitle\_text correspond to options and
settings in the Start Show dialog box.

Initialslide\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to the
number of slides in the slide show and specifies which slide to display
first. If omitted, it is assumed to be 1.

Repeat\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying
whether to repeat or end the slide show after displaying the last slide.
If repeat\_logical is TRUE, the slide show repeats; if FALSE or omitted,
the slide show ends.

Dialogtitle\_text&nbsp;&nbsp;&nbsp;&nbsp;is text enclosed in quotation
marks that specifies the title of the dialog boxes displayed during the
slide show. If dialogtitle\_text is omitted, it is assumed to be "Slide
Show".

Allownav\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying
whether to enable or disable navigational keys (arrow keys, PAGE UP,
PAGE DOWN, and so on) or the mouse during the slide show. If
allownav\_logical is TRUE or omitted, you can press navigational keys or
use the mouse to move between slides; if FALSE, all movement is
controlled by the slide show sheet settings.

Allowcontrol\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value
specifying whether to enable or disable the Slide Show Options dialog
box during the slide show. If allowcontrol\_logical is TRUE or omitted,
you can press ESC to interrupt the slide show and display the dialog
box; if FALSE, pressing ESC stops the slide show but does not display
the dialog box.

**Tip&nbsp;&nbsp;&nbsp;**If you want to display the last slide in a show
but don't know its number, use SLIDE.GET(1) as the initialslide\_num
argument.

**Remarks**

SLIDE.SHOW returns the values shown in the following table:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Situation</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returned value</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>The slide show ends normally.</p>
</blockquote></td>
<td><blockquote>
<p>TRUE</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>You press the Cancel button when using the dialog-box form.</p>
</blockquote></td>
<td><blockquote>
<p>FALSE</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>The active sheet is not a slide show or is protected.</p>
</blockquote></td>
<td><blockquote>
<p>#N/A</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>You interrupt the slide show, and then stop it.</p>
</blockquote></td>
<td><blockquote>
<p>1</p>
</blockquote></td>
</tr>
</tbody>
</table>

Return to [top](#Q)

# SOLVER.ADD

Equivalent to clicking the Solver command on the Tools menu and clicking
the Add button in the Solver Parameters dialog box. Adds a constraint to
the current problem. For an explanation of constraints, see "Remarks"
later in this topic.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.ADD**(**cell\_ref, relation**, formula)

Cell\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to a cell or range of
cells on the active sheet and forms the left side of the constraint.

Relation&nbsp;&nbsp;&nbsp;&nbsp;specifies the arithmetic relationship
between the left and right sides, or whether cell\_ref must be an
integer.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Relation</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Arithmetic relationship</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>&lt;=</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>=</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>&gt;=</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Int (cell_ref is an integer)</p>
</blockquote></td>
</tr>
</tbody>
</table>

Formula&nbsp;&nbsp;&nbsp;&nbsp;is the right side of the constraint and
will often be a single number, but it may be a formula (as text) or a
reference to a range of cells.

  - > If relation is 4, cell\_ref must be a subset of the references in
    > the By Changing cells text box.

  - > if relation is 4, formula must be either "=integer" or "integer".

  - > Any cell reference in a formula must use the R1C1 reference style.

  - > If formula is a reference to a range of cells, the number of cells
    > in the range usually matches the number of cells in cell\_ref,
    > although the shape of the areas need not be the same. For example,
    > cell\_ref could be a row and formula could refer to a column, as
    > long as the number of cells is the same. Formula can also be a
    > single reference, as in the following relationship:
    > &nbsp;A1:A4&nbsp;\<=&nbsp;B1.

> &nbsp;

**Remarks**

  - > The SOLVER.ADD, SOLVER.CHANGE, and SOLVER.DELETE functions
    > correspond to the Add, Change, and Delete buttons in the Solver
    > Parameters dialog box. You use these functions to define
    > constraints. For many macro applications, however, you may find it
    > more convenient to load the problem specifications from the sheet
    > in a single step using the SOLVER.LOAD function.

  - > Each constraint is uniquely identified by the combination of the
    > cell reference on the left and the relationship (\<=, =, or \>=)
    > between its left and right sides, or the cell reference may be
    > defined as an integer only. This takes the place of selecting the
    > appropriate constraint in the Tools Solver Parameters dialog box.
    > You can manipulate the constraints with SOLVER.CHANGE or
    > SOLVER.DELETE. The constraints in a Solver problem can refer to a
    > maximum of 400 cells.

> &nbsp;

**Related Function**

SOLVER.DELETE&nbsp;&nbsp;&nbsp;Deletes an existing constraint

Return to [top](#Q)

# SOLVER.CHANGE

Equivalent to clicking the Solver command on the Tools menu and clicking
the Change button in the Solver Parameters dialog box. Changes the right
side of an existing constraint.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.CHANGE**(**cell\_ref, relation**, formula)

For an explanation of the arguments and constraints, see SOLVER.ADD.

**Remarks**

  - > If the combination of cell\_ref and relation does not match any
    > existing constraint, the function returns the value 4 and no
    > action is taken.

  - > To change the cell\_ref or relation of an existing constraint, use
    > SOLVER.DELETE to delete the old constraint and then use SOLVER.ADD
    > to add the constraint in the form you want.

> &nbsp;

**Related Functions**

SOLVER.DELETE&nbsp;&nbsp;&nbsp;Deletes an existing constraint

SOLVER.ADD&nbsp;&nbsp;&nbsp;Adds a constraint to the current problem

Return to [top](#Q)

# SOLVER.DELETE

Equivalent to clicking the Solver command on the Tools menu and clicking
the Delete button in the Solver Parameters dialog box. Deletes an
existing constraint.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.DELETE**(**cell\_ref, relation**, formula)

For an explanation of the arguments and constraints, see SOLVER.ADD.

**Remarks**

If the combination of cell\_ref and relation does not match any existing
constraint, the function returns the value 4 and no action is taken. If
the constraint is found, it is deleted, and the function returns the
value 0.

**Related Function**

SOLVER.ADD&nbsp;&nbsp;&nbsp;Adds a constraint to the current problem

Return to [top](#Q)

# SOLVER.FINISH

Equivalent to clicking OK in the Solver Results dialog box that appears
when the solution process is complete. The dialog-box form displays the
dialog box with the arguments that you supply as defaults. This function
must be used if you supplied the value TRUE for the userfinish argument
to SOLVER.SOLVE.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.FINISH**(keep\_final, report\_array)

**SOLVER.FINISH**?(keep\_final, report\_array)

Keep\_final&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
whether to keep the final solution. If keep\_final is 1 or omitted, the
final solution values are kept in the changing cells. If keep\_final is
2, the final solution values are discarded and the former values of the
changing cells are restored.

Report\_array&nbsp;&nbsp;&nbsp;&nbsp;is an array argument specifying
what reports to create when Solver is finished.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>If report_array is</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Microsoft Excel creates</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>{1}</p>
</blockquote></td>
<td><blockquote>
<p>An answer report</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>{2}</p>
</blockquote></td>
<td><blockquote>
<p>A sensitivity report</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>{3}</p>
</blockquote></td>
<td><blockquote>
<p>A limit report</p>
</blockquote></td>
</tr>
</tbody>
</table>

Any combination of these produces multiple reports. For example, if
report\_array is {1, 2}, Microsoft Excel creates an answer report and a
sensitivity report.

**Related Function**

SOLVER.SOLVE&nbsp;&nbsp;&nbsp;Equivalent to clicking the Solver command
on the Tools menu and clicking the Solve button in the Solver Parameters
dialog box

Return to [top](#Q)

# SOLVER.GET

Returns information about current settings for Solver. The settings are
specified in the Solver Parameters and Solver Options dialog boxes.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.GET**(**type\_num**, sheet\_name)

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the type of
information you want.

The following settings are specified in the Solver Parameters dialog
box.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the Equal To option<br />
1 = Max<br />
2 = Min<br />
3 = Value of</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>The value in the Value Of box</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The reference (as a multiple reference if necessary) in the By Changing Cells box</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>The number of constraints</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>An array of the left sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>An array of numbers corresponding to the relationships between the left and right sides of the constraints:<br />
1 = &lt;=<br />
2 = =<br />
3 = &gt;=<br />
4 = int</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>An array of the right sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
</tbody>
</table>

The following settings are specified in the Solver Options dialog box:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>The maximum number of iterations</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>The precision</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>The integer tolerance value</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Assume Linear Model check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Show Iteration Results check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>15</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Use Automatic Scaling check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of estimates:<br />
1 = Tangent<br />
2 = Quadratic</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>17</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of derivatives:<br />
1 = Forward<br />
2 = Central</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>18</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of search:<br />
1 = Quasi-Newton<br />
2 = Conjugate Gradient</p>
</blockquote></td>
</tr>
</tbody>
</table>

Sheet\_name&nbsp;&nbsp;&nbsp;&nbsp;is the name of a sheet that contains
the scenario for which you want information. If sheet\_name is omitted,
it is assumed to be the active sheet.

Return to [top](#Q)

# SOLVER.LOAD

Equivalent to clicking the Solver command on the Tools menu, clicking
the Options button in the Solver Parameters dialog box, and clicking the
Load Model button in the Solver Options dialog box. Loads Solver problem
specifications that you have previously saved on the worksheet.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.LOAD**(**load\_area**)

Load\_area&nbsp;&nbsp;&nbsp;&nbsp;is a reference on the active sheet to
a range of cells from which you want to load a complete problem
specification.

  - > The first cell in load\_area contains a formula for the Set Cell
    > box; the second cell contains a formula for the changing cells;
    > subsequent cells contain constraints in the form of logical
    > formulas. The last cell optionally contains an array of Solver
    > option values. The order of the Solver option values is the same
    > as the top-to-bottom order in the Solver Options dialog box.

  - > Although load\_area must be on the active sheet, it need not be
    > the current selection.

Return to [top](#Q)

# SOLVER.OK

Equivalent to clicking the Solver command on the Tools menu and
specifying options in the Solver Parameters dialog box. Specifies basic
Solver options, except that constraints are added via SOLVER.ADD.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.OK**(set\_cell, max\_min\_val, value\_of, **by\_changing**)

**SOLVER.OK**?(set\_cell, max\_min\_val, value\_of, by\_changing)

Set\_cell&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the Set Target Cell box
in the Solver Parameters dialog box.

  - > Set\_cell must be a reference to a cell on the active worksheet.

  - > If you enter a cell reference, you must also enter a value for
    > max\_min\_val. If you do not enter a cell, you must include three
    > commas before the by\_changing value.

> &nbsp;

Max\_min\_val&nbsp;&nbsp;&nbsp;&nbsp;corresponds to the options Max,
Min, and Value Of in the Solver Parameters dialog box. Use this option
only if you entered a reference for set\_cell.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Max_min_val</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Option specified</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Maximize</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Minimize</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Match specific value</p>
</blockquote></td>
</tr>
</tbody>
</table>

Value\_of&nbsp;&nbsp;&nbsp;&nbsp;is a number that becomes the target for
the cell in the Set Target Cell box if max\_min\_val is 3. Value\_of is
ignored if the cell is being maximized or minimized.

By\_changing&nbsp;&nbsp;&nbsp;&nbsp;indicates the changing cells, as
entered in the By Changing Cells box. By\_changing must refer to a cell
or range of cells on the active worksheet, and can be a multiple
selection.

**Remarks**

The constraints in a Solver problem can refer to a maximum of 400 cells.

**Related Function**

SOLVER.SOLVE&nbsp;&nbsp;&nbsp;Returns an integer value indicating the
condition that caused Solver to stop

Return to [top](#Q)

# SOLVER.OPTIONS

Equivalent to clicking the Solver command on the Tools menu and then
clicking the Options button in the Solver Parameters dialog box.
Specifies the available options.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.OPTIONS**(max\_time, iterations, precision, assume\_linear,
step\_thru, estimates, derivatives, search, int\_tolerance, scaling)

The arguments correspond to the options in the dialog box. If an
argument is omitted, Microsoft Excel uses an appropriate value based on
the current situation. If any of the arguments are the wrong type, the
function returns the \#N/A error value. If an argument has the correct
type but an invalid value, the function returns a positive integer
corresponding to its position. A zero indicates all options were
accepted.

Max\_time&nbsp;&nbsp;&nbsp;&nbsp;must be an integer greater than zero
and less than 32768. It corresponds to the Max Time box.

Iterations&nbsp;&nbsp;&nbsp;&nbsp;must be an integer greater than zero
and less than 32768. It corresponds to the Iterations box.

Precision&nbsp;&nbsp;&nbsp;&nbsp;must be a number between zero and one,
but not equal to zero or one. It corresponds to the Precision box.

Assume\_linear&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding
to the Assume Linear Model check box and allows Solver to arrive at a
solution more quickly. If TRUE, Solver assumes that the underlying model
is linear; if FALSE, it does not.

Step\_thru&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Show Iteration Results check box. If you have supplied SOLVER.SOLVE
with a valid command macro reference, your macro will be called each
time Solver pauses. If TRUE, Solver pauses at each trial solution; if
FALSE, it does not.

Estimates&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and corresponds to
the Estimates options: 1 for the Tangent option and 2 for the Quadratic
option.

Derivatives&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and corresponds
to the Derivatives options: 1 for the Forward option and 2 for the
Central option.

Search&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and corresponds to
the Search options: 1 for the Quasi-Newton option and 2 for the
Conjugate Gradient option.

Int\_tolerance&nbsp;&nbsp;&nbsp;&nbsp;is a decimal number corresponding
to the Tolerance box in the Solver Options dialog box, and must be
between zero and 1, inclusively. This argument applies only if integer
constraints have been defined.

Scaling&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
Use Automatic Scaling check box. If scaling is TRUE, then if two or more
constraints differ by several orders of magnitude, Solver scales the
constraints to similar orders of magnitude during computation. If
scaling is FALSE, Solver calculates normally.

Return to [top](#Q)

# SOLVER.RESET

Equivalent to clicking the Solver command on the Tools menu and clicking
the Reset All button in the Solver Parameters dialog box. Erases all
cell selections and constraints from the Solver Parameters dialog box
and restores all the settings in the Solver Options dialog box to their
defaults.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.RESET**( )

Return to [top](#Q)

# SOLVER.SAVE

Equivalent to clicking the Solver command on the Tools menu, clicking
the Options button in the Solver Parameters dialog box, and clicking the
Save Model button in the Solver Options dialog box. Saves the Solver
problem specifications on the worksheet.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.SAVE** (**save\_area**)

Save\_area&nbsp;&nbsp;&nbsp;&nbsp;is a reference on the active sheet to
a range of cells or to the upper-left corner of a range of cells into
which you want to paste the current problem specification.

  - > If you specify only one cell for save\_area, the area is extended
    > downwards for as many cells as are required to hold the problem
    > specifications (3 plus the number of constraints).

  - > If you specify more than one cell and if the area is too small,
    > the last constraints (in alphabetic order by cell reference) or
    > options will be omitted and the function will return a nonzero
    > value.

  - > Save\_area must be on the active worksheet, but it need not be the
    > current selection.

Return to [top](#Q)

# SOLVER.SOLVE

Equivalent to clicking the Solver command on the Tools menu and clicking
the Solve button in the Solver Parameters dialog box. If successful,
returns an integer value indicating the condition that caused Solver to
stop as described in "Remarks" later in this topic.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.SOLVE**(user\_finish, show\_ref)

User\_finish&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying
whether to display the Solver Results dialog box.

  - > If user\_finish is TRUE, SOLVER.SOLVE returns its integer value
    > without displaying anything. Your macro should decide what action
    > to take (for example, by examining the return value or presenting
    > its own dialog box); it must call SOLVER.FINISH in any case to
    > restore the sheet to its proper state.

  - > If user\_finish is FALSE or omitted, Solver displays the Solver
    > Results dialog box, which allows you to keep or discard the final
    > solution and run reports.

> &nbsp;

Show\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a macro to be called in place of the
Show Trial Solution dialog box. It is used when you want to regain
control whenever Solver finds a new intermediate solution value.

  - > For this argument to have an effect, the Show Iteration Results
    > check box must be selected in the Solver Options dialog box. This
    > can be done manually by selecting the check box, or automatically
    > by calling SOLVER.OPTIONS in your macro.

  - > The macro you call can inspect the current solution values on the
    > sheet or take other actions such as saving or charting the
    > intermediate values. It must return the value TRUE with a
    > statement such as =RETURN(TRUE) if the solution process is to
    > continue, or FALSE if the solution process should stop at this
    > point.

> &nbsp;

**Remarks**

If a problem has not been completely defined, SOLVER.SOLVE returns the
\#N/A error value. Otherwise, the Solver application is started and the
problem specifications are passed to it. When the solution process is
complete, SOLVER.SOLVE returns an integer value indicating the stopping
condition:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Value</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Stopping condition</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>0</p>
</blockquote></td>
<td><blockquote>
<p>Solver found a solution. All constraints and optimality conditions are satisfied.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Solver has converged to the current solution. All constraints are satisfied.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Solver cannot improve the current solution. All constraints are satisfied.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Stop chosen when the maximum iteration limit was reached.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The Set Cells values do not converge.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Solver could not find a feasible solution.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Solver stopped at user's request.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>The conditions for Assume Linear Model are not satisfied.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>The problem is too large for Solver to solve.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>Solver encountered an error value in a target or constraint cell.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>Stop chosen when the maximum time limit was reached.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>There is not enough memory available to solve the problem.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>Another Excel instance is using SOLVER.DLL. Try again later.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>Error in model. Please verify that all cells and constraints are valid.</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Related Function**

SOLVER.FINISH&nbsp;&nbsp;&nbsp;Equivalent to clicking OK in the Solver
Results dialog box that appears when the solution process is complete

Return to [top](#Q)

# SORT

Equivalent to clicking the Sort command on the Data menu. Sorts the rows
or columns of the selection according to the contents of a key row or
column within the selection. Use SORT to rearrange information into
ascending or descending order.

**Syntax 1**

For Worksheet and macro sheets

**SORT**(orientation, key1, order1, key2, order2, key3, order3, header,
custom, case)

**SORT**?(orientation, key1, order1, key2, order2, key3, order3, header,
custom, case)

**Syntax 2**

For PivotTable reports

**SORT**(orientation, key1, order1, type, custom)

**SORT**?(orientation, key1, order1, type, custom)

Orientation&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying whether to
sort by rows or columns. Enter 1 to sort top to bottom or 2 to sort left
to right.

Key1&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cell or cells you want
to use as the first sort key. The sort key identifies which column to
sort by when sorting rows or which row to sort by when sorting columns.
For a PivotTable report, if type is 1, then key1 is a cell reference
which indicates what value to sort by. There are two ways to specify
sort keys:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type of key</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Examples</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>An R1C1-style reference in the form of text. If the reference is relative, it is assumed to be relative to the active cell in the selection.</p>
</blockquote></td>
<td><blockquote>
<p>"C2" or "C[1]" or "Price"</p>
</blockquote></td>
</tr>
</tbody>
</table>

Order1&nbsp;&nbsp;&nbsp;&nbsp;specifies whether to sort the row or
column containing key1 in ascending or descending order. Enter 1 to sort
in ascending order or 2 to sort in descending order.

Key2, order2, key3, and order3&nbsp;&nbsp;&nbsp;&nbsp;are similar to
key1 and order1. Key2 specifies the second sort key, and order2
specifies whether to sort the row or column containing key2 in ascending
or descending order. Key3 and order3 work similarly.

Header&nbsp;&nbsp;&nbsp;&nbsp;is a number indicating how Microsoft Excel
is to handle headers on list.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Header</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Defined</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>0</p>
</blockquote></td>
<td><blockquote>
<p>Microsoft Excel will guess if there is a header</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Forces Microsoft Excel to assume there is a header</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2 or omitted</p>
</blockquote></td>
<td><blockquote>
<p>Forces Microsoft Excel to assume there is no header</p>
</blockquote></td>
</tr>
</tbody>
</table>

Type&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying whether to sort the
field by labels or values. Use one to sort by values or two to sort by
labels.

Custom&nbsp;&nbsp;&nbsp;&nbsp;is a number that specifies what kind of
custom sorting you want. This corresponds to the First Key Sort Order
drop-down box in the Sort Options dialog box. For a PivotTable report,
custom is a number indicating what custom sort order to use when sorting
labels.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Number</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Type of sort</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
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
<p>Weekdays in abbreviated form ("Sun", "Mon", and so on)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Weekdays</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Months in abbreviated form ("Jan" "Feb", and so on)</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Months</p>
</blockquote></td>
</tr>
</tbody>
</table>

Case&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that determines whether
the sort is case sensitive. If TRUE, the sort is case sensitive. If
FALSE or omitted, the sort will not be case sensitive.

**Tip&nbsp;&nbsp;&nbsp;**If you want to sort using more than three keys,
then sort the data three keys at a time, starting with the least
important group of keys and progressing to the most important group, but
listing the most important key first within each group.

**Remarks**

In the dialog box form of this function, if the header argument is
omitted, then Microsoft Excel will guess whether or not there are
headers.

Return to [top](#Q)

# SOUND.NOTE

This function should not be used in Microsoft Excel 97 or later because
sound notes are available only in Microsoft Excel 95 or earlier
versions.

Records sound into or erases sound from a cell note or imports sound
from another file into a cell note. This function requires that you have
recording hardware installed in your computer, and you must be running
Microsoft Windows version 3.1 or later, or Apple system software version
6.07 or later.

**Syntax 1**

Recording or erasing sound

**SOUND.NOTE**(cell\_ref, erase\_snd)

**Syntax 2**

Importing sound from another file

**SOUND.NOTE**(cell\_ref, file\_text, resource)

Cell\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cell containing a
note into which you want to record or import sounds or from which you
want to erase a sound.

Erase\_snd&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether
to erase the sound in the note. If erase\_snd is TRUE, Microsoft Excel
erases only the sound from the note. If FALSE or omitted, Microsoft
Excel displays the Record dialog box so that you can record sound into
the note.

File\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a file containing
sounds.

Resource&nbsp;&nbsp;&nbsp;&nbsp;is the number or name of a sound
resource in file\_text that you want to import into your note.

  - > This argument applies only to Microsoft Excel for the Macintosh.

  - > If resource is omitted, Microsoft Excel uses the first resource in
    > the file.

  - > If the file does not contain a sound resource with the specified
    > name or number, Microsoft Excel halts the macro and displays an
    > error message.

**Remarks**

  - > To find out if a cell has sound attached to it, use GET.CELL(47).

  - > Sounds notes are not available in Microsoft Excel 97 or later.

**Examples**

The following macro formula erases the sound, if present, from cell A1
on the active sheet:

SOUND.NOTE(\!$A$1, TRUE)

The following macro formula displays the Record dialog box so that you
can record sound into a note for cell A1 on the active sheet:

SOUND.NOTE(\!$A$1)

In Microsoft Excel for Windows, the following macro formula imports the
sound from a file named CHIMES.WAV into a note for the cell named
Doorbell on the active sheet:

SOUND.NOTE(\!Doorbell, "C:\\SOUNDS\\CHIMES.WAV")

In Microsoft Excel for the Macintosh, the following macro formula
imports a sound called Chimes from a file named SOFT SOUNDS into a note
for the cell named Doorbell on the active sheet:

SOUND.NOTE(\!Doorbell, "HARD DISK:SOUNDS:SOFT SOUNDS", "Chimes")

**Related Functions**

NOTE&nbsp;&nbsp;&nbsp;Creates or changes a cell note

SOUND.PLAY&nbsp;&nbsp;&nbsp;Plays the sound from a cell note or a file

Return to [top](#Q)

# SOUND.PLAY

This function should not be used in Microsoft Excel 97 or later because
sound notes are available only in Microsoft Excel 95 or earlier
versions.

Plays the sound from a cell note or a file. Equivalent to clicking the
Note command on the Insert menu and clicking the Play button, or
clicking the Note command on the Insert menu, clicking the Import
button, and then opening a file, selecting a sound, and clicking the
Play button. To play sounds in Microsoft Excel for Windows, you must
have a sound board installed in your computer.

**Syntax**

**SOUND.PLAY**(cell\_ref, file\_text, resource)

Cell\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cell note
containing sound that you want to play. If cell\_ref is omitted,
Microsoft Excel plays the sound from the active cell, or from a file if
you specify one.

File\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a file containing
sounds. If cell\_ref is specified, file\_text is ignored.

Resource&nbsp;&nbsp;&nbsp;&nbsp;is a number or name given as text
specifying a sound resource in file\_text that you want to play.

  - > This argument applies only to Microsoft Excel for the Macintosh.

  - > If cell\_ref is specified, resource is ignored.

  - > If resource is omitted, Microsoft Excel uses the first sound
    > resource in the file.

  - > If the file does not contain a sound resource with the specified
    > name or number, Microsoft Excel halts the macro and displays an
    > error message.

> &nbsp;

**Related Function**

SOUND.NOTE&nbsp;&nbsp;&nbsp;Records or imports sound into or erases
sound from cell notes

Return to [top](#Q)

# SPELLING

Equivalent to clicking the Spelling command on the Tools menu. Checks
the spelling of words in the current selection.

**Syntax**

**SPELLING**(custom\_dic, ignore\_uppercase, always\_suggest)

Custom\_dic&nbsp;&nbsp;&nbsp;&nbsp;is the filename of the custom
dictionary to examine if words are not found in the main dictionary. If
custom\_dic is omitted, the currently specified dictionary is used.

Ignore\_uppercase&nbsp;&nbsp;&nbsp;&nbsp;is a logical value
corresponding to the Ignore UPPERCASE check box.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>If ignore_uppercase is</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Microsoft Excel will</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>TRUE</p>
</blockquote></td>
<td><blockquote>
<p>Ignore words in all uppercase letters</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>FALSE</p>
</blockquote></td>
<td><blockquote>
<p>Check words in all uppercase letters</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Omitted</p>
</blockquote></td>
<td><blockquote>
<p>Use the current setting</p>
</blockquote></td>
</tr>
</tbody>
</table>

Always\_suggest&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding
to the Always Suggest check box.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>If always_suggest is</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Microsoft Excel will</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>TRUE</p>
</blockquote></td>
<td><blockquote>
<p>Display a list of suggested alternate spellings when an incorrect spelling is found</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>FALSE</p>
</blockquote></td>
<td><blockquote>
<p>Wait for user to input the correct spelling</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>Omitted</p>
</blockquote></td>
<td><blockquote>
<p>Use the current setting</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Related Function**

SPELLING.CHECK&nbsp;&nbsp;&nbsp;Checks the spelling of a word

Return to [top](#Q)

# SPELLING.CHECK

Checks the spelling of a word. Returns TRUE if the word is spelled
correctly; FALSE otherwise.

**Syntax**

**SPELLING.CHECK**(**word\_text**, custom\_dic, ignore\_uppercase)

Word\_text&nbsp;&nbsp;&nbsp;&nbsp;is the word whose spelling you want to
check. It can be text or a reference to text.

Custom\_dic&nbsp;&nbsp;&nbsp;&nbsp;is the filename of a custom
dictionary to examine if the word is not found in the main dictionary.

Ignore\_uppercase&nbsp;&nbsp;&nbsp;&nbsp;is a logical value
corresponding to the Ignore Words In Uppercase check box. If
ignore\_uppercase is TRUE, the check box is selected, and Microsoft
Excel ignores words in all uppercase letters; if FALSE, the check box is
cleared, and Microsoft Excel checks all words; if omitted, the current
setting is used.

**Remarks**

This function does not have a dialog-box form. To display the Spelling
dialog box, use SPELLING.

**Related Function**

SPELLING&nbsp;&nbsp;&nbsp;Checks the spelling of words in the current
selection

Return to [top](#Q)

# SPLIT

Equivalent to choosing the Split command from the Window menu or to
dragging the split bar in the active window's scroll bar. Splits the
active window into panes. Use SPLIT when you want to view different
parts of the active sheet at the same time.

**Syntax**

**SPLIT**(col\_split, row\_split)

Col\_split&nbsp;&nbsp;&nbsp;&nbsp;specifies where to split the window
vertically and is measured in columns from the left of the window.

Row\_split&nbsp;&nbsp;&nbsp;&nbsp;specifies where to split the window
horizontally and is measured in rows from the top of the window.

If an argument is 0 and there is a split in that direction, the split is
removed. If an argument is omitted, a split in that direction is not
changed.

**Related Function**

FREEZE.PANES&nbsp;&nbsp;&nbsp;Freezes or unfreezes the panes of a window

Return to [top](#Q)

# SQL.BIND

Specifies where results from a SQL query are placed when they are
retrieved with SQL.RETRIEVE. If this function is not available, you must
install the Microsoft ODBC add-in (XLODBC.XLA).

**Syntax**

**SQL.BIND**(**connection\_num**, column, reference)

Connection\_num&nbsp;&nbsp;&nbsp;&nbsp;is the unique connection ID of
the data source for which you want to define storage.

  - > Connection\_num was returned by a previously executed SQL.OPEN
    > function.

  - > If connection\_num is not valid, then SQL.BIND returns the
    > \#VALUE\! error value.

> &nbsp;

Column&nbsp;&nbsp;&nbsp;&nbsp;is the number of the result column that is
to be bound. Result columns are numbered from left to right starting
with 1. If column is omitted then all bindings for connection\_num are
removed. Column number 0 contains row numbers for the result set. If
column number 0 is bound then SQL.RETRIEVE will return row numbers in
the bound location.

Reference&nbsp;&nbsp;&nbsp;&nbsp; is a single cell reference on the
worksheet where the results should be placed. If reference is omitted,
the binding is removed for the column.

When SQL.RETRIEVE is called, the result rows in this column will be
placed in the reference cell and the cells immediately below reference.
The number of rows that will be retrieved is one of the SQL.RETRIEVE
arguments.

**Remarks**

  - > If SQL.BIND is completed successfully then it will return a
    > vertical array listing the bound columns on the current
    > connection. If SQL.BIND is unable to bind the result column then
    > it will return the error value \#N/A. In such a case SQL.BIND will
    > place error information in memory for the SQL.ERROR function, if
    > such information is available.

  - > SQL.BIND tells the ODBC interface where to place results when they
    > are retrieved using SQL.RETRIEVE. Binding is not necessary but can
    > be useful if you want the results from different columns to be
    > placed in disjoint worksheet locations.

  - > If bindings are used, SQL.BIND must be called once for each column
    > in the result set. If a result column is not bound then it will
    > not be returned. A binding remains valid for as long as
    > connection\_num is open.

  - > Call SQL.BIND after calling SQL.OPEN and SQL.EXEC.QUERY, but
    > before calling SQL.RETRIEVE or SQL.RETRIEVE.TO.FILE. Calls to
    > SQL.BIND will not affect results that have already been retrieved.

**Example**

SQL.BIND(conn1,1,"\[Book1\]Sheet1\!C1") stores data obtained from the
data source conn1 on Sheet1 from left to right in cell C1, starting with
column1.

**Related Functions**

SQL.OPEN&nbsp;&nbsp;&nbsp;Establishes a connection with a data source

SQL.EXEC.QUERY&nbsp;&nbsp;&nbsp;Sends a query to a data source

SQL.RETRIEVE.TO.FILE&nbsp;&nbsp;&nbsp;Retrieves query results and places
them in a file

SQL.RETRIEVE&nbsp;&nbsp;&nbsp;Retrieves query results

SQL.GET.SCHEMA&nbsp;&nbsp;&nbsp;Gets information about a connected data
source.

SQL.CLOSE&nbsp;&nbsp;&nbsp;Closes a connection to a data source.

SQL.ERROR&nbsp;&nbsp;&nbsp;Returns detailed error information

Return to [top](#Q)

# SQL.CLOSE

Terminates a connection to an external data source. If this function is
not available, you must install the Microsoft ODBC add-in (XLODBC.XLA).

**Syntax**

**SQL.CLOSE**(**connection\_num**)

Connection\_num&nbsp;&nbsp;&nbsp;&nbsp;is the unique connection ID of
the data source from which you wish to disconnect.

  - > Connection\_num is returned by a previously executed SQL.OPEN
    > function.

  - > If connection\_num is not valid, SQL.CLOSE returns the \#VALUE\!
    > error value.

> &nbsp;

**Remarks**

  - > If the connection is successfully terminated SQL.CLOSE will return
    > zero and the connection ID number is then no longer valid.

  - > If SQL.CLOSE is unable to disconnect with the data source then it
    > will return the error value the \#N/A error value. In such a case
    > SQL.CLOSE will place error information in memory for the SQL.ERROR
    > function, if such information is available.

  - > SQL.CLOSE works with data sources in much the same manner as
    > FCLOSE works with files. If the call is successful then SQL.CLOSE
    > will terminate the specified data source connection.

**Example**

SQL.CLOSE(conn1) will close the connection with connection\_num conn1

**Related Functions**

SQL.OPEN&nbsp;&nbsp;&nbsp;Establishes a connection with a data source

SQL.EXEC.QUERY&nbsp;&nbsp;&nbsp;Sends a query to a data source

SQL.BIND&nbsp;&nbsp;&nbsp;Specifies storage for a result column

SQL.RETRIEVE.TO.FILE&nbsp;&nbsp;&nbsp;Retrieves query results and places
them in a file

SQL.RETRIEVE&nbsp;&nbsp;&nbsp;Retrieves query results

SQL.GET.SCHEMA&nbsp;&nbsp;&nbsp;Gets information about a connected data
source.

SQL.ERROR&nbsp;&nbsp;&nbsp;Returns detailed error information

Return to [top](#Q)

# SQL.ERROR

Returns detailed error information when it is called after a previous
XLODBC.XLA function call has failed. If this function is not available,
you must install the Microsoft ODBC add-in (XLODBC.XLA).

**Syntax**

**SQL.ERROR**()

Calling SQL.ERROR returns detailed error information in a two
dimensional array. Each row in the array describes exactly one error. If
a function call generates multiple errors, a row will be created for
each error. When SQL.ERROR is processed successfully, all SQL.ERROR
information is cleared. Also, all SQL.ERROR information is automatically
removed whenever an ODBC function completes successfully.

Each row will have exactly three fields. The information in these three
fields is obtained through the SQLERROR API function call. These fields
are:

  - > A textual message describing the error.

  - > The ODBC error class and subclass as a character string.

  - > The data source native error code as a numeric value.

> &nbsp;

If one or more of these fields is not available for the type of error
that was encountered, the field will be left blank. For more information
on the meaning of these three fields, refer to Chapter 24, "ODBC
Function Reference", in the Microsoft Open Database Connectivity
Programmer's Reference for the SQLError API function. See also Appendix
A, "ODBC Error Codes" in the same manual.

**Remarks**

  - > SQL.ERROR cannot provide information on Excel errors.

  - > If no error information is available when SQL.ERROR is called,
    > then it well return the error value \#N/A but does not post any
    > error information to SQL.ERROR.

  - > SQL.ERROR stores and returns error information by processing
    > SQL.ERROR (in the ODBC API reference) in a loop until
    > SQL\_NO\_DATA\_FOUND is encountered. In the SQL.ERROR function,
    > the error information is automatically defined and stored in
    > memory whenever an XLODBC.XLA function fails. If the call is
    > successful then SQL.ERROR will return the error information
    > available. If SQL.ERROR fails, it will return the error value
    > \#N/A.

**Example**

When entered as an array formula, the following example will return
error information about each argument in, for example, SQL.GET.QUERY.

SQL.ERROR()

**Related Functions**

SQL.OPEN&nbsp;&nbsp;&nbsp;Establishes a connection with a data source

SQL.EXEC.QUERY&nbsp;&nbsp;&nbsp;Sends a query to a data source

SQL.BIND&nbsp;&nbsp;&nbsp;Specifies storage for a result column

SQL.RETRIEVE.TO.FILE&nbsp;&nbsp;&nbsp;Retrieves query results and places
them in a file

SQL.RETRIEVE&nbsp;&nbsp;&nbsp;Retrieves query results

SQL.CLOSE&nbsp;&nbsp;&nbsp;Closes a data source connection

Return to [top](#Q)

# SQL.EXEC.QUERY

Sends a query to a data source using an existing connection. If this
function is not available, you must install the Microsoft ODBC add-in
(XLODBC.XLA).

**Syntax**

**SQL.EXEC.QUERY**(**connection\_num**, **query\_text**)

Connection\_num&nbsp;&nbsp;&nbsp;&nbsp;is the unique connection ID of
the data source you want to query.

  - > Connection\_num is returned by a previously executed SQL.OPEN
    > function.

  - > If connection\_num is not valid, SQL.EXEC.QUERY returns the
    > \#VALUE\! error value.

Query\_text is the SQL language query that is to be executed on the data
source. The query must follow the SQL syntax guidelines in the Appendix
of the Microsoft Excel ODBC Developers Guide.

  - > If SQL.EXEC.QUERY is unable to execute query\_text on the
    > specified data source, SQL.EXEC.QUERY returns the \#N/A error
    > value.

  - > Excel limits strings to a length of 255 characters. If query\_text
    > needs to be longer than 255 characters then query\_text should be
    > a vertical array or vertical range of cells. The values in the
    > array will be joined together to form the complete SQL query.

> &nbsp;

**Remarks**

  - > Before calling SQL.EXEC.QUERY a connection must be established
    > with a data source using SQL.OPEN. A successful call to SQL.OPEN
    > returns a unique connection ID number. SQL.EXEC.QUERY uses that
    > connection ID number to send SQL language queries to the data
    > source.

  - > Any results generated from the query will not be returned
    > immediately-- SQL.EXEC.QUERY only executes the query. Retrieving
    > results is handled by the functions SQL.RETRIEVE and
    > SQL.RETRIEVE.TO.FILE.

  - > If SQL.EXEC.QUERY is called using a previously used connection ID
    > number, all pending results on that connection will automatically
    > be discarded. The connection ID will then refer to the new query
    > and its results.

  - > If SQL.EXEC.QUERY is unable to successfully execute the query on
    > the specified data source then an error value will be returned. In
    > such a case SQL.EXEC.QUERY will place error information in memory
    > for the SQL.ERROR function, if such information is available. If
    > SQL.EXEC.QUERY is able to successfully execute the query on the
    > specified connection it will return one of three values depending
    > on the type of SQL statement that was executed.
    
      - > If it was a SELECT statement then SQL.EXEC.QUERY will return
        > the number of result columns available.
    
      - > If it was an UPDATE, INSERT, or DELETE statement then
        > SQL.EXEC.QUERY will return the number of rows affected by the
        > statement.
    
      - > If it was a legal SQL query that is not in one of the
        > categories above, SQL.EXEC.QUERY will return 0 (zero).

**Example**

SQL.EXEC.QUERY(conn1, "SELECT Custmr\_ID, Due\_Date FROM Orders WHERE
Order\_Amt \> 100") executes a SQL query from a SQL table named "Orders"

**Related Functions**

SQL.OPEN&nbsp;&nbsp;&nbsp;Establishes a connection with a data source

SQL.BIND&nbsp;&nbsp;&nbsp;Specifies storage for a result column

SQL.RETRIEVE.TO.FILE&nbsp;&nbsp;&nbsp;Retrieves query results and places
them in a file

SQL.RETRIEVE&nbsp;&nbsp;&nbsp;Retrieves query results

SQL.GET.SCHEMA&nbsp;&nbsp;&nbsp;Gets information about a connected data
source.

SQL.CLOSE&nbsp;&nbsp;&nbsp;Closes a data source connection

SQL.ERROR&nbsp;&nbsp;&nbsp;Returns detailed error information

Return to [top](#Q)

# SQL.GET.SCHEMA

Returns information about the structure of the data source on a
particular connection. The return value from a successful call to
SQL.GET.SCHEMA depends on the type of information that was requested. A
list of the accepted requests and their respective return values is
listed in the syntax section below.

If this function is not available, you must install the Microsoft ODBC
add-in (XLODBC.XLA).

**Syntax**

**SQL.GET.SCHEMA**(**connection\_num**, **type\_num**, qualifier\_text)

Connection\_num&nbsp;&nbsp;&nbsp;&nbsp;is the unique connection ID of
the data source you want information about.

  - > Connection\_num is returned by a previously executed SQL.OPEN
    > function.

  - > If connection\_num is not valid, SQL.GET.SCHEMA returns the
    > \#VALUE\! error value.

> &nbsp;

Type\_num specifies the type of information you want returned. The
following is a list of valid type\_num values.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>A list of available data sources, as a vertical array.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A list of databases on the current connection, as a vertical array .</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>A list of owners in a database on the current connection, as a vertical array.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>A list of tables for a given owner and database on the current connection, as a vertical array.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>A list of columns in a particular table and their data types, as a two-dimensional array. The returned array will have two fields and will have a row for each column in the table. The first field will be the name of the column. The second field is the data type of the column. The data type will be a number that corresponds to the ODBC C header file data types. These #define numbers are found in Microsoft Excel ODBC Developer Guide.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>User ID of the current user</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>Name of the current database.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>The name of the data source as given in the ODBC.INI file.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>The name of the data source DBMS (i.e. Oracle, SQL Server, etc.).</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>The server name for the data source.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>The terminology used by the data source to refer to owners ( i.e. "owner", "Authorization ID", "Schema", etc.).</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>The terminology used by the data source to refer to tables ( i.e. "table", "file", etc.).</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>The terminology used by the data source to refer to qualifiers (i.e. "database" or "directory").</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>The terminology used by the data source to refer to procedures (i.e. "database procedure", "stored procedure", or "procedure").</p>
</blockquote></td>
</tr>
</tbody>
</table>

Qualifier\_text&nbsp;&nbsp;&nbsp;&nbsp;is only included for type\_num
values of 3, 4 and 5. It is a text string used to qualify the search for
the requested information and should be enclosed by quotation marks.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Qualifier_text</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>The value of qualifier_text should be the name of a database in the current data source. SQL.GET.SCHEMA will then only return the names of table owners in that database.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The value of qualifier_text should be both a database name and an owner name. The syntax of qualifier_text is "DatabaseName.OwnerName". A period is used to separate the two names. SQL.GET.SCHEMA will then return an array of table names that are located in the given database and owned by the given owner.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>The value of qualifier_text should be the name of a table. Information about the columns in that table will be returned.</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Remarks**

  - > If SQL.GET.SCHEMA is unable to find the requested information then
    > it will return the error value \#N/A. In such a case
    > SQL.GET.SCHEMA will place error information in memory for the
    > SQL.ERROR function, if such information is available.

  - > SQL.GET.SCHEMA works with the ODBC functions SQLGetInfo and
    > SQLTables to find the requested information. Refer to the
    > Microsoft Excel ODBC Programmer Guide for more information on
    > these two functions.

**Example**

SQL.GET.SCHEMA(conn1,7) returns the name of the current database.

SQL.GET.SCHEMA(conn1,9) returns the name of the DBMS used by the data
source.

**Related Functions**

SQL.OPEN&nbsp;&nbsp;&nbsp;Establishes a connection with a data source

SQL.EXEC.QUERY&nbsp;&nbsp;&nbsp;Sends a query to a data source

SQL.BIND&nbsp;&nbsp;&nbsp;Specifies storage for a result column

SQL.RETRIEVE.TO.FILE&nbsp;&nbsp;&nbsp;Retrieves query results and places
them in a file

SQL.RETRIEVE&nbsp;&nbsp;&nbsp;Retrieves query results

SQL.CLOSE&nbsp;&nbsp;&nbsp;Closes a data source connection

SQL.ERROR&nbsp;&nbsp;&nbsp;Returns detailed error information

Return to [top](#Q)

# SQL.OPEN

Establishes a connection with a data source. If the connection is
successfully established SQL.OPEN will return a connection ID number.
Use the connection ID number with other ODBC macro functions to identify
a connection.

If this function is not available, you must install the Microsoft ODBC
add-in (XLODBC.XLA).

**Syntax**

**SQL.OPEN**(**connection\_string**, output\_ref, driver\_prompt)

Connection\_string&nbsp;&nbsp;&nbsp;&nbsp;is a text string that contains
the information necessary to establish a connection to a data source.
Any data-source-name that is used in connection\_string must be an
existing data source name defined with ODBC Setup or the ODBC
Administration Utility.

  - > Connection\_string must follow the format described in Chapter 24,
    > "ODBC Function Reference", of the Microsoft Open Database
    > Connectivity Programmer's Reference for SQLDriverConnect. In this
    > string the user supplies the data source name, one or more user
    > ID's, one or more passwords, and any other information necessary
    > to successfully connect to a DBMS. An example of a SQL.OPEN
    > connection\_string entered would be: "DSN=MyServer; UID=dbayer;
    > PWD=123; Database=pubs"

  - > Enter the connection\_string as an array when the length exceeds
    > 255 characters. Or enter connection\_string as an array of cells
    > containing the same information. The connection string should be
    > horizontal array.

Output\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a single cell reference where you
want the completed connection string to be placed. Use output\_ref when
you want SQL.OPEN to return the completed connection string. If
output\_ref is omitted, a completed connection string will not be
returned.

Driver\_prompt&nbsp;&nbsp;&nbsp;&nbsp; is a number from 1 to 4
specifying if and how you want to be prompted by the driver. This sets
the fDriverCompletion flag in ODBC's SQLDriverConnect.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Number</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Description</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Always brings up a dialog box.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Bring up dialog only if there is not enough information to connect. The driver uses information from the connection string and from the data source specification as defaults.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Same as 2, but the driver grays and disables any prompts for information not needed.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>If the connection string is unsuccessful, do not bring up a dialog box.</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Remarks**

  - > If SQL.OPEN is unable to connect with the information provided
    > then it will return the error value \#N/A. In such a case,
    > SQL.OPEN will place error information in memory for the SQL.ERROR
    > function, if more information is available.

  - > If the call is successful then SQL.OPEN will return a unique
    > connection ID number that can be used in future function calls to
    > identify the connection.

  - > If connection\_array does not allow SQL.OPEN to connect to a data
    > source, then the error value \#N/A will be returned.

**Example**

conn1=SQL.OPEN('DSN=NWind;DBQ=C:\\MSQUERY;FIL=dBASE4',C15, 2) sets the
name conn1 to the return value of SQL.OPEN, which connects to the NWind
data source, specifies where to place the connection string, and
displays the driver dialog box only if additional information is needed.

**Related Functions**

SQL.EXEC.QUERY&nbsp;&nbsp;&nbsp;Sends a query to a data source

SQL.BIND&nbsp;&nbsp;&nbsp;Specifies storage for a result column

SQL.RETRIEVE.TO.FILE&nbsp;&nbsp;&nbsp;Retrieves query results and places
them in a file

SQL.RETRIEVE&nbsp;&nbsp;&nbsp;Retrieves query results

SQL.GET.SCHEMA&nbsp;&nbsp;&nbsp;Gets information about a connected data
source.

SQL.CLOSE&nbsp;&nbsp;&nbsp;Closes a data source connection

SQL.ERROR&nbsp;&nbsp;&nbsp;Returns detailed error information

Return to [top](#Q)

# SQL.RETRIEVE

Retrieves all or part of the results from a previously executed query.
The connection used must have already been established using the macro
function SQL.OPEN. Also, a query must already have been executed using
SQL.EXEC.QUERY and results must be pending.

If this function is not available, you must install the Microsoft ODBC
add-in (XLODBC.XLA).

**Syntax**

**SQL.RETRIEVE**(**connection\_num**, destination\_ref, max\_columns,
max\_rows, col\_names\_logical, row\_nums\_logical, named\_rng\_logical,
fetch\_first\_logical)

Connection\_num&nbsp;&nbsp;&nbsp;&nbsp;is the unique connection ID for a
data source. The data source specified must have pending query results.
Pending query results are generated by a call to SQL.EXEC.QUERY on the
same connection.

  - > If there are no pending results on the connection SQL.RETRIEVE
    > returns the \#N/A error value.

  - > If connection\_num is not valid, SQL.EXEC.QUERY returns the
    > \#VALUE\! error value.

> &nbsp;

Destination\_ref&nbsp;&nbsp;&nbsp;&nbsp;specifies where the results
should be placed. It is either a reference to a single cell or it is
omitted.

  - > If destination\_ref refers to a single cell then SQL.RETRIEVE will
    > return all of the pending results in the cells to the right,
    > below, and including destination\_ref. This is the same convention
    > used in Microsoft Excel when multiple cells are pasted into a
    > single-cell selection. Any previous values contained in the
    > destination cells will be overwritten without confirmation.

  - > If destination\_ref is omitted then the bindings established by
    > previous calls to SQL.BIND will be used to return results. If no
    > such bindings exist for the current connection then SQL.RETRIEVE
    > will return the \#REF\! error value. If a particular result column
    > has not been bound then its results will be discarded. Max\_rows
    > specifies the number of rows that will be returned under each
    > bound column. The first row of results will be placed in the bound
    > cell and any additional rows will be placed in the rows
    > immediately under the bound cell.

> &nbsp;

Max\_columns&nbsp;&nbsp;&nbsp;&nbsp;is the maximum number of columns to
be retrieved. It is only used when destination\_ref is not omitted.

  - > If max\_columns specifies more columns than are available in the
    > results, SQL.RETRIEVE will place data in the columns for which
    > data is available and clear the additional columns.

  - > If max\_columns specifies fewer columns than are available in the
    > results, the rightmost result columns will be discarded to fit the
    > chosen size. Column position will be determined by the order in
    > which the data source returned them.

  - > If max\_columns is omitted then all of the result columns will be
    > returned.

> &nbsp;

Max\_rows&nbsp;&nbsp;&nbsp;&nbsp;is the maximum number of rows to be
returned.

  - > If max\_rows specifies more rows than are available in the
    > results, SQL.RETRIEVE will place data in the rows for which data
    > is available and clear the additional rows.

  - > If max\_rows specifies fewer rows than are available in the
    > results, SQL.RETRIEVE will place data in the selected rows but
    > will not discard the additional rows. These extra rows can be
    > retrieved via additional calls to SQL.RETRIEVE. This process is
    > described in the fetch\_first\_logical argument description.

  - > If max\_rows is omitted then all rows in the result set will be
    > returned.

> &nbsp;

Col\_names\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value which, if
TRUE, causes the column names to be returned as the first row of
results. It FALSE or omitted, the column names will not be returned.

Row\_nums\_logical&nbsp;&nbsp;&nbsp;&nbsp;is used only when
destination\_ref is included. If row\_nums\_logical is TRUE then the
first column in the result set will contain row numbers. If FALSE then
row numbers will not be returned. This column of row numbers will not
have a column name and the column heading will be left blank. Row
numbers can also be retrieved by binding column number 0 with SQL.BIND.

Named\_rng\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value which, if
TRUE, sets each column of results to be declared as a named range on the
worksheet. The name of the each range will be the result column name.
The named range will only include the rows that were fetched with this
SQL.RETRIEVE function call. If FALSE, the results will not be declared
as a named range.

Fetch\_first\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that
allows you to request results from the beginning of the result set.

  - > If the first call to SQL.RETRIEVE did not return all of the rows
    > in the result set then SQL.RETRIEVE may be called again to return
    > the next set of rows. This process can be repeated until no more
    > result rows are available, at which time SQL.RETRIEVE will return
    > the value 0 (zero). This will not halt the running of the macro.
    > During each of these calls, including the first call,
    > fetch\_first\_logical should be set to FALSE.

  - > If you want to move the cursor back to the beginning of the result
    > set then fetch\_first\_logical should be set to TRUE. This causes
    > the same SQL query text to be executed again on the data source.
    > The cursor will then be positioned at the top of the result set
    > and SQL.RETRIEVE will fill destination\_ref beginning with the
    > first row of results. Further calls to SQL.RETRIEVE, for the
    > purpose of retrieving additional rows, can then be made with
    > fetch\_first\_logical set to FALSE .

> &nbsp;

**Remarks**

  - > Before calling SQL.RETRIEVE a connection must be established with
    > a data source using SQL.OPEN.

  - > If SQL.RETRIEVE is unable to retrieve the results on the specified
    > data source then an error value will be returned. In such a case
    > SQL.RETRIEVE will place error information in memory for the
    > SQL.ERROR function, if such information is available.

  - > If SQL.RETRIEVE is able to successfully return rows of results on
    > the specified connection it will return the number of rows that
    > were actually returned. If there were no results pending on the
    > connection then SQL.RETRIEVE will return the \#N/A error value.If
    > no data was found then SQL.RETRIEVE returns 0 (zero).

  - > A successful call to SQL.OPEN returns a unique connection ID
    > number, which is used in a call to SQL.EXEC.QUERY to send a SQL
    > language query. Following this call to SQL.EXEC.QUERY,
    > SQL.RETRIEVE uses the same connection ID number to retrieve query
    > results.

**Example**

SQL.RETRIEVE(conn1,sheet1\!C1,1) stores data obtained from the data
source conn1 on Sheet1 from left to right in cell C1, using only column
1.

**Related Functions**

SQL.OPEN&nbsp;&nbsp;&nbsp;Establishes a connection with a data source

SQL.EXEC.QUERY&nbsp;&nbsp;&nbsp;Sends a query to a data source

SQL.BIND&nbsp;&nbsp;&nbsp;Specifies storage for a result column

SQL.RETRIEVE.TO.FILE&nbsp;&nbsp;&nbsp;Retrieves query results and places
them in a file

SQL.GET.SCHEMA&nbsp;&nbsp;&nbsp;Gets information about a connected data
source.

SQL.CLOSE&nbsp;&nbsp;&nbsp;Close a data source connection

SQL.ERROR&nbsp;&nbsp;&nbsp;Returns detailed error information

Return to [top](#Q)

# SQL.RETRIEVE.TO.FILE

Retrieves all of the results from a previously executed query and places
them in a file. The connection used must have already been established
using the macro function SQL.OPEN. Also, a query must already have been
executed using SQL.EXEC.QUERY and results must be pending.

If this function is not available, you must install the Microsoft ODBC
add-in (XLODBC.XLA).

**Syntax**

**SQL.RETRIEVE.TO.FILE**(**connection\_num**, **destination**,
col\_names\_logical, column\_delimiter)

Connection\_num&nbsp;&nbsp;&nbsp;&nbsp;is the unique connection ID for a
data source. The data source specified must have query results pending.
Pending results were generated by a previous call to SQL.EXEC.QUERY on
the same connection.

  - > If there are no pending results on the connection
    > SQL.RETRIEVE.TO.FILE returns the \#N/A error value. The file is
    > not affected.

  - > If connection\_num is not valid, SQL.RETRIEVE.TO.FILE returns the
    > \#VALUE\! error value.

> &nbsp;

Destination&nbsp;&nbsp;&nbsp;&nbsp;specifies the name and path of the
file where the results should be placed. SQL.RETRIEVE.TO.FILE will open
the specified file and fill it with the entire result set.

  - > The format of the data in the file will be compatible with the
    > Microsoft Excel ".CSV" format. The overall format will be that
    > columns will be separated by the value in column\_delimiter (see
    > below) and the individual rows will be separated by a
    > linefeed/carriage-return.

  - > If the file specified by destination cannot be opened then the
    > error value \#N/A will be returned by SQL.RETRIEVE.TO.FILE.

  - > If the file already exists its previous contents will be
    > overwritten by SQL.RETRIEVE.TO.FILE.

> &nbsp;

Col\_names\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that, if
TRUE, allows the column names to be returned as the first row of data.
If FALSE or omitted, the column names will not be returned.

Column\_delimiter&nbsp;&nbsp;&nbsp;&nbsp;is the value that will be used
to separate the elements in each row. If column\_delimiter is omitted
then a tab will be used. If another value is desired then it should be
enclosed in quotation marks. Possible values for column\_delimiter might
be: "," or ";" or " ". The string "tab" can also be used to specify a
tab separator (even though this is redundant, since a tab is the
default).

**Remarks**

  - > If SQL.RETRIEVE.TO.FILE is unable to retrieve the results on the
    > specified connection then an error value will be returned. In such
    > a case SQL.RETRIEVE.TO.FILE will place error information in memory
    > for the SQL.ERROR function, if such information is available.

  - > If SQL.RETRIEVE.TO.FILE is able to successfully return rows of
    > results on the specified connection and place them in a file it
    > will return the number of rows that were actually written to the
    > file. If there were no results pending on the connection then
    > SQL.RETRIEVE.TO.FILE will return the \#N/A error value and the
    > file will not be created or modified.

  - > Before calling SQL.RETRIEVE.TO.FILE a connection must be
    > established with a data source using SQL.OPEN.

  - > A successful call to SQL.OPEN returns a unique connection ID
    > number, which can be used in a call to SQL.EXEC.QUERY to send a
    > SQL language query. Following this call to SQL.EXEC.QUERY,
    > SQL.RETRIEVE.TO.FILE uses the same connection ID number to
    > retrieve query results and place them in a file.

**Example**

SQL.RETRIEVE.TO.FILE(conn1,"C:\\MSQUERY\\RESULTS1.QRY",TRUE,",")
retrieves the results of a previously executed query and places them in
the file RESULTS1.QRY, with column names that are comma delimited.

**Related Functions**

SQL.OPEN&nbsp;&nbsp;&nbsp;Establishes a connection with a data source

SQL.EXEC.QUERY&nbsp;&nbsp;&nbsp;Sends a query to a data source

SQL.BIND&nbsp;&nbsp;&nbsp;Specifies storage for a result column

SQL.RETRIEVE&nbsp;&nbsp;&nbsp;Retrieves query results

SQL.GET.SCHEMA&nbsp;&nbsp;&nbsp;Gets information about a connected data
source.

SQL.CLOSE&nbsp;&nbsp;&nbsp;Closes a data source connection

SQL.ERROR&nbsp;&nbsp;&nbsp;Returns detailed error information

Return to [top](#Q)

# STANDARD.FONT

Sets the attributes of the standard font in Microsoft Excel version 2.2
and earlier. This function is included only for macro compatibility. To
define and apply a style in Microsoft Excel version 5.0 or later, use
the APPLY.STYLE and DEFINE.STYLE functions.

**Syntax**

**STANDARD.FONT**(name\_text, size\_num, bold, italic, underline,
strike, color, outline, shadow)

The arguments for this function are the same as those for FORMAT.FONT.

**Related Functions**

APPLY.STYLE&nbsp;&nbsp;&nbsp;Applies a style to the selection

DEFINE.STYLE&nbsp;&nbsp;&nbsp;Defines a cell style

FORMAT.FONT&nbsp;&nbsp;&nbsp;Applies a font to the selection

Return to [top](#Q)

# STANDARD.WIDTH

Sets the default width used for all columns that you have not previously
adjusted on the active worksheet.

**Syntax**

**STANDARD.WIDTH**(standard\_num)

Standard\_num&nbsp;&nbsp;&nbsp;&nbsp;specifies how wide you want the
columns to be in units of one character of the font corresponding to the
Normal cell style.

Return to [top](#Q)

# STEP

Stops the normal flow of a macro and calculates it one cell at a time.
Running a macro one cell at a time is called single-stepping and is very
useful when you are debugging a macro. Use the STEP function, instead of
clicking the Step Into button in the Macro dialog box when you want to
start single-stepping at a specific line in a macro. The Macro dialog
box appears when you click the Macros command (Tools menu, Macro
submenu).

**Syntax**

**STEP**( )

**Remarks**

  - > When Microsoft Excel encounters a STEP function, it stops running
    > the macro and displays a dialog box. The dialog box tells you
    > which cell in the macro Microsoft Excel is about to calculate, and
    > what formula is in that cell. You can click Step to carry out the
    > next instruction; click Evaluate to calculate part of the formula;
    > click Halt to interrupt the macro; or click Continue to continue
    > the macro without single-stepping.

  - > When placed at the beginning of a macro, STEP is equivalent to
    > clicking the Macro command on the Tools menu and clicking the Step
    > Into button in the Macro dialog box.

  - > To step through the calculation of a custom function, place the
    > STEP function at the start of the custom function.

> &nbsp;

**Related Functions**

HALT&nbsp;&nbsp;&nbsp;Stops all macros from running

RUN&nbsp;&nbsp;&nbsp;Runs a macro

Return to [top](#Q)

# STYLE

Checks the fonts for a bold and/or italic font and applies it to the
current selection in Microsoft Excel for the Macintosh version 1.5 or
earlier. If no appropriate font is available, Microsoft Excel finds the
most similar font available and formats the selection using that font.
This function is included only for macro compatibility. If you want to
change a font to bold or italic, use the FONT.PROPERTIES function.

**Syntax**

**STYLE**(bold,italic)

**STYLE**?(bold,italic)

**Related Function**

FONT.PROPERTIES&nbsp;&nbsp;&nbsp;Applies a font to the selection

Return to [top](#Q)

# SUBSCRIBE.TO

Inserts the contents of the edition into the active sheet at the point
of the current selection. Use SUBSCRIBE.TO to incorporate editions
published from other workbooks into your Microsoft Excel worksheets and
macro sheets. SUBSCRIBE.TO returns TRUE if successful.

**Syntax**

**SUBSCRIBE.TO**(**file\_text, format\_num**)

**Important&nbsp;&nbsp;&nbsp;**This function is only available if you
are using Microsoft Excel for the Macintosh with system software version
7.0 or later.

File\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name, as a text string, of the
edition you want to insert into the active sheet. Unless file\_text is
in the current folder, supply the full path of the workbook. If
file\_text cannot be found, SUBSCRIBE.TO returns the \#VALUE\! error
value and interrupts the macro.

**Remarks**

  - > If a single cell is selected, the data from the edition file is
    > placed into as large a range of cells as is required by the data.
    > Data already present in those cells is replaced. If the data is a
    > picture, it is inserted from the upper-left corner of the selected
    > cell.

  - > If a range of cells is selected, and the range is not big enough
    > to contain the edition data, Microsoft Excel displays a dialog box
    > asking if you want to clip the data to fit the range.

> &nbsp;

Format\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and specifies
the format type of the file you are subscribing to.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Format_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Format type</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1 or omitted</p>
</blockquote></td>
<td><blockquote>
<p>Picture</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Text (includes BIFF, VALU, TEXT, and CSV formats)</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Related Functions**

CREATE.PUBLISHER&nbsp;&nbsp;&nbsp;Creates a publisher from the selection

EDITION.OPTIONS&nbsp;&nbsp;&nbsp;Sets publisher and subscriber options

GET.LINK.INFO&nbsp;&nbsp;&nbsp;Returns information about a link

Return to [top](#Q)

# SUBTOTAL.CREATE

Equivalent to clicking the Subtotals command on the Data menu. Generates
a subtotal in a list or database.

**Syntax**

**SUBTOTAL.CREATE**(**at\_change\_in**, **function\_num**, **total**,
replace, pagebreaks, summary\_below)

**SUBTOTAL.CREATE**?(at\_change\_in, function\_num, total, replace,
pagebreaks, summary\_below)

At\_change\_in&nbsp;&nbsp;&nbsp;&nbsp;is a column offset corresponding
to the At Each Change In text box on the Subtotal dialog box.

Function\_Num&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the
Use Function list box specifying which function you want to use in
subtotaling your data.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Function</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Function_Num</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>SUM</p>
</blockquote></td>
<td><blockquote>
<p>1</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>COUNTA</p>
</blockquote></td>
<td><blockquote>
<p>2</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>AVERAGE</p>
</blockquote></td>
<td><blockquote>
<p>3</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>MAX</p>
</blockquote></td>
<td><blockquote>
<p>4</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>MIN</p>
</blockquote></td>
<td><blockquote>
<p>5</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>PRODUCT</p>
</blockquote></td>
<td><blockquote>
<p>6</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>COUNT</p>
</blockquote></td>
<td><blockquote>
<p>7</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>STDEV</p>
</blockquote></td>
<td><blockquote>
<p>8</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>STDEVP</p>
</blockquote></td>
<td><blockquote>
<p>9</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>VAR</p>
</blockquote></td>
<td><blockquote>
<p>10</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>VARP</p>
</blockquote></td>
<td><blockquote>
<p>11</p>
</blockquote></td>
</tr>
</tbody>
</table>

Total&nbsp;&nbsp;&nbsp;&nbsp;is an array of column offsets corresponding
to the Add Subtotal To list box. Indicates which columns you want
aggregated according to function\_num; for example, {4,5}

Replace&nbsp;&nbsp;&nbsp;&nbsp;is a logical value which, if TRUE, causes
any previous subtotals to be replaced by new subtotals. If FALSE or
omitted, subtotals will not be replaced.

PageBreaks&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Page Break Between Groups check box which, if TRUE, creates a page
break below each subtotal. If FALSE or omitted, does not create a page
break below each subtotal.

Summary\_Below&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding
to the Summary Below Data check box which, if TRUE, places subtotal rows
below the data they refer to, and a grand total at the bottom of the
list. If FALSE, places subtotal rows above the data they refer to, and a
grand total just below the header.

**Related Function**

SUBTOTAL.REMOVE&nbsp;&nbsp;&nbsp;Removes all previously existing
subtotals and grand totals in a list

Return to [top](#Q)

# SUBTOTAL.REMOVE

Equivalent to clicking the Subtotal command on the Data menu, and then
clicking the Remove All button in the Subtotal dialog box. Removes all
previously existing subtotals and grand totals in a list. Any page
breaks and outlines will also be removed.

**Syntax**

**SUBTOTAL.REMOVE**()

**Related Function**

SUBTOTAL.CREATE&nbsp;&nbsp;&nbsp;Generates a subtotal in a list or
database

Return to [top](#Q)

# SUMMARY.INFO

Equivalent to clicking the Properties command on the File menu.
Generates the summary information for the active workbook.

**Syntax**

**SUMMARY.INFO**(title, subject, author, keywords, comments)

**SUMMARY.INFO**?(title, subject, author, keywords, comments)

The arguments correspond to the text boxes on the Summary tab of the
Properties dialog box. If any arguments are omitted, that text box is
left empty.

Title&nbsp;&nbsp;&nbsp;&nbsp;specifies a title for the file, not
necessarily a file name. Long names can be entered, up to 255
characters.

Subject&nbsp;&nbsp;&nbsp;&nbsp; is information pertaining to the subject
matter of the workbook.

Author&nbsp;&nbsp;&nbsp;&nbsp; is initially the name specified in the
User Name box on the General tab in the Options dialog box, which
appears from the Options command from the Tools menu. If this is
omitted, the registered user of the copy of Microsoft Excel will be
used.

Keywords&nbsp;&nbsp;&nbsp;&nbsp;are keywords that can be later used in
searching for the contents in the file.

Comments&nbsp;&nbsp;&nbsp;&nbsp; is a comment that can be entered to
help a user learn more about the contents or subject matter of the
workbook.

**Related Functions**

FIND.FILE&nbsp;&nbsp;&nbsp;Lets you search for files based on criteria
such as author or creation date

GET.WORKBOOK&nbsp;&nbsp;&nbsp;Returns information about a workbook sheet

Return to [top](#Q)
