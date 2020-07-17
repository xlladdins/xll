PUSHBUTTON.PROPERTIES

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

