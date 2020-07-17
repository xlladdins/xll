ON.DATA

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


