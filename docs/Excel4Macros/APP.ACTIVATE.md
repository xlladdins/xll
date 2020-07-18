APP.ACTIVATE

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

[T](T.md)he first five functions following are only for Microsoft Excel for
[W](W.md)indows.

[APP.MAXIMIZE](APP.MAXIMIZE.md)&nbsp;&nbsp;&nbsp;Maximizes the Microsoft Excel application
window

[APP.MINIMIZE](APP.MINIMIZE.md)&nbsp;&nbsp;&nbsp;Minimizes the Microsoft Excel application
window

[APP.MOVE](APP.MOVE.md)&nbsp;&nbsp;&nbsp;Moves the Microsoft Excel application window

[APP.RESTORE](APP.RESTORE.md)&nbsp;&nbsp;&nbsp;Restores the Microsoft Excel application
window

[APP.SIZE](APP.SIZE.md)&nbsp;&nbsp;&nbsp;Changes the size of the Microsoft Excel
application window

[EXEC](EXEC.md)&nbsp;&nbsp;&nbsp;Starts another application



Return to [README](README.md)

