ON.WINDOW

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

[GET.WINDOW](GET.WINDOW.md)   Returns information about a window

[ON.KEY](ON.KEY.md)   Runs a macro when a specified key is pressed

[ON.SHEET](ON.SHEET.md)   Triggers a macro whenever the specified sheet is activated
from another sheet

[WINDOWS](WINDOWS.md)   Returns the names of all open windows



Return to [README](README.md)

