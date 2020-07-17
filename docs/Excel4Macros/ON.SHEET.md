ON.SHEET

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

[ON.WINDOW](ON.WINDOW.md)   Runs a specified macro when you switch to a particular
window.



Return to [README.md](README.md)

ON.SHEET

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

[ON.WINDOW](ON.WINDOW.md)   Runs a specified macro when you switch to a particular
window.



Return to [README.md](README.md)

