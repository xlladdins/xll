RESUME
======

Equivalent to choosing the Resume button on the toolbar. Resumes a
paused macro. Returns TRUE if successful or the \#VALUE! error value if
no macro is paused. A macro can be paused by using the PAUSE function or
choosing Pause from the Single Step dialog box, which appears when you
choose the Step Into button from the Macro dialog box.

**Syntax**

**RESUME**(type\_num)

Type\_num    is a number from 1 to 4 specifying how to resume.

+-----------------+---------------------------------------------------+
| > **Type\_num** | > **How Microsoft Excel resumes**                 |
+-----------------+---------------------------------------------------+
| > 1 or omitted  | > If paused by a PAUSE function, continues        |
|                 | > running the macro. If paused from the Single    |
|                 | > Step dialog box, returns to that dialog box.    |
+-----------------+---------------------------------------------------+
| > 2             | > Halts the paused macro                          |
+-----------------+---------------------------------------------------+
| > 3             | > Continues running the macro                     |
+-----------------+---------------------------------------------------+
| > 4             | > Opens the Single Step dialog box                |
+-----------------+---------------------------------------------------+

**Tip**   You can use Microsoft Excel\'s ON functions to resume based on
an event. For an example, see ENTER.DATA.

**Remarks**

-   If one macro runs a second macro that pauses, and you need to halt
    > only the paused macro, use RESUME(2) instead of HALT. HALT halts
    > all macros and prevents resuming or returning to any macro.

-   If the macro was paused from the Single Step dialog box, RESUME
    > returns to the Single Step dialog box.

>  

**Related Functions**

HALT   Stops all macros from running

PAUSE   Pauses a macro

RETURN   Ends the currently running macro

Return to [top](#Q)

RETURN
