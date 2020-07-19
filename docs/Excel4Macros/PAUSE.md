# PAUSE

Pauses a macro. Use the PAUSE function, instead of clicking the Pause
button in the Single Step dialog box, as a debugging tool when you do
not wish to step through a macro. You can also use PAUSE to enter and
edit data, to work directly with Microsoft Excel commands, or to perform
other actions that are not normally available when a macro is running.

**Syntax**

**PAUSE**(no\_tool)

No\_tool&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether to
display the Resume Macro button when the macro is paused. If no\_tool is
TRUE, the toolbar is not displayed; if FALSE, the toolbar is displayed;
if omitted, the toolbar is displayed unless you previously clicked the
close box on the toolbar.

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

**Tip**&nbsp;&nbsp;&nbsp;Since the automatic Resume Macro button can be
customized, you can create a custom toolbar that will appear whenever a
macro pauses.

**Example**

The following macro formula checks to see if a variable named TestValue
is greater than 9. If it is, the macro pauses; otherwise, the macro
continues normally.

IF(TestValue\>9, PAUSE())

**Related Functions**

[HALT](HALT.md)&nbsp;&nbsp;&nbsp;Stops all macros from running

[RESUME](RESUME.md)&nbsp;&nbsp;&nbsp;Resumes a paused macro

[STEP](STEP.md)&nbsp;&nbsp;&nbsp;Turns on macro single-stepping



Return to [README](README.md)

