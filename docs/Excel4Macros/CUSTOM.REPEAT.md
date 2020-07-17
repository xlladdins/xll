CUSTOM.REPEAT
=============

Allows custom commands to be repeated using the Repeat tool or the
Repeat command on the Edit menu. Also allows custom commands to be
recorded using the macro recorder.

**Syntax**

**CUSTOM.REPEAT**(macro\_text, repeat\_text, record\_text)

Macro\_text    is the name of, or a reference to, the macro you want to
run when the Repeat command is chosen. If macro\_text is omitted, no
repeat macro is run, but the custom command can still be recorded.

Repeat\_text    is the text you want to use as the repeat command on the
Edit menu (for example, \"Repeat Reports\"). You can omit repeat\_text
and macro\_text if you only want to record the formula specified by
record\_text when using the macro recorder.

Record\_text    is the formula you want to record. For example, if the
user clicks a command named Run Reports in Macro 1, the record\_text
argument would be \"=Macro1!RunReports()\", where RunReports is the name
of the macro called by the Run Reports command.

-   References in record\_text must be in R1C1 format.

-   If record\_text is omitted, the macro recorder records normally (a
    > RUN function with the first cell of the macro as its argument).

-   If you are not recording a macro, record\_text is ignored.

>  

**Tip**   Place CUSTOM.REPEAT at the end of the macro you will want to
repeat. If you place it before the end, then the macro formulas that
follow CUSTOM.REPEAT may interfere with the desired effects of
CUSTOM.REPEAT. The Repeat tool and the Repeat command continue to change
as you click subsequent commands that can be repeated.

**Example**

The following macro formula specifies that the macro RepeatReport on the
MenuMacros macro sheet in the current workbook will be run when the
Repeat Report command is chosen:

CUSTOM.REPEAT(\"MenuMacros!RepeatReport\", \"Repeat Report\")

**Related Function**

CUSTOM.UNDO   Specifies a macro to run to undo a custom command

Return to [top](#A)

CUSTOM.UNDO
