# CUSTOM.UNDO

Creates a customized Undo tool and Undo or Redo command on the Edit menu
for custom commands.

**Syntax**

**CUSTOM.UNDO**(**macro\_text**, undo\_text)

Macro\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of, or an R1C1-style
reference to, the macro you want to run when the Undo command is chosen.
Macro\_text can be the name or cell reference of a macro.

Undo\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text you want to use as the
Undo command.

**Example**

The following macro function runs the UndoMult macro when the user
clicks the Undo Times100 command, a custom command that multiples the
current cell by 100.

\=CUSTOM.UNDO("UndoMult", "\&Undo Times100")

**Tip**&nbsp;&nbsp;&nbsp;Use CUSTOM.UNDO directly after the macro
functions you want to be able to repeat, because other macro functions
following CUSTOM.UNDO might reset the Undo command.

**Related Function**

[CUSTOM.REPEAT](CUSTOM.REPEAT.md)&nbsp;&nbsp;&nbsp;Specifies a macro to run to repeat a
custom command



Return to [README](README.md#C)

# CUSTOM.UNDO

Creates a customized Undo tool and Undo or Redo command on the Edit menu
for custom commands.

**Syntax**

**CUSTOM.UNDO**(**macro\_text**, undo\_text)

Macro\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of, or an R1C1-style
reference to, the macro you want to run when the Undo command is chosen.
Macro\_text can be the name or cell reference of a macro.

Undo\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text you want to use as the
Undo command.

**Example**

The following macro function runs the UndoMult macro when the user
clicks the Undo Times100 command, a custom command that multiples the
current cell by 100.

\=CUSTOM.UNDO("UndoMult", "\&Undo Times100")

**Tip**&nbsp;&nbsp;&nbsp;Use CUSTOM.UNDO directly after the macro
functions you want to be able to repeat, because other macro functions
following CUSTOM.UNDO might reset the Undo command.

**Related Function**

[CUSTOM.REPEAT](CUSTOM.REPEAT.md)&nbsp;&nbsp;&nbsp;Specifies a macro to run to repeat a
custom command



Return to [README](README.md#C)

