ON.KEY

Runs a specified macro when a particular key or key combination is
pressed.

**Syntax**

**ON.KEY**(**key\_text**, macro\_text)

Key\_text&nbsp;&nbsp;&nbsp;&nbsp;can specify any single key, or any key
combined with ALT, CTRL, or SHIFT, or any combination of those keys (in
Microsoft Excel for Windows) or COMMAND, CTRL, OPTION, or SHIFT or any
combination of those keys (in Microsoft Excel for the Macintosh). Each
key is represented by one or more characters, such as "a" for the
character a, or "{ENTER}" for the ENTER key.

To specify characters that aren't displayed when you press the key, such
as ENTER or TAB, use the codes shown in the following table. Each code
in the table represents one key on the keyboard.

|                        |                         |
| ---------------------- | ----------------------- |
| **Key**                | **Code**                |
| BACKSPACE              | "{BACKSPACE}" or "{BS}" |
| BREAK                  | "{BREAK}"               |
| CAPS LOCK              | "{CAPSLOCK}"            |
| CLEAR                  | "{CLEAR}"               |
| DELETE or DEL          | "{DELETE}" or "{DEL}"   |
| DOWN                   | "{DOWN}"                |
| END                    | "{END}"                 |
| ENTER (numeric keypad) | "{ENTER}"               |
| ENTER                  | "\~" (tilde)            |
| ESC                    | "{ESCAPE} or {ESC}"     |
| HELP                   | "{HELP}"                |
| HOME                   | "{HOME}"                |
| INS                    | "{INSERT}"              |
| LEFT                   | "{LEFT}"                |
| NUM LOCK               | "{NUMLOCK}"             |
| PAGE DOWN              | "{PGDN}"                |
| PAGE UP                | "{PGUP}"                |
| RETURN                 | "{RETURN}"              |
| RIGHT                  | "{RIGHT}"               |
| SCROLL LOCK            | "{SCROLLLOCK}"          |
| TAB                    | "{TAB}"                 |
| UP                     | "{UP}"                  |
| F1 through F15         | "{F1}" through "{F15}"  |

In Microsoft Excel for Windows, you can also specify keys combined with
SHIFT and/or CTRL and/or ALT. In Microsoft Excel for the Macintosh, you
can also specify keys combined with SHIFT and/or CTRL and/or OPTION
and/or COMMAND. To specify a key combined with another key or keys, use
the following table.

|                     |                             |
| ------------------- | --------------------------- |
| **To combine with** | **Precede the key code by** |
| SHIFT               | "+" (plus sign)             |
| CTRL                | "^" (caret)                 |
| ALT or OPTION       | "%" (percent sign)          |
| COMMAND             | "\*" (asterisk)             |

To assign a macro to one of the special characters (+, ^, %, and so on),
enclose the character in brackets. For example, ON.KEY("^{+}",
"InsertItem") assigns a macro named InsertItem to the key sequence
CTRL+PLUS SIGN.

Macro\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a macro that you want
to run when key\_text is pressed. The reference must be in text form.

  - > If macro\_text is "" (empty text), nothing happens when key\_text
    > is pressed. This form of ON.KEY disables the normal meaning of
    > keystrokes in Microsoft Excel.

  - > If macro\_text is omitted, key\_text reverts to its normal meaning
    > in Microsoft Excel, and any special key assignments made with
    > previous ON.KEY functions are cleared.

> &nbsp;

**Remarks**

  - > ON.KEY remains in effect until you clear it or quit Microsoft
    > Excel. You can clear ON.KEY by specifying key\_text and omitting
    > the macro\_text argument.

  - > If the macro sheet containing macro\_text is closed when you press
    > key\_text, an error is returned.

  - > If another macro is running when you press key\_text, macro\_text
    > will not run.

> &nbsp;

**Examples**

Suppose you wanted the key combination SHIFT+CTRL+RIGHT to run the macro
Print. You use the following macro formula:

ON.KEY("+^{RIGHT}", "Print")

To return SHIFT+CTRL+RIGHT to its normal meaning, you would use the
following macro formula:

ON.KEY("+^{RIGHT}")

To disable SHIFT+CTRL+RIGHT altogether, you would use the following
macro formula:

ON.KEY("+^{RIGHT}", "")

**Related Functions**

[CANCEL.KEY](CANCEL.KEY.md)&nbsp;&nbsp;&nbsp;Disables macro interruption

[ERROR](ERROR.md)&nbsp;&nbsp;&nbsp;Specifies what action to take if an error is
encountered while a macro is running

[SEND.KEYS](SEND.KEYS.md)&nbsp;&nbsp;&nbsp;Sends a key sequence to an application



Return to [README](README.md)

