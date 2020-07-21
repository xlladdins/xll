# MESSAGE

Displays and removes messages in the message area of the status bar.
MESSAGE is useful for displaying text that doesn't need a response, such
as descriptions of commands in user-defined menus.

**Syntax**

**MESSAGE**(**logical**, text)

Logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether to
display or remove a message.

  - > If logical is TRUE, Microsoft Excel displays text in the message
    > area of the status bar.

  - > If logical is FALSE, Microsoft Excel removes any messages, and the
    > status bar is returned to normal (that is, command help messages
    > are displayed).

Text&nbsp;&nbsp;&nbsp;&nbsp;is the message you want to display in the
status bar. If text is "" (empty text), Microsoft Excel removes any
messages currently displayed in the status bar.

**Remarks**

  - > Only one message can be displayed in the status bar at a time.
    > Messages are always displayed in the same place.

  - > MESSAGE works the same way whether the status bar is displayed or
    > not. You can, for example, use MESSAGE while the status bar isn't
    > displayed. As soon as you display the status bar, you see your
    > message.

  - > If you display any message (even empty text) and don't remove it
    > with MESSAGE(FALSE), that message is displayed until you quit
    > Microsoft Excel.

  - > You can also use the ALERT function to get the user's attention;
    > however, this interrupts the macro and requires the user's
    > intervention before the macro can continue.


**Example**

The following lines in a macro display a message warning that you must
wait for a moment while the macro calls a subroutine.

MESSAGE(TRUE, "One moment please...")

**Related Functions**

[ALERT](ALERT.md)&nbsp;&nbsp;&nbsp;Displays a dialog box and a message

[BEEP](BEEP.md)&nbsp;&nbsp;&nbsp;Sounds a tone



Return to [README](README.md#M)

