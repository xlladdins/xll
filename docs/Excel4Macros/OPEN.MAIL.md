# OPEN.MAIL

Equivalent to clicking the Open Mail command on the Mail submenu on File
menu.

**Note**&nbsp;&nbsp;&nbsp;This function is available for only Microsoft
Excel for the Macintosh and Microsoft Mail version 2.0 or later.

**Syntax**

**OPEN.MAIL**(subject, comments)

**OPEN.MAIL**?(subject, comments)

Subject&nbsp;&nbsp;&nbsp;&nbsp;is the subject of the message containing
a file that Microsoft Excel can open.

  - > For each message whose subject matches the subject argument and
    > contains a file that Microsoft Excel can open, the file is opened
    > in Microsoft Excel; if the message has no unread enclosures, it is
    > deleted from the list of pending mail.

  - > If subject is omitted, then for all messages containing files that
    > Microsoft Excel can open, the files are opened; each message that
    > has no unread enclosures is deleted from the list of pending mail.


Comments&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether comments associated with the Microsoft Excel files are
displayed. If comments is TRUE, Microsoft Excel displays the comments;
if FALSE, comments are not displayed. If omitted, the current setting is
not changed.

**Tips**

  - > If you use consistent subjects in your Microsoft Mail messages,
    > you can easily create a macro that always opens mail messages with
    > certain files attached. For example, an OPEN.MAIL formula with
    > subject specified as "Weekly Report" would open the Microsoft
    > Excel file attached to the message containing that subject each
    > week.

  - > In OPEN.MAIL?, the dialog-box form of the function, the currently
    > running macro pauses while the Microsoft Mail documents window is
    > displayed. The macro resumes after you close the Microsoft Mail
    > documents window.

**Related Function**

[SEND.MAIL](SEND.MAIL.md)&nbsp;&nbsp;&nbsp;Sends the active workbook



Return to [README](README.md#O)

# OPEN.MAIL

Equivalent to clicking the Open Mail command on the Mail submenu on File
menu.

**Note**&nbsp;&nbsp;&nbsp;This function is available for only Microsoft
Excel for the Macintosh and Microsoft Mail version 2.0 or later.

**Syntax**

**OPEN.MAIL**(subject, comments)

**OPEN.MAIL**?(subject, comments)

Subject&nbsp;&nbsp;&nbsp;&nbsp;is the subject of the message containing
a file that Microsoft Excel can open.

  - > For each message whose subject matches the subject argument and
    > contains a file that Microsoft Excel can open, the file is opened
    > in Microsoft Excel; if the message has no unread enclosures, it is
    > deleted from the list of pending mail.

  - > If subject is omitted, then for all messages containing files that
    > Microsoft Excel can open, the files are opened; each message that
    > has no unread enclosures is deleted from the list of pending mail.


Comments&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that specifies
whether comments associated with the Microsoft Excel files are
displayed. If comments is TRUE, Microsoft Excel displays the comments;
if FALSE, comments are not displayed. If omitted, the current setting is
not changed.

**Tips**

  - > If you use consistent subjects in your Microsoft Mail messages,
    > you can easily create a macro that always opens mail messages with
    > certain files attached. For example, an OPEN.MAIL formula with
    > subject specified as "Weekly Report" would open the Microsoft
    > Excel file attached to the message containing that subject each
    > week.

  - > In OPEN.MAIL?, the dialog-box form of the function, the currently
    > running macro pauses while the Microsoft Mail documents window is
    > displayed. The macro resumes after you close the Microsoft Mail
    > documents window.

**Related Function**

[SEND.MAIL](SEND.MAIL.md)&nbsp;&nbsp;&nbsp;Sends the active workbook



Return to [README](README.md#O)

