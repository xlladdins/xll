SEND.MAIL

Equivalent to clicking the Send Mail command on the File menu. Sends the
active workbook using email.

**Syntax**

**SEND.MAIL**(**recipients**, subject, return\_receipt)

**SEND.MAIL**?(recipients, subject, return\_receipt)

**Important****&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;To use SEND.MAIL in Microsoft Excel for
Windows, you must be using a mail client that supports the Messaging
Applications Programming Interface (MAPI) or Vendor-Independent
Messaging (VIM). To use SEND.MAIL in Microsoft Excel for the Macintosh,
you must be using Microsoft Mail version 2.0 or later.

Recipients**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is the name of the person to whom you
want to send the mail. The name should be given as text.

  - > To specify more than one name, give the list of names as an array.
    > For example, SEND.MAIL({"John", "Paul", "George", "Ringo"}) would
    > send the active workbook to the four names in the array. You can
    > also refer to a range on a sheet or macro sheet that contains a
    > list of names to whom you want the mail to be sent.

  - > To send mail to users on different Microsoft Mail for the
    > Macintosh servers, specify the server name along with the user
    > name. The following text, as the recipients argument, sends mail
    > to wandagr on server2, gregpr on the current server, and victorge
    > on server7:

> {"wandagr@server2", "gregpr", "victorge@server7"}
> 

Subject**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a text string that specifies the
subject of the message. If subject is omitted, the name of the active
workbook is used as the subject.

Return\_receipt**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value that
corresponds to the Return Receipt check box. If return\_receipt is TRUE,
Microsoft Excel selects the check box and sends a return receipt; if
FALSE or omitted, Microsoft Excel clears the check box.

**Related Function**

[OPEN.MAIL](OPEN.MAIL.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Opens files sent via Microsoft Mail that
[M](M.md)icrosoft Excel can open



Return to [README](README.md)

