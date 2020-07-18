# MAIL.LOGON

Starts a mail session.

**Important**&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;To use MAIL.LOGON in Microsoft Excel for
Windows, you must be using a mail client that supports the Messaging
Applications Programming Interface (MAPI) or Vendor-Independent
Messaging (VIM). The function is available for only Microsoft Excel for
Windows.

**Syntax**

**MAIL.LOGON**(name\_text, password\_text, download\_logical)

**MAIL.LOGON**?(name\_text, password\_text, download\_logical)

Name\_text&nbsp;&nbsp;&nbsp;&nbsp;is the username of the mail account or
Microsoft Exchange profile name. If omitted, prompts for username.

Password\_text&nbsp;&nbsp;&nbsp;&nbsp;is the password for the account.
If omitted, prompts for password. Ignored when the dialog box form is
used. This argument is ignored in Microsoft Exchange.

Download\_logical&nbsp;&nbsp;&nbsp;&nbsp;specifies whether to download
new mail. Use TRUE to download new mail; use FALSE or leave blank to
skip downloading new mail.

**Remarks**

Returns FALSE if you cancel the dialog box or \#VALUE\! if you can't log
on.

If you omit both name\_text and password\_text, Microsoft Excel tries to
log on using an existing mail session.

**Related Function**

[MAIL.LOGOFF](MAIL.LOGOFF.md)&nbsp;&nbsp;&nbsp;Ends the current mail session



Return to [README](README.md)

