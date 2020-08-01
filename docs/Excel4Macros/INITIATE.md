# INITIATE

Opens a dynamic data exchange (DDE) channel to an application and
returns the number of the open channel. Once you have opened a channel
to another application with INITIATE, you can use EXECUTE and SEND.KEYS
to control the other application from a Microsoft Excel macro.
(SEND.KEYS is available only with Microsoft Excel for Windows.) If
INITIATE is successful, it returns the number of the open channel. All
the subsequent DDE macro functions use this number to specify the
channel. If INITIATE is unsuccessful, FALSE is returned.

**Important**&nbsp;&nbsp;&nbsp;Microsoft Excel for the Macintosh
requires system software version 7.0 or later for this function.

**Syntax**

**INITIATE**(**app\_text, topic\_text**)

App\_text&nbsp;&nbsp;&nbsp;&nbsp;is the DDE name of the application with
which you want to begin a DDE session, in text form. The form of
app\_text depends on the application you are accessing. The DDE name of
Microsoft Excel, for example, is "Excel".

Topic\_text&nbsp;&nbsp;&nbsp;&nbsp;describes something, such as a
document or a record in a database, in the application that you are
accessing; the form of topic\_text depends on the application you are
accessing. Microsoft Excel accepts the names of the current documents as
topic\_text, as well as the name "System".

**Remarks**

  - > You can specify an instance of an application by appending the
    > application's task ID number to the app\_text argument. If you
    > start an application by using the EXEC function, EXEC returns the
    > task ID number for that instance of the application.

  - > If more than one instance of an application is running and you do
    > not specify which instance you would like to open a channel to,
    > INITIATE displays a dialog box from which you can choose the
    > instance you want. You can prevent this dialog box from appearing
    > by disabling or redirecting errors with the ERROR function.


**Example**

The following macro formula opens a channel to the document named MEMO
in the application named WORD:

INITIATE("WORD", "MEMO")

**Related Functions**

[POKE](POKE.md)&nbsp;&nbsp;&nbsp;Sends data to another application

[REQUEST](REQUEST.md)&nbsp;&nbsp;&nbsp;Returns data from another application

[TERMINATE](TERMINATE.md)&nbsp;&nbsp;&nbsp;Closes a channel to another application

[EXECUTE](EXECUTE.md)&nbsp;&nbsp;&nbsp;Carries out a command in another application

[EXEC](EXEC.md)&nbsp;&nbsp;&nbsp;Starts a separate program



Return to [README](README.md#I)

# INITIATE

Opens a dynamic data exchange (DDE) channel to an application and
returns the number of the open channel. Once you have opened a channel
to another application with INITIATE, you can use EXECUTE and SEND.KEYS
to control the other application from a Microsoft Excel macro.
(SEND.KEYS is available only with Microsoft Excel for Windows.) If
INITIATE is successful, it returns the number of the open channel. All
the subsequent DDE macro functions use this number to specify the
channel. If INITIATE is unsuccessful, FALSE is returned.

**Important**&nbsp;&nbsp;&nbsp;Microsoft Excel for the Macintosh
requires system software version 7.0 or later for this function.

**Syntax**

**INITIATE**(**app\_text, topic\_text**)

App\_text&nbsp;&nbsp;&nbsp;&nbsp;is the DDE name of the application with
which you want to begin a DDE session, in text form. The form of
app\_text depends on the application you are accessing. The DDE name of
Microsoft Excel, for example, is "Excel".

Topic\_text&nbsp;&nbsp;&nbsp;&nbsp;describes something, such as a
document or a record in a database, in the application that you are
accessing; the form of topic\_text depends on the application you are
accessing. Microsoft Excel accepts the names of the current documents as
topic\_text, as well as the name "System".

**Remarks**

  - > You can specify an instance of an application by appending the
    > application's task ID number to the app\_text argument. If you
    > start an application by using the EXEC function, EXEC returns the
    > task ID number for that instance of the application.

  - > If more than one instance of an application is running and you do
    > not specify which instance you would like to open a channel to,
    > INITIATE displays a dialog box from which you can choose the
    > instance you want. You can prevent this dialog box from appearing
    > by disabling or redirecting errors with the ERROR function.


**Example**

The following macro formula opens a channel to the document named MEMO
in the application named WORD:

INITIATE("WORD", "MEMO")

**Related Functions**

[POKE](POKE.md)&nbsp;&nbsp;&nbsp;Sends data to another application

[REQUEST](REQUEST.md)&nbsp;&nbsp;&nbsp;Returns data from another application

[TERMINATE](TERMINATE.md)&nbsp;&nbsp;&nbsp;Closes a channel to another application

[EXECUTE](EXECUTE.md)&nbsp;&nbsp;&nbsp;Carries out a command in another application

[EXEC](EXEC.md)&nbsp;&nbsp;&nbsp;Starts a separate program



Return to [README](README.md#I)

