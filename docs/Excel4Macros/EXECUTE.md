EXECUTE

Carries out commands in another program with which you have a dynamic
data exchange (DDE) link. Use with EXEC, INITIATE, and SEND.KEYS to run
another program through Microsoft Excel. (SEND.KEYS is available only in
Microsoft Excel for Windows.)

**Important**&nbsp;&nbsp;&nbsp;Microsoft Excel for the Macintosh
requires system software version 7.0 or later for this function.

**Syntax**

**EXECUTE**(**channel\_num, execute\_text**)

Channel\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number returned by a previously
run INITIATE function. Channel\_num refers to a channel through which
Microsoft Excel communicates with another program.

Execute\_text&nbsp;&nbsp;&nbsp;&nbsp;is a text string representing
commands you want to carry out in the program specified by channel\_num.
The form of execute\_text depends on the program you are referring to.
To include specific key sequences in execute\_text, use the format
described under key\_text in the ON.KEY function.

If EXECUTE is not successful, it returns one of the following error
values:

|                    |                                                                                                                  |
| ------------------ | ---------------------------------------------------------------------------------------------------------------- |
| **Value returned** | **Situation**                                                                                                    |
| \#VALUE\!          | Channel\_num is not a valid channel number.                                                                      |
| \#N/A              | The program you are accessing is busy.                                                                           |
| \#DIV/0\!          | The program you are accessing does not respond after a certain length of time or you have pressed ESC to cancel. |
| \#REF\!            | The keys specified in execute\_text are refused by the application which you want to access.                     |

**Remarks**

Commands sent to another program with EXECUTE will not work when a
dialog box is displayed in the program. In Microsoft Excel for Windows,
you can use SEND.KEYS to send commands that make selections in a dialog
box.

**Examples**

The following macro formula sends the number 25 and a carriage return to
the application identified by channel\_num 14:

EXECUTE(14, "25\~")

**Related Functions**

[EXEC](EXEC.md)&nbsp;&nbsp;&nbsp;Starts another application

[INITIATE](INITIATE.md)&nbsp;&nbsp;&nbsp;Opens a channel to another application

[POKE](POKE.md)&nbsp;&nbsp;&nbsp;Sends data to another application

[REQUEST](REQUEST.md)&nbsp;&nbsp;&nbsp;Returns data from another application

[SEND.KEYS](SEND.KEYS.md)&nbsp;&nbsp;&nbsp;Sends a key sequence to an application

[TERMINATE](TERMINATE.md)&nbsp;&nbsp;&nbsp;Closes a channel to another application



Return to [README](README.md)

