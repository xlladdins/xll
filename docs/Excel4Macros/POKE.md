# POKE

Sends data to another application. Use POKE to send data to documents in
other applications you are communicating with through dynamic data
exchange (DDE).

**Syntax**

**POKE**(**channel\_num, item\_text**, **data\_ref**)

**Important**&nbsp;&nbsp;&nbsp;Microsoft Excel for the Macintosh
requires system software version 7.0 or later for this function.

Channel\_num&nbsp;&nbsp;&nbsp;&nbsp;is the channel number returned by a
previously run INITIATE function.

Item\_text&nbsp;&nbsp;&nbsp;&nbsp;is text that identifies the item you
want to send data to in the application you are accessing through
channel\_num. The form of item\_text depends on the application
connected to channel\_num.

Data\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the workbook
containing the data to send.

If POKE is not successful, it returns the following values.

|                    |                                                                                                                 |
| ------------------ | --------------------------------------------------------------------------------------------------------------- |
| **Value returned** | **Meaning**                                                                                                     |
| \#VALUE\!          | Channel\_num is not a valid channel number.                                                                     |
| \#DIV/0\!          | The application you are accessing does not respond after a certain length of time, and you press ESC to cancel. |
| \#REF\!            | POKE is refused.                                                                                                |

**Examples**

In Microsoft Excel for Windows, the following macro inserts the text
from cell C3 into the Microsoft Word for Windows document SALES.DOC at
the start of the document.

\=POKE(SendChanl, "StartOfDoc", C3)

In Microsoft Excel for the Macintosh, the following macro inserts the
text from cell C3 into the Microsoft Word document named Report.

\=POKE(SendChanl, "TopicName", C3)

**Related Functions**

[INITIATE](INITIATE.md)&nbsp;&nbsp;&nbsp;Opens a channel to another application

[REQUEST](REQUEST.md)&nbsp;&nbsp;&nbsp;Returns data from another application

[TERMINATE](TERMINATE.md)&nbsp;&nbsp;&nbsp;Closes a channel to another application



Return to [README](README.md#P)

# POKE

Sends data to another application. Use POKE to send data to documents in
other applications you are communicating with through dynamic data
exchange (DDE).

**Syntax**

**POKE**(**channel\_num, item\_text**, **data\_ref**)

**Important**&nbsp;&nbsp;&nbsp;Microsoft Excel for the Macintosh
requires system software version 7.0 or later for this function.

Channel\_num&nbsp;&nbsp;&nbsp;&nbsp;is the channel number returned by a
previously run INITIATE function.

Item\_text&nbsp;&nbsp;&nbsp;&nbsp;is text that identifies the item you
want to send data to in the application you are accessing through
channel\_num. The form of item\_text depends on the application
connected to channel\_num.

Data\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the workbook
containing the data to send.

If POKE is not successful, it returns the following values.

|                    |                                                                                                                 |
| ------------------ | --------------------------------------------------------------------------------------------------------------- |
| **Value returned** | **Meaning**                                                                                                     |
| \#VALUE\!          | Channel\_num is not a valid channel number.                                                                     |
| \#DIV/0\!          | The application you are accessing does not respond after a certain length of time, and you press ESC to cancel. |
| \#REF\!            | POKE is refused.                                                                                                |

**Examples**

In Microsoft Excel for Windows, the following macro inserts the text
from cell C3 into the Microsoft Word for Windows document SALES.DOC at
the start of the document.

\=POKE(SendChanl, "StartOfDoc", C3)

In Microsoft Excel for the Macintosh, the following macro inserts the
text from cell C3 into the Microsoft Word document named Report.

\=POKE(SendChanl, "TopicName", C3)

**Related Functions**

[INITIATE](INITIATE.md)&nbsp;&nbsp;&nbsp;Opens a channel to another application

[REQUEST](REQUEST.md)&nbsp;&nbsp;&nbsp;Returns data from another application

[TERMINATE](TERMINATE.md)&nbsp;&nbsp;&nbsp;Closes a channel to another application



Return to [README](README.md#P)

# POKE

Sends data to another application. Use POKE to send data to documents in
other applications you are communicating with through dynamic data
exchange (DDE).

**Syntax**

**POKE**(**channel\_num, item\_text**, **data\_ref**)

**Important**&nbsp;&nbsp;&nbsp;Microsoft Excel for the Macintosh
requires system software version 7.0 or later for this function.

Channel\_num&nbsp;&nbsp;&nbsp;&nbsp;is the channel number returned by a
previously run INITIATE function.

Item\_text&nbsp;&nbsp;&nbsp;&nbsp;is text that identifies the item you
want to send data to in the application you are accessing through
channel\_num. The form of item\_text depends on the application
connected to channel\_num.

Data\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the workbook
containing the data to send.

If POKE is not successful, it returns the following values.

|                    |                                                                                                                 |
| ------------------ | --------------------------------------------------------------------------------------------------------------- |
| **Value returned** | **Meaning**                                                                                                     |
| \#VALUE\!          | Channel\_num is not a valid channel number.                                                                     |
| \#DIV/0\!          | The application you are accessing does not respond after a certain length of time, and you press ESC to cancel. |
| \#REF\!            | POKE is refused.                                                                                                |

**Examples**

In Microsoft Excel for Windows, the following macro inserts the text
from cell C3 into the Microsoft Word for Windows document SALES.DOC at
the start of the document.

\=POKE(SendChanl, "StartOfDoc", C3)

In Microsoft Excel for the Macintosh, the following macro inserts the
text from cell C3 into the Microsoft Word document named Report.

\=POKE(SendChanl, "TopicName", C3)

**Related Functions**

[INITIATE](INITIATE.md)&nbsp;&nbsp;&nbsp;Opens a channel to another application

[REQUEST](REQUEST.md)&nbsp;&nbsp;&nbsp;Returns data from another application

[TERMINATE](TERMINATE.md)&nbsp;&nbsp;&nbsp;Closes a channel to another application



Return to [README](README.md#P)

