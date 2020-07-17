GET.WORKBOOK
============

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

+---------------+-----------------------------------------------------+
| **Type\_num** | **Returns**                                         |
+---------------+-----------------------------------------------------+
| 1             | The names of all sheets in the workbook, as a       |
|               | horizontal array of text values. Names are returned |
|               | as \[book\]sheet.                                   |
+---------------+-----------------------------------------------------+
| 2             | This will always return the \#N/A error value.      |
+---------------+-----------------------------------------------------+
| 3             | The names of the currently selected sheets in the   |
|               | workbook, as a horizontal array of text values.     |
+---------------+-----------------------------------------------------+
| 4             | The number of sheets in the workbook.               |
+---------------+-----------------------------------------------------+
| 5             | TRUE if the workbook has a routing slip; otherwise, |
|               | FALSE.                                              |
+---------------+-----------------------------------------------------+
| 6             | The names of all of the workbook routing recipients |
|               | who have not received the workbook, as a horizontal |
|               | array of text values.                               |
+---------------+-----------------------------------------------------+
| 7             | The subject line for the current routing slip, as   |
|               | text.                                               |
+---------------+-----------------------------------------------------+
| 8             | The message text for the routing slip, as text.     |
+---------------+-----------------------------------------------------+
| 9             | If the workbook is to be routed to recipients one   |
|               | after another, returns 1. If it is to be routed all |
|               | at once, returns 2.                                 |
+---------------+-----------------------------------------------------+
| 10            | TRUE, if the Return When Done check box in the      |
|               | Routing Slip dialog box is selected; otherwise,     |
|               | FALSE.                                              |
+---------------+-----------------------------------------------------+
| 11            | TRUE, if the current recipient has already          |
|               | forwarded the current workbook; otherwise, FALSE.   |
+---------------+-----------------------------------------------------+
| 12            | TRUE, if the Track Status checkbox in the Routing   |
|               | Slip dialog box is selected; otherwise, FALSE.      |
+---------------+-----------------------------------------------------+
| 13            | Status of the workbook routing slip:                |
|               |                                                     |
|               | 0 = Unrouted                                        |
|               |                                                     |
|               | 1 = Routing in progress, or the workbook has been   |
|               | routed to a user                                    |
|               |                                                     |
|               | 2 = Routing is finished                             |
+---------------+-----------------------------------------------------+
| 14            | TRUE, if the workbook structure is protected;       |
|               | otherwise, FALSE.                                   |
+---------------+-----------------------------------------------------+
| 15            | TRUE, if the workbook windows are protected;        |
|               | otherwise, FALSE.                                   |
+---------------+-----------------------------------------------------+
| 16            | Name of the workbook as text. The workbook name     |
|               | does not include the drive, directory or folder, or |
|               | window number.                                      |
+---------------+-----------------------------------------------------+
| 17            | TRUE if the workbook is read only; otherwise,       |
|               | FALSE. This is the equivalent of GET.DOCUMENT(34).  |
+---------------+-----------------------------------------------------+
| 18            | TRUE if sheet is write-reserved; otherwise, FALSE.  |
|               | This is the equivalent of GET.DOCUMENT(35).         |
+---------------+-----------------------------------------------------+
| 19            | Name of the user with current write permission for  |
|               | the workbook. This is the equivalent of             |
|               | GET.DOCUMENT(36).                                   |
+---------------+-----------------------------------------------------+
| 20            | Number corresponding to the file type of the        |
|               | document as displayed in the Save As dialog box.    |
|               | This is the equivalent of GET.DOCUMENT(37).         |
+---------------+-----------------------------------------------------+
| 21            | TRUE if the Always Create Backup check box is       |
|               | selected in the Save Options dialog box; otherwise, |
|               | FALSE. This is the equivalent of GET.DOCUMENT(40).  |
+---------------+-----------------------------------------------------+
| 22            | TRUE if the Save External Link Values check box is  |
|               | selected in the Calculation tab of the Options      |
|               | dialog box. This is the equivalent of               |
|               | GET.DOCUMENT(43).                                   |
+---------------+-----------------------------------------------------+
| 23            | TRUE if the workbook has a PowerTalk mailer;        |
|               | otherwise, FALSE. Returns \#N/A if no OCE mailer is |
|               | installed.                                          |
+---------------+-----------------------------------------------------+
| 24            | TRUE if changes have been made to the workbook      |
|               | since the last time it was saved; FALSE if book is  |
|               | unchanged (or when closed, will not prompt to be    |
|               | saved).                                             |
+---------------+-----------------------------------------------------+
| 25            | The recipients on the To line of a PowerTalk        |
|               | mailer, as a horizontal array of text.              |
+---------------+-----------------------------------------------------+
| 26            | The recipients on the Cc line of a PowerTalk        |
|               | mailer, as a horizontal array of text.              |
+---------------+-----------------------------------------------------+
| 27            | The recipients on the Bcc line of a PowerTalk       |
|               | mailer, as a horizontal array of text.              |
+---------------+-----------------------------------------------------+
| 28            | The subject of the PowerTalk mailer, as text.       |
+---------------+-----------------------------------------------------+
| 29            | The enclosures of the PowerTalk mailer, as a        |
|               | horizontal array of text.                           |
+---------------+-----------------------------------------------------+
| 30            | TRUE, if the PowerTalk mailer has been received     |
|               | from another user (as opposed to just being added   |
|               | but not sent). FALSE, if the mailer has not been    |
|               | received from another user.                         |
+---------------+-----------------------------------------------------+
| 31            | The date and time the PowerTalk mailer was sent, as |
|               | a serial number. Returns the \#N/A error value if   |
|               | the mailer has not yet been sent.                   |
+---------------+-----------------------------------------------------+
| 32            | The sender name of the PowerTalk mailer, as text.   |
|               | Returns the \#N/A error value if the mailer has not |
|               | yet been sent.                                      |
+---------------+-----------------------------------------------------+
| 33            | The title of the workbook as displayed on the       |
|               | Summary tab of the Properties dialog box, as text.  |
+---------------+-----------------------------------------------------+
| 34            | The subject of the workbook as displayed on the     |
|               | Summary tab of the Properties dialog box, as text.  |
+---------------+-----------------------------------------------------+
| 35            | The author of the workbook as displayed on the      |
|               | Summary tab of the Properties dialog box, as text.  |
+---------------+-----------------------------------------------------+
| 36            | The keywords for the workbook as displayed on the   |
|               | Summary tab of the Properties dialog box, as text.  |
+---------------+-----------------------------------------------------+
| 37            | The comments for the workbook as displayed on the   |
|               | Summary tab of the Properties dialog box, as text.  |
+---------------+-----------------------------------------------------+
| 38            | The name of the active sheet.                       |
+---------------+-----------------------------------------------------+

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

**Example**

The following macro formula returns the name of the active sheet in the
workbook named SALES.XLS:

GET.WORKBOOK(38, \"SALES.XLS\")

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

WORKBOOK.SELECT   Selects the specified documents in a workbook

Return to [top](#E)

GET.WORKSPACE
