GET.LINK.INFO

Returns information about the specified link. Use GET.LINK.INFO to get
information about the update settings of a link.

**Syntax**

**GET.LINK.INFO**(**link\_text, type\_num**, type\_of\_link, reference)

Link\_text    is the path of the link as displayed in the Links dialog
box, which appears when you choose the Links command from the Edit menu.
The path to the file you wish to return DDE information on must be
surrounded by single quotes.

Type\_num    is a number that specifies what type of information about
the currently selected link to return. Type\_num 2 applies only to
publishers and subscribers in Microsoft Excel for the Macintosh.

|               |                                                                                                                |
| ------------- | -------------------------------------------------------------------------------------------------------------- |
| **Type\_num** | **Returns**                                                                                                    |
| 1             | If the link is set to automatic update, returns 1; otherwise 2.                                                |
| 2             | Date of the latest edition as a serial number. Returns \#N/A if link\_text is not a publisher or a subscriber. |

Type\_of\_link    is a number from 1 to 6 that specifies what type of
link you want to get information about.

|                    |                              |
| ------------------ | ---------------------------- |
| **Type\_of\_link** | **Link document type**       |
| 1                  | Not applicable               |
| 2                  | DDE link (Microsoft Windows) |
| 3                  | Not applicable               |
| 4                  | Not applicable               |
| 5                  | Publisher (Macintosh)        |
| 6                  | Subscriber (Macintosh)       |

Reference    specifies the cell range in R1C1 format of the publisher or
subscriber that you want information about. Reference is required if you
have more than one publisher or subscriber of a single edition name on
the active workbook. Use reference to specify the location of the
subscriber you want to return information about. If the subscriber is a
picture, or if the publisher is an embedded chart, reference is the
number of the object as displayed in the Name box.

**Remarks**

  - > If Microsoft Excel cannot find link\_text, or if type\_of\_link
    > does not match the link specified by link\_text, GET.LINK.INFO
    > returns the \#VALUE\! error value.

  - > If you have more than one subscriber to the edition link\_text or
    > if the same area is published more than once, you must specify
    > reference.

>  

**Example**

In Microsoft Excel for Windows, the following macro formula returns
information about a DDE link to a Microsoft Word for Windows document.
The document is named NEWPROD.DOC.

GET.LINK.INFO("WinWord|'C:\\WINWORD\\NEWPROD.DOC'\!DDE\_LINK1", 1, 2)

In Microsoft Excel for the Macintosh, the following macro formula
returns information about a link to a publisher defined in cells A1:C3
on a workbook named New Products.

GET.LINK.INFO("A1:C3 New Products Edition \#1", 2, 5, "'New
Products'\!R1C1:R3C3")

**Related Functions**

[CREATE.PUBLISHER](CREATE.PUBLISHER.md)   Creates a publisher from the selection

[SUBSCRIBE.TO](SUBSCRIBE.TO.md)   Inserts contents of an edition into the active workbook

[UPDATE.LINK](UPDATE.LINK.md)   Updates a link to another workbook



Return to [README](README.md)

