CREATE.PUBLISHER

Equivalent to clicking the Create Publisher command on the Publishing
submenu of the Edit menu. Publishes the selected range or chart to an
edition file for use by other Macintosh applications.

**Important****&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;This function is only available if you
are using Microsoft Excel for the Macintosh with system software version
7.0 or later.

**Syntax**

**CREATE.PUBLISHER**(file\_text, appearance, size, formats)

**CREATE.PUBLISHER**?(file\_text, appearance, size, formats)

File\_text**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a text string to be used as the
name of the new file that will contain the selected data. If file\_text
is omitted, Microsoft Excel uses the format "\<WorkbookName\> Edition
\#n", where WorkbookName is the name of the workbook from which the
publisher is being created, Edition indicates that the file is an
edition file, and n is a unique integer.

For example, if you omit file\_text and are publishing a selection from
a workbook named Seasonal, and it is your third publisher from that
workbook in the current work session, the default name of the publisher
would be "Seasonal Edition \#3".

Appearance**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies whether the selection is to
be published as shown on screen or as shown when printed. The default
value for appearance is 1 if the selection is a sheet and 2 if the
selection is a chart.

|                |                            |
| -------------- | -------------------------- |
| **Appearance** | **Selection is published** |
| 1              | As shown on screen         |
| 2              | As shown when printed      |

Size**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;specifies the size at which to publish a
chart. Size is only available if a chart is to be published.

|              |                        |
| ------------ | ---------------------- |
| **Size**     | **Chart is published** |
| 1 or omitted | As shown on screen     |
| 2            | As shown when printed  |

Formats**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is number specifying what file format or
formats CREATE.PUBLISHER should use when it creates the Edition file.

|             |                 |
| ----------- | --------------- |
| **Formats** | **File format** |
| 1           | PICT            |
| 2           | BIFF            |
| 4           | RTF             |
| 8           | VALU            |

&nbsp;

  - > You can also use the sum of the allowable file formats for
    > formats. For example, a value of 6 specifies BIFF and RTF.

  - > If formats is omitted and the document is a sheet, formats is
    > assumed to be 15 (all formats); if the document is a chart,
    > formats is assumed to be 1 (PICT).


**Related Functions**

[EDITION.OPTIONS](EDITION.OPTIONS.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Sets publisher and subscriber options

[GET.LINK.INFO](GET.LINK.INFO.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns information about a link

[SUBSCRIBE.TO](SUBSCRIBE.TO.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Inserts contents of an edition into the
active workbook

[UPDATE.LINK](UPDATE.LINK.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Updates a link to another workbook



Return to [README](README.md)

