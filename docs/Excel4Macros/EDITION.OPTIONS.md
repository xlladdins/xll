EDITION.OPTIONS

Sets options in, or performs actions on, the specified publisher or
subscriber. In Microsoft Excel for Windows, EDITION.OPTIONS also allows
you to cancel a publisher or subscriber created in Microsoft Excel for
the Macintosh.

**Syntax**

**EDITION.OPTIONS**(**edition\_type**, edition\_name, reference,
**option**, appearance, size, formats)

Edition\_type    is the number 1 or 2 specifying the type of edition.

|                   |                     |
| ----------------- | ------------------- |
| **Edition\_type** | **Type of edition** |
| 1                 | Publisher           |
| 2                 | Subscriber          |

Edition\_name    is the name of the edition you want to change the
edition options for or to perform actions on. If edition\_name is
omitted, reference is required.

Reference    specifies the range (given in text form as a name or an
R1C1-style reference) occupied by the publisher or subscriber.

  - > Reference is required if you have more than one publisher or
    > subscriber of edition\_name on the active workbook. Use reference
    > to specify the location of the publisher or subscriber for which
    > you want to set options.

  - > If edition\_type is 1 and the publisher is an embedded chart, or
    > if edition\_type is 2 and the subscriber is a picture, reference
    > is the object identifier as displayed in the reference area.

  - > If reference is omitted, edition\_name is required.

>  

Option    is a number from 1 to 6 specifying the edition option you want
to set or the action you want to take, according to the following two
tables. Options 2 to 6 are only available if you are using Microsoft
Excel for the Macintosh with system software version 7.0 or later.

If a publisher is specified, then option applies as follows.

|            |                                                                        |
| ---------- | ---------------------------------------------------------------------- |
| **Option** | **Action**                                                             |
| 1          | Cancels the publisher                                                  |
| 2          | Sends the edition now                                                  |
| 3          | Selects the range or object published to the specified edition         |
| 4          | Automatically updates the edition when the file is saved               |
| 5          | Updates the edition on request only                                    |
| 6          | Changes the edition file as specified by appearance, size, and formats |

If a subscriber is specified, then option applies as follows.

|            |                                                  |
| ---------- | ------------------------------------------------ |
| **Option** | **Action**                                       |
| 1          | Cancels the subscriber                           |
| 2          | Gets the latest edition                          |
| 3          | Opens the publisher workbook                     |
| 4          | Automatically updates when new data is available |
| 5          | Update on request only                           |

The following three arguments are available only when option is 6.

Appearance    specifies whether the selection is published as shown on
screen or as shown when printed. The default value for appearance is 1
if the selection is a sheet or macro sheet and 2 if the selection is a
chart.

|                |                            |
| -------------- | -------------------------- |
| **Appearance** | **Selection is published** |
| 1              | As shown on screen         |
| 2              | As shown when printed      |

Size    specifies the size of a published chart. Size is only available
if a chart is to be published.

|              |                             |
| ------------ | --------------------------- |
| **Size**     | **Chart size is published** |
| 1 or omitted | As shown on screen          |
| 2            | As shown when printed       |

Formats    is a number specifying the format of the file.

|              |                 |
| ------------ | --------------- |
| **Formats**  | **File format** |
| 1 or omitted | PICT            |
| 2            | BIFF            |
| 4            | RTF             |
| 8            | VALU            |

You can also use the sum of the allowable file formats. For example, a
value of 6 specifies BIFF and RTF.

**Example**

The following macro formula opens the workbook (and application) that
published the edition named Monthly Totals:

EDITION.OPTIONS(2, "Monthly Totals", , 3)

**Related Functions**

[CREATE.PUBLISHER](CREATE.PUBLISHER.md)   Creates a publisher from the selection

[GET.LINK.INFO](GET.LINK.INFO.md)   Returns information about a link

[SUBSCRIBE.TO](SUBSCRIBE.TO.md)   Inserts contents of an edition into the active workbook



Return to [README.md](README.md)

EDITION.OPTIONS

Sets options in, or performs actions on, the specified publisher or
subscriber. In Microsoft Excel for Windows, EDITION.OPTIONS also allows
you to cancel a publisher or subscriber created in Microsoft Excel for
the Macintosh.

**Syntax**

**EDITION.OPTIONS**(**edition\_type**, edition\_name, reference,
**option**, appearance, size, formats)

Edition\_type    is the number 1 or 2 specifying the type of edition.

|                   |                     |
| ----------------- | ------------------- |
| **Edition\_type** | **Type of edition** |
| 1                 | Publisher           |
| 2                 | Subscriber          |

Edition\_name    is the name of the edition you want to change the
edition options for or to perform actions on. If edition\_name is
omitted, reference is required.

Reference    specifies the range (given in text form as a name or an
R1C1-style reference) occupied by the publisher or subscriber.

  - > Reference is required if you have more than one publisher or
    > subscriber of edition\_name on the active workbook. Use reference
    > to specify the location of the publisher or subscriber for which
    > you want to set options.

  - > If edition\_type is 1 and the publisher is an embedded chart, or
    > if edition\_type is 2 and the subscriber is a picture, reference
    > is the object identifier as displayed in the reference area.

  - > If reference is omitted, edition\_name is required.

>  

Option    is a number from 1 to 6 specifying the edition option you want
to set or the action you want to take, according to the following two
tables. Options 2 to 6 are only available if you are using Microsoft
Excel for the Macintosh with system software version 7.0 or later.

If a publisher is specified, then option applies as follows.

|            |                                                                        |
| ---------- | ---------------------------------------------------------------------- |
| **Option** | **Action**                                                             |
| 1          | Cancels the publisher                                                  |
| 2          | Sends the edition now                                                  |
| 3          | Selects the range or object published to the specified edition         |
| 4          | Automatically updates the edition when the file is saved               |
| 5          | Updates the edition on request only                                    |
| 6          | Changes the edition file as specified by appearance, size, and formats |

If a subscriber is specified, then option applies as follows.

|            |                                                  |
| ---------- | ------------------------------------------------ |
| **Option** | **Action**                                       |
| 1          | Cancels the subscriber                           |
| 2          | Gets the latest edition                          |
| 3          | Opens the publisher workbook                     |
| 4          | Automatically updates when new data is available |
| 5          | Update on request only                           |

The following three arguments are available only when option is 6.

Appearance    specifies whether the selection is published as shown on
screen or as shown when printed. The default value for appearance is 1
if the selection is a sheet or macro sheet and 2 if the selection is a
chart.

|                |                            |
| -------------- | -------------------------- |
| **Appearance** | **Selection is published** |
| 1              | As shown on screen         |
| 2              | As shown when printed      |

Size    specifies the size of a published chart. Size is only available
if a chart is to be published.

|              |                             |
| ------------ | --------------------------- |
| **Size**     | **Chart size is published** |
| 1 or omitted | As shown on screen          |
| 2            | As shown when printed       |

Formats    is a number specifying the format of the file.

|              |                 |
| ------------ | --------------- |
| **Formats**  | **File format** |
| 1 or omitted | PICT            |
| 2            | BIFF            |
| 4            | RTF             |
| 8            | VALU            |

You can also use the sum of the allowable file formats. For example, a
value of 6 specifies BIFF and RTF.

**Example**

The following macro formula opens the workbook (and application) that
published the edition named Monthly Totals:

EDITION.OPTIONS(2, "Monthly Totals", , 3)

**Related Functions**

[CREATE.PUBLISHER](CREATE.PUBLISHER.md)   Creates a publisher from the selection

[GET.LINK.INFO](GET.LINK.INFO.md)   Returns information about a link

[SUBSCRIBE.TO](SUBSCRIBE.TO.md)   Inserts contents of an edition into the active workbook



Return to [README.md](README.md)

