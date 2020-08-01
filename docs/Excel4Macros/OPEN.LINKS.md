# OPEN.LINKS

Equivalent to clicking the Links command on the Edit menu. Use
OPEN.LINKS with the LINKS function to open workbooks linked to a
particular sheet.

**Syntax**

**OPEN.LINKS**(**document\_text1**, document\_text2, ..., read\_only,
type\_of\_link)

**OPEN.LINKS**?(document\_text1, document\_text2, ..., read\_only,
type\_of\_link)

Document\_text1, document\_text2,&nbsp;&nbsp;&nbsp;&nbsp;are 1 to 12
arguments that are the names of supporting documents in the form of
text, or arrays or references that contain text.

Read\_only&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the read/write status of the linked worksheet. If read\_only is TRUE,
the sheet can be modified but changes cannot be saved; if FALSE or
omitted, changes to the sheet can be saved. Read\_only applies only to
Microsoft Excel, WKS, and SYLK documents.

Type\_of\_link&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 6 that
specifies what type of link you want to get information about.

|                    |                        |
| ------------------ | ---------------------- |
| **Type\_of\_link** | **Link document type** |
| 1                  | Microsoft Excel link   |
| 2                  | DDE link               |
| 3                  | Reserved               |
| 4                  | Not applicable         |
| 5                  | Subscriber             |
| 6                  | Publisher              |

**Remarks**

You can generate an array of the names of linked workbooks with the
LINKS function.

**Related Functions**

[CHANGE.LINK](CHANGE.LINK.md)&nbsp;&nbsp;&nbsp;Changes supporting workbook links

[GET.LINK.INFO](GET.LINK.INFO.md)&nbsp;&nbsp;&nbsp;Returns information about a link

[LINKS](LINKS.md)&nbsp;&nbsp;&nbsp;Returns the name of all linked workbooks

[UPDATE.LINK](UPDATE.LINK.md)&nbsp;&nbsp;&nbsp;Updates a link to another document



Return to [README](README.md#O)

# OPEN.LINKS

Equivalent to clicking the Links command on the Edit menu. Use
OPEN.LINKS with the LINKS function to open workbooks linked to a
particular sheet.

**Syntax**

**OPEN.LINKS**(**document\_text1**, document\_text2, ..., read\_only,
type\_of\_link)

**OPEN.LINKS**?(document\_text1, document\_text2, ..., read\_only,
type\_of\_link)

Document\_text1, document\_text2,&nbsp;&nbsp;&nbsp;&nbsp;are 1 to 12
arguments that are the names of supporting documents in the form of
text, or arrays or references that contain text.

Read\_only&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the read/write status of the linked worksheet. If read\_only is TRUE,
the sheet can be modified but changes cannot be saved; if FALSE or
omitted, changes to the sheet can be saved. Read\_only applies only to
Microsoft Excel, WKS, and SYLK documents.

Type\_of\_link&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 6 that
specifies what type of link you want to get information about.

|                    |                        |
| ------------------ | ---------------------- |
| **Type\_of\_link** | **Link document type** |
| 1                  | Microsoft Excel link   |
| 2                  | DDE link               |
| 3                  | Reserved               |
| 4                  | Not applicable         |
| 5                  | Subscriber             |
| 6                  | Publisher              |

**Remarks**

You can generate an array of the names of linked workbooks with the
LINKS function.

**Related Functions**

[CHANGE.LINK](CHANGE.LINK.md)&nbsp;&nbsp;&nbsp;Changes supporting workbook links

[GET.LINK.INFO](GET.LINK.INFO.md)&nbsp;&nbsp;&nbsp;Returns information about a link

[LINKS](LINKS.md)&nbsp;&nbsp;&nbsp;Returns the name of all linked workbooks

[UPDATE.LINK](UPDATE.LINK.md)&nbsp;&nbsp;&nbsp;Updates a link to another document



Return to [README](README.md#O)

