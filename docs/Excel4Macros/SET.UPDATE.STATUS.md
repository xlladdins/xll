SET.UPDATE.STATUS
=================

Sets the update status of a link to automatic or manual. Use
SET.UPDATE.STATUS to change the way a link is updated.

**Syntax**

**SET.UPDATE.STATUS**(**link\_text, status**, type\_of\_link)

Link\_text    is the path of the linked file for which you want to
change the update status.

Status    is the number 1 or 2 and describes how you want the link to be
updated.

+--------------+---------------------+
| > **Status** | > **Update method** |
+--------------+---------------------+
| > 1          | > Automatic         |
+--------------+---------------------+
| > 2          | > Manual            |
+--------------+---------------------+

Type\_of\_link    is a number from 1 to 4 that specifies what type of
link you want to get information about.

+----------------------+--------------------------+
| > **Type\_of\_link** | > **Link document type** |
+----------------------+--------------------------+
| > 1                  | > Not available          |
+----------------------+--------------------------+
| > 2                  | > DDE/OLE link           |
+----------------------+--------------------------+
| > 3                  | > Not available          |
+----------------------+--------------------------+
| > 4                  | > Not available          |
+----------------------+--------------------------+

**Example**

In Microsoft Excel for Windows, the following macro formula sets the
update status of the DDE link to Microsoft Word for Windows to manual:

SET.UPDATE.STATUS(\"WordDocument\|\'C:\\MEMO.DOC\'!DDE.LINK1\", 2, 2)

**Related Functions**

GET.LINK.INFO   Returns information about a link

UPDATE.LINK   Updates a link to another document

Return to [top](#Q)

SET.VALUE
