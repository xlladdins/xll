SELECT.SPECIAL
==============

Equivalent to clicking the Go To command on the Edit menu and then
selecting the Special button. Use SELECT.SPECIAL to select groups of
similar cells in one of a variety of categories.

**Syntax**

**SELECT.SPECIAL**(**type\_num**, value\_type, levels)

**SELECT.SPECIAL**?(type\_num, value\_type, levels)

Type\_num    is a number from 1 to 13 corresponding to options in the Go
To Special dialog box and describes what to select.

+-----------------+----------------------------------+
| > **Type\_num** | > **Description**                |
+-----------------+----------------------------------+
| > 1             | > Notes/comments                 |
+-----------------+----------------------------------+
| > 2             | > Constants                      |
+-----------------+----------------------------------+
| > 3             | > Formulas                       |
+-----------------+----------------------------------+
| > 4             | > Blanks                         |
+-----------------+----------------------------------+
| > 5             | > Current region                 |
+-----------------+----------------------------------+
| > 6             | > Current array                  |
+-----------------+----------------------------------+
| > 7             | > Row differences                |
+-----------------+----------------------------------+
| > 8             | > Column differences             |
+-----------------+----------------------------------+
| > 9             | > Precedents                     |
+-----------------+----------------------------------+
| > 10            | > Dependents                     |
+-----------------+----------------------------------+
| > 11            | > Last cell                      |
+-----------------+----------------------------------+
| > 12            | > Visible cells only (outlining) |
+-----------------+----------------------------------+
| > 13            | > All objects                    |
+-----------------+----------------------------------+

Value\_type    is a number specifying which types of constants or
formulas you want to select. Value\_type is available only when
type\_num is 2 or 3.

+-------------------+------------------+
| > **Value\_type** | > **Selects**    |
+-------------------+------------------+
| > 1               | > Numbers        |
+-------------------+------------------+
| > 2               | > Text           |
+-------------------+------------------+
| > 4               | > Logical values |
+-------------------+------------------+
| > 16              | > Error values   |
+-------------------+------------------+

These values can be added to select more than one type. The default for
value\_type is 23, which select all value types.

Levels    is a number specifying how precedents and dependents are
selected. Levels is available only when type\_num is 9 or 10. The
default is 1.

+--------------+---------------+
| > **Levels** | > **Selects** |
+--------------+---------------+
| > 1          | > Direct only |
+--------------+---------------+
| > 2          | > All levels  |
+--------------+---------------+

Return to [top](#Q)

SEND.KEYS
