NAMES

Returns, as a horizontal array of text, the specified names defined in
the specified workbook. The returned array lists the names in alphabetic
order. Use NAMES instead of LIST.NAMES when you want to return the names
to the macro sheet instead of to the active worksheet.

**Syntax**

**NAMES**(document\_text, type\_num, match\_text)

Document\_text    is text that specifies the workbook whose names you
want returned. If document\_text is omitted, it is assumed to be the
active workbook.

Type\_num    is a number from 1 to 3 that specifies whether to include
hidden names in the returned array.

|                     |                   |
| ------------------- | ----------------- |
| **If type\_num is** | **NAMES returns** |
| 1 or omitted        | Normal names only |
| 2                   | Hidden names only |
| 3                   | All names         |

Match\_text    is text that specifies the names you want returned and
can include wildcard characters. If match\_text is omitted, all names
are returned.

**Remarks**

  - > Hidden names are defined using the DEFINE.NAME macro function and
    > do not appear in the Paste Name, Define Name, or Go To dialog
    > boxes.

  - > NAMES returns a horizontal array, so you will normally enter this
    > function as an array in several horizontal cells or define a name
    > to refer to the array that NAMES returns. If you want the names in
    > a vertical array instead, use the TRANSPOSE function.

  - > You can use the COLUMNS function to count the number of entries in
    > the horizontal array.

>  

**Example**

The following macro formula returns all names on the active workbook
starting with the letter P.

NAMES(, 3, "P\*")

**Related Functions**

[DEFINE.NAME](DEFINE.NAME.md)   Defines a name on the active worksheet or macro sheet

[DELETE.NAME](DELETE.NAME.md)   Deletes a name

[GET.DEF](GET.DEF.md)   Returns a name matching a definition

[GET.NAME](GET.NAME.md)   Returns the definition of a name

[LIST.NAMES](LIST.NAMES.md)   Lists names and their associated information

[SET.NAME](SET.NAME.md)   Defines a name as a value


