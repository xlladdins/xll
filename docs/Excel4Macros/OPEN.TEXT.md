# OPEN.TEXT

Equivalent to using the Text Import Wizard to open a text file in
Microsoft Excel.

**Syntax**

**OPEN.TEXT**(file\_name, file\_origin, start\_row, file\_type,
text\_qualifier, consecutive\_delim, tab, semicolon, comma, space,
other, other\_char, {field\_info1; field\_info2;...})

File\_name&nbsp;&nbsp;&nbsp;&nbsp;is the full pathname of the text file
you want to open.

File\_origin&nbsp;&nbsp;&nbsp;&nbsp;specifies the operating environment
the text file was created in.

|                  |                               |
| ---------------- | ----------------------------- |
| **File\_origin** | **Operating system**          |
| 1                | Macintosh                     |
| 2                | Windows (ANSI)                |
| 3                | MS DOS (PC-8)                 |
| Omitted          | Current operating environment |

Start\_row&nbsp;&nbsp;&nbsp;&nbsp;is a number greater than or equal to
one, specifying the row in the text file where you want to start
importing into Microsoft Excel. This number should be less than the
number of lines in the text file.

File\_type&nbsp;&nbsp;&nbsp;&nbsp;specifies the type of delimited text
file to import:

|                |                  |
| -------------- | ---------------- |
| **File\_type** | **Type of file** |
| 1 or omitted   | Delimited        |
| 2              | Fixed width      |

Text\_qualifier&nbsp;&nbsp;&nbsp;&nbsp;indicates the character-enclosing
text fields in the text file:

|                           |                           |
| ------------------------- | ------------------------- |
| **Text\_qualifier value** | **Qualifier**             |
| 1 or "                    | " (double quotation mark) |
| 2 or '                    | ' (single quotation mark) |
| 3 or {None}               | No text qualifier         |

Consecutive\_delim&nbsp;&nbsp;&nbsp;&nbsp;is a logical value
corresponding to the Treat Consecutive Delimiters as One check box
which, if TRUE, allows consecutive delimiters (such as ",,,") to be
treated as a single delimiter. If FALSE, all consecutive delimiters are
considered separate column breaks.

Tab, semicolon, comma, space&nbsp;&nbsp;&nbsp;&nbsp;are logical values
corresponding to the check boxes in the Column Delimiters group. If an
argument is TRUE, the check box is selected. These arguments apply only
when the file\_type argument is 1 or omitted (delimited file type).

Other&nbsp;&nbsp;&nbsp;&nbsp;is a logical value indicating a custom
delimiter to use in opening the text file.

Other\_char&nbsp;&nbsp;&nbsp;&nbsp;specifies the custom delimiter to use
or FALSE if no custom delimiter is used.

Field\_info&nbsp;&nbsp;&nbsp;&nbsp;is an array which consists of the
following elements: "column number, data\_format", if file\_type is 1;
or "start\_pos, data\_format" if file\_type is 2.

**Related Function**

[TEXT.TO.COLUMNS](TEXT.TO.COLUMNS.md)&nbsp;&nbsp;&nbsp;Parses text, as in a text file, into
columns of data



Return to [README](README.md)

