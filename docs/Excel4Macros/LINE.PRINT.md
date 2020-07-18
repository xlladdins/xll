# LINE.PRINT

Prints the active worksheet using methods compatible with those of Lotus
1-2-3. LINE.PRINT does not use the Microsoft Windows printer drivers.
Unless you have a specific need for the LINE.PRINT function, use the
PRINT function instead.

**Note**&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;&nbsp;&nbsp;&nbsp;**nbsp;This function is only available in Microsoft
Excel for Windows.

**Syntax 1**

Go, Line, Page, Align, and Clear

**LINE.PRINT**(**command**, file, append)

**Syntax 2**

Worksheet settings

**LINE.PRINT**(**command**, setup\_text, leftmarg, rightmarg, topmarg,
botmarg, pglen, formatted)

**Syntax 3**

Global settings

**LINE.PRINT**(**command**, setup\_text, leftmarg, rightmarg, topmarg,
botmarg, pglen, wait, autolf, port, update)

Command&nbsp;&nbsp;&nbsp;&nbsp;is a number corresponding to the command
you want LINE.PRINT to carry out. For syntax 2, command must be 5. For
syntax 3, command must be 6.

|             |                                           |
| ----------- | ----------------------------------------- |
| **Command** | **Command that is carried out**           |
| 1           | Go                                        |
| 2           | Line                                      |
| 3           | Page                                      |
| 4           | Align                                     |
| 5           | Worksheet settings                        |
| 6           | Global settings (saved in EXCEL5.INI)     |
| 7           | Clear (change to current global settings) |

File&nbsp;&nbsp;&nbsp;&nbsp;is the name of a file to which you want to
print. If omitted, Microsoft Excel prints to the printer port determined
by the current global settings.

Append&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether to
append text to file. If TRUE, the file you are printing is appended to
file; if FALSE or omitted, the file you are printing overwrites the
contents of file.

Setup\_text&nbsp;&nbsp;&nbsp;&nbsp;is text that includes a printer
initialization sequence or other control codes to prepare your printer
for printing. If omitted, no setup text is used.

Leftmarg&nbsp;&nbsp;&nbsp;&nbsp;is the size of the left margin measured
in characters from the left side of the page. If omitted, it is assumed
to be 4.

Rightmarg&nbsp;&nbsp;&nbsp;&nbsp;is the size of the right margin
measured in characters from the left side of the page. If omitted, it is
assumed to be 76.

Topmarg&nbsp;&nbsp;&nbsp;&nbsp;is the size of the top margin measured in
lines from the top of the page. If omitted, it is assumed to be 2.

Botmarg&nbsp;&nbsp;&nbsp;&nbsp;is the size of the bottom margin measured
in lines from the bottom of the page. If omitted, it is assumed to be 2.

Pglen&nbsp;&nbsp;&nbsp;&nbsp;is the number of lines on one page. If
omitted, it is assumed to be 66 (11 inches with 6 lines per inch). If
you're using an HP LaserJet or compatible printer, set pglen to 60 (the
printer reserves six lines).

Formatted&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether
to format the output. If TRUE or omitted, the output is formatted; if
FALSE, it is not formatted.

Wait&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether to
wait after printing a page. If TRUE, Microsoft Excel waits; if FALSE or
omitted, Microsoft Excel continues printing.

Autolf&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether your
printer has automatic line feeding. If TRUE, Microsoft Excel prints
lines normally; if FALSE or omitted, Microsoft Excel sends an additional
line feed character after printing each line.

Port&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 8 specifying which
port to use when printing.

|              |                             |
| ------------ | --------------------------- |
| **Port**     | **Port used when printing** |
| 1 or omitted | LPT1                        |
| 2            | COM1                        |
| 3            | LPT2                        |
| 4            | COM2                        |
| 5            | LPT1                        |
| 6            | LPT2                        |
| 7            | LPT3                        |
| 8            | LPT4                        |

Update&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether to
update and save global settings. If TRUE, the settings are saved in the
EXCEL5.INI file; if FALSE or omitted, the global settings are not saved.

**Remarks**

The default values for print settings on your worksheet are determined
by the current global settings.

**Example**

The following macro formula prints the currently defined print area to
the currently defined printer port:

LINE.PRINT(1)

**Related Function**

[PRINT](PRINT.md)&nbsp;&nbsp;&nbsp;Prints the active sheet



Return to [README](README.md)

