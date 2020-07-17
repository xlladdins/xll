FWRITELN

Writes text, followed by a carriage return and linefeed, to a file,
starting at the current position in that file. (For more information
about a file's position, see FPOS.) If FWRITELN can't write to the file,
it returns the \#N/A error value. Use FWRITELN instead of FWRITE when
you want to append a carriage return and linefeed to each group of
characters that you write to a text file.

**Syntax**

**FWRITELN**(**file\_num, text**)

File\_num    is the unique ID number of the file you want to write data
to. File\_num is returned by a previously executed FOPEN function. If
file\_num is not valid, FWRITELN returns the \#VALUE\! error value.

Text    is the text you want to write to the file.

**Remarks**

In Microsoft Excel for Windows, FWRITELN writes text followed by a
carriage return and a line feed. In Microsoft Excel for the Macintosh,
FWRITELN writes text followed by a carriage return only.

**Example**

The following function writes the current month to the open file
identified as FileNumber and starts a new line in the file:

FWRITELN(FileNumber, TEXT(MONTH(NOW()),"mmmm"))

**Related Functions**

[FOPEN](FOPEN.md)   Opens a file with the type of permission specified

[FPOS](FPOS.md)   Sets the position in a text file

[FREAD](FREAD.md)   Reads characters from a text file

[FWRITE](FWRITE.md)   Writes characters to a text file



Return to [README.md](README.md)

FWRITELN

Writes text, followed by a carriage return and linefeed, to a file,
starting at the current position in that file. (For more information
about a file's position, see FPOS.) If FWRITELN can't write to the file,
it returns the \#N/A error value. Use FWRITELN instead of FWRITE when
you want to append a carriage return and linefeed to each group of
characters that you write to a text file.

**Syntax**

**FWRITELN**(**file\_num, text**)

File\_num    is the unique ID number of the file you want to write data
to. File\_num is returned by a previously executed FOPEN function. If
file\_num is not valid, FWRITELN returns the \#VALUE\! error value.

Text    is the text you want to write to the file.

**Remarks**

In Microsoft Excel for Windows, FWRITELN writes text followed by a
carriage return and a line feed. In Microsoft Excel for the Macintosh,
FWRITELN writes text followed by a carriage return only.

**Example**

The following function writes the current month to the open file
identified as FileNumber and starts a new line in the file:

FWRITELN(FileNumber, TEXT(MONTH(NOW()),"mmmm"))

**Related Functions**

[FOPEN](FOPEN.md)   Opens a file with the type of permission specified

[FPOS](FPOS.md)   Sets the position in a text file

[FREAD](FREAD.md)   Reads characters from a text file

[FWRITE](FWRITE.md)   Writes characters to a text file



Return to [README.md](README.md)

