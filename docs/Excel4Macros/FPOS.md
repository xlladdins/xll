FPOS

Sets the position of a file. The position of a file is where a character
is read from or written to by an FREAD, FREADLN, FWRITE, or FWRITELN
function. Use FPOS when you want to write characters to or read
characters from specific locations. For example, to append text to the
end of a file, you must set the position to the end of the file;
otherwise, you might accidentally overwrite existing characters in the
file.

**Syntax**

**FPOS**(**file\_num**, position\_num)

File\_num&nbsp;&nbsp;&nbsp;&nbsp;is the unique ID number of the file for
which you want to set the position. File\_num is returned by a
previously executed FOPEN function. If file\_num is not valid, FPOS
returns the \#VALUE\! error value.

Position\_num&nbsp;&nbsp;&nbsp;&nbsp;is the location in the file that a
character will be read from or written to.

  - > The first position in a file is 1, the location of the first byte.

  - > The last position in the file is the same as the value returned by
    > FSIZE. For example, the last position in a file with 280 bytes is
    > 280.

  - > If position\_num is omitted, FPOS returns the current position of
    > the file&mdash;that is, the number corresponding to where the next
    > character will be read from or written to.


Whenever you read a character from or write a character to a file, the
file's position is automatically incremented.

**Examples**

The following statement starts a loop that executes until the position
in the open file identified as FileNumber reaches the end of the file:

\=WHILE(FPOS(FileNumber)\<=FSIZE(FileNumber))

**Related Functions**

[FCLOSE](FCLOSE.md)&nbsp;&nbsp;&nbsp;Closes a text file

[FOPEN](FOPEN.md)&nbsp;&nbsp;&nbsp;Opens a file with the type of permission
specified

[FREAD](FREAD.md)&nbsp;&nbsp;&nbsp;Reads characters from a text file

[FREADLN](FREADLN.md)&nbsp;&nbsp;&nbsp;Reads a line from a text file

[FWRITE](FWRITE.md)&nbsp;&nbsp;&nbsp;Writes characters to a text file

[FWRITELN](FWRITELN.md)&nbsp;&nbsp;&nbsp;Writes a line to a text file



Return to [README](README.md)

