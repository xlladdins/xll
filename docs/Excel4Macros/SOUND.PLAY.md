# SOUND.PLAY

This function should not be used in Microsoft Excel 97 or later because
sound notes are available only in Microsoft Excel 95 or earlier
versions.

Plays the sound from a cell note or a file. Equivalent to clicking the
Note command on the Insert menu and clicking the Play button, or
clicking the Note command on the Insert menu, clicking the Import
button, and then opening a file, selecting a sound, and clicking the
Play button. To play sounds in Microsoft Excel for Windows, you must
have a sound board installed in your computer.

**Syntax**

**SOUND.PLAY**(cell\_ref, file\_text, resource)

Cell\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cell note
containing sound that you want to play. If cell\_ref is omitted,
Microsoft Excel plays the sound from the active cell, or from a file if
you specify one.

File\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a file containing
sounds. If cell\_ref is specified, file\_text is ignored.

Resource&nbsp;&nbsp;&nbsp;&nbsp;is a number or name given as text
specifying a sound resource in file\_text that you want to play.

  - > This argument applies only to Microsoft Excel for the Macintosh.

  - > If cell\_ref is specified, resource is ignored.

  - > If resource is omitted, Microsoft Excel uses the first sound
    > resource in the file.

  - > If the file does not contain a sound resource with the specified
    > name or number, Microsoft Excel halts the macro and displays an
    > error message.


**Related Function**

[SOUND.NOTE](SOUND.NOTE.md)&nbsp;&nbsp;&nbsp;Records or imports sound into or erases
sound from cell notes



Return to [README](README.md)

