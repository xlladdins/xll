# SOUND.NOTE

This function should not be used in Microsoft Excel 97 or later because
sound notes are available only in Microsoft Excel 95 or earlier
versions.

Records sound into or erases sound from a cell note or imports sound
from another file into a cell note. This function requires that you have
recording hardware installed in your computer, and you must be running
Microsoft Windows version 3.1 or later, or Apple system software version
6.07 or later.

**Syntax 1**

Recording or erasing sound

**SOUND.NOTE**(cell\_ref, erase\_snd)

**Syntax 2**

Importing sound from another file

**SOUND.NOTE**(cell\_ref, file\_text, resource)

Cell\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cell containing a
note into which you want to record or import sounds or from which you
want to erase a sound.

Erase\_snd&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether
to erase the sound in the note. If erase\_snd is TRUE, Microsoft Excel
erases only the sound from the note. If FALSE or omitted, Microsoft
Excel displays the Record dialog box so that you can record sound into
the note.

File\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a file containing
sounds.

Resource&nbsp;&nbsp;&nbsp;&nbsp;is the number or name of a sound
resource in file\_text that you want to import into your note.

  - > This argument applies only to Microsoft Excel for the Macintosh.

  - > If resource is omitted, Microsoft Excel uses the first resource in
    > the file.

  - > If the file does not contain a sound resource with the specified
    > name or number, Microsoft Excel halts the macro and displays an
    > error message.

**Remarks**

  - > To find out if a cell has sound attached to it, use GET.CELL(47).

  - > Sounds notes are not available in Microsoft Excel 97 or later.

**Examples**

The following macro formula erases the sound, if present, from cell A1
on the active sheet:

SOUND.NOTE(\!$A$1, TRUE)

The following macro formula displays the Record dialog box so that you
can record sound into a note for cell A1 on the active sheet:

SOUND.NOTE(\!$A$1)

In Microsoft Excel for Windows, the following macro formula imports the
sound from a file named CHIMES.WAV into a note for the cell named
Doorbell on the active sheet:

SOUND.NOTE(\!Doorbell, "C:\\SOUNDS\\CHIMES.WAV")

In Microsoft Excel for the Macintosh, the following macro formula
imports a sound called Chimes from a file named SOFT SOUNDS into a note
for the cell named Doorbell on the active sheet:

SOUND.NOTE(\!Doorbell, "HARD DISK:SOUNDS:SOFT SOUNDS", "Chimes")

**Related Functions**

[NOTE](NOTE.md)&nbsp;&nbsp;&nbsp;Creates or changes a cell note

[SOUND.PLAY](SOUND.PLAY.md)&nbsp;&nbsp;&nbsp;Plays the sound from a cell note or a file



Return to [README](README.md#S)

# SOUND.NOTE

This function should not be used in Microsoft Excel 97 or later because
sound notes are available only in Microsoft Excel 95 or earlier
versions.

Records sound into or erases sound from a cell note or imports sound
from another file into a cell note. This function requires that you have
recording hardware installed in your computer, and you must be running
Microsoft Windows version 3.1 or later, or Apple system software version
6.07 or later.

**Syntax 1**

Recording or erasing sound

**SOUND.NOTE**(cell\_ref, erase\_snd)

**Syntax 2**

Importing sound from another file

**SOUND.NOTE**(cell\_ref, file\_text, resource)

Cell\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to the cell containing a
note into which you want to record or import sounds or from which you
want to erase a sound.

Erase\_snd&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying whether
to erase the sound in the note. If erase\_snd is TRUE, Microsoft Excel
erases only the sound from the note. If FALSE or omitted, Microsoft
Excel displays the Record dialog box so that you can record sound into
the note.

File\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a file containing
sounds.

Resource&nbsp;&nbsp;&nbsp;&nbsp;is the number or name of a sound
resource in file\_text that you want to import into your note.

  - > This argument applies only to Microsoft Excel for the Macintosh.

  - > If resource is omitted, Microsoft Excel uses the first resource in
    > the file.

  - > If the file does not contain a sound resource with the specified
    > name or number, Microsoft Excel halts the macro and displays an
    > error message.

**Remarks**

  - > To find out if a cell has sound attached to it, use GET.CELL(47).

  - > Sounds notes are not available in Microsoft Excel 97 or later.

**Examples**

The following macro formula erases the sound, if present, from cell A1
on the active sheet:

SOUND.NOTE(\!$A$1, TRUE)

The following macro formula displays the Record dialog box so that you
can record sound into a note for cell A1 on the active sheet:

SOUND.NOTE(\!$A$1)

In Microsoft Excel for Windows, the following macro formula imports the
sound from a file named CHIMES.WAV into a note for the cell named
Doorbell on the active sheet:

SOUND.NOTE(\!Doorbell, "C:\\SOUNDS\\CHIMES.WAV")

In Microsoft Excel for the Macintosh, the following macro formula
imports a sound called Chimes from a file named SOFT SOUNDS into a note
for the cell named Doorbell on the active sheet:

SOUND.NOTE(\!Doorbell, "HARD DISK:SOUNDS:SOFT SOUNDS", "Chimes")

**Related Functions**

[NOTE](NOTE.md)&nbsp;&nbsp;&nbsp;Creates or changes a cell note

[SOUND.PLAY](SOUND.PLAY.md)&nbsp;&nbsp;&nbsp;Plays the sound from a cell note or a file



Return to [README](README.md#S)

