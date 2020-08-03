# SLIDE.PASTE

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Paste button on a slide show sheet. Pastes
the contents of the Clipboard as the next available slide of the active
slide show sheet, and gives the slide the attributes you specify.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.PASTE**(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

**SLIDE.PASTE**?(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

Effect\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the transition
effect you want to use when displaying the slide.

  - > The numbers correspond to the effects in the Effect list in the
    > Edit Slide dialog box. The first effect in the list is 1 (None).

  - > If effect\_num is omitted, the default setting is used.


Speed\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 10 specifying
the speed of the transition effect.

  - > If speed\_num is omitted, the default setting is used.

  - > If speed\_num is greater than 10, Microsoft Excel uses the value
    > 10 anyway.

  - > If effect\_num is 1 (none), speed\_num is ignored.


Advance\_rate\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying how
long (in seconds) the slide is displayed before advancing to the next
one.

  - > If advance\_rate\_num is omitted, the default setting is used.

  - > If advance\_rate\_num is 0, you must press a key or click with the
    > mouse to advance to the next slide.


Soundfile\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a file enclosed in
quotation marks and specifies sound that will be played when the slide
is displayed.

  - > If soundfile\_text is omitted, Microsoft Excel plays the default
    > sound defined for the slide show sheet, if any.

  - > If soundfile\_text is empty text (""), no sound is played.

  - > In Microsoft Excel for the Macintosh, soundfile\_text also
    > includes the number or name of the sound resource to play in the
    > file.


Resource&nbsp;&nbsp;&nbsp;&nbsp;is the number or name of a sound
resource in soundfile\_text.

  - > This argument applies only to Microsoft Excel for the Macintosh.

  - > If resource is omitted, Microsoft Excel uses the first resource in
    > the file.

  - > If the file does not contain a sound resource with the specified
    > name or number, Microsoft Excel halts the macro and displays an
    > error message.


**Remarks**

  - > SLIDE.PASTE returns TRUE if it successfully pastes the slide, or
    > FALSE if you click the Cancel button when using the dialog-box
    > form. If the active sheet is not a slide show or is protected,
    > SLIDE.PASTE returns the \#N/A error value. If the Clipboard format
    > is not compatible with the slide show sheet's format, SLIDE.PASTE
    > returns the \#VALUE error value.


**Examples**

In Microsoft Excel for Windows, the following macro formula pastes the
contents of the Clipboard into the active slide show sheet. The slide's
transition effect is fade, at a speed of 8; it is displayed for five
seconds; and Microsoft Excel plays the specified sound file:

SLIDE.PASTE(3, 8, 5, "C:\\SLIDES\\SOUND\\MACHINES.WAV")

In Microsoft Excel for the Macintosh, the formula is:

SLIDE.PASTE(3, 8, 5, "HARD DISK:SLIDES:SOUND:MACHINE SOUNDS")



Return to [README](README.md#S)

# SLIDE.PASTE

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Paste button on a slide show sheet. Pastes
the contents of the Clipboard as the next available slide of the active
slide show sheet, and gives the slide the attributes you specify.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.PASTE**(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

**SLIDE.PASTE**?(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

Effect\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the transition
effect you want to use when displaying the slide.

  - > The numbers correspond to the effects in the Effect list in the
    > Edit Slide dialog box. The first effect in the list is 1 (None).

  - > If effect\_num is omitted, the default setting is used.


Speed\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 10 specifying
the speed of the transition effect.

  - > If speed\_num is omitted, the default setting is used.

  - > If speed\_num is greater than 10, Microsoft Excel uses the value
    > 10 anyway.

  - > If effect\_num is 1 (none), speed\_num is ignored.


Advance\_rate\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying how
long (in seconds) the slide is displayed before advancing to the next
one.

  - > If advance\_rate\_num is omitted, the default setting is used.

  - > If advance\_rate\_num is 0, you must press a key or click with the
    > mouse to advance to the next slide.


Soundfile\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a file enclosed in
quotation marks and specifies sound that will be played when the slide
is displayed.

  - > If soundfile\_text is omitted, Microsoft Excel plays the default
    > sound defined for the slide show sheet, if any.

  - > If soundfile\_text is empty text (""), no sound is played.

  - > In Microsoft Excel for the Macintosh, soundfile\_text also
    > includes the number or name of the sound resource to play in the
    > file.


Resource&nbsp;&nbsp;&nbsp;&nbsp;is the number or name of a sound
resource in soundfile\_text.

  - > This argument applies only to Microsoft Excel for the Macintosh.

  - > If resource is omitted, Microsoft Excel uses the first resource in
    > the file.

  - > If the file does not contain a sound resource with the specified
    > name or number, Microsoft Excel halts the macro and displays an
    > error message.


**Remarks**

  - > SLIDE.PASTE returns TRUE if it successfully pastes the slide, or
    > FALSE if you click the Cancel button when using the dialog-box
    > form. If the active sheet is not a slide show or is protected,
    > SLIDE.PASTE returns the \#N/A error value. If the Clipboard format
    > is not compatible with the slide show sheet's format, SLIDE.PASTE
    > returns the \#VALUE error value.


**Examples**

In Microsoft Excel for Windows, the following macro formula pastes the
contents of the Clipboard into the active slide show sheet. The slide's
transition effect is fade, at a speed of 8; it is displayed for five
seconds; and Microsoft Excel plays the specified sound file:

SLIDE.PASTE(3, 8, 5, "C:\\SLIDES\\SOUND\\MACHINES.WAV")

In Microsoft Excel for the Macintosh, the formula is:

SLIDE.PASTE(3, 8, 5, "HARD DISK:SLIDES:SOUND:MACHINE SOUNDS")



Return to [README](README.md#S)

# SLIDE.PASTE

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Equivalent to clicking the Paste button on a slide show sheet. Pastes
the contents of the Clipboard as the next available slide of the active
slide show sheet, and gives the slide the attributes you specify.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.PASTE**(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

**SLIDE.PASTE**?(effect\_num, speed\_num, advance\_rate\_num,
soundfile\_text)

Effect\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the transition
effect you want to use when displaying the slide.

  - > The numbers correspond to the effects in the Effect list in the
    > Edit Slide dialog box. The first effect in the list is 1 (None).

  - > If effect\_num is omitted, the default setting is used.


Speed\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number from 1 to 10 specifying
the speed of the transition effect.

  - > If speed\_num is omitted, the default setting is used.

  - > If speed\_num is greater than 10, Microsoft Excel uses the value
    > 10 anyway.

  - > If effect\_num is 1 (none), speed\_num is ignored.


Advance\_rate\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying how
long (in seconds) the slide is displayed before advancing to the next
one.

  - > If advance\_rate\_num is omitted, the default setting is used.

  - > If advance\_rate\_num is 0, you must press a key or click with the
    > mouse to advance to the next slide.


Soundfile\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of a file enclosed in
quotation marks and specifies sound that will be played when the slide
is displayed.

  - > If soundfile\_text is omitted, Microsoft Excel plays the default
    > sound defined for the slide show sheet, if any.

  - > If soundfile\_text is empty text (""), no sound is played.

  - > In Microsoft Excel for the Macintosh, soundfile\_text also
    > includes the number or name of the sound resource to play in the
    > file.


Resource&nbsp;&nbsp;&nbsp;&nbsp;is the number or name of a sound
resource in soundfile\_text.

  - > This argument applies only to Microsoft Excel for the Macintosh.

  - > If resource is omitted, Microsoft Excel uses the first resource in
    > the file.

  - > If the file does not contain a sound resource with the specified
    > name or number, Microsoft Excel halts the macro and displays an
    > error message.


**Remarks**

  - > SLIDE.PASTE returns TRUE if it successfully pastes the slide, or
    > FALSE if you click the Cancel button when using the dialog-box
    > form. If the active sheet is not a slide show or is protected,
    > SLIDE.PASTE returns the \#N/A error value. If the Clipboard format
    > is not compatible with the slide show sheet's format, SLIDE.PASTE
    > returns the \#VALUE error value.


**Examples**

In Microsoft Excel for Windows, the following macro formula pastes the
contents of the Clipboard into the active slide show sheet. The slide's
transition effect is fade, at a speed of 8; it is displayed for five
seconds; and Microsoft Excel plays the specified sound file:

SLIDE.PASTE(3, 8, 5, "C:\\SLIDES\\SOUND\\MACHINES.WAV")

In Microsoft Excel for the Macintosh, the formula is:

SLIDE.PASTE(3, 8, 5, "HARD DISK:SLIDES:SOUND:MACHINE SOUNDS")



Return to [README](README.md#S)

