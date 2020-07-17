SLIDE.GET
=========

This function should not be used in Microsoft Excel 95 or later because
the Slide Show add-in is available only in Microsoft Excel version 5.0
or earlier versions.

Returns the specified information about a slide show or a specific slide
in the slide show.

If this function is not available, you must install the Slide Show
add-in.

**Syntax**

**SLIDE.GET**(**type\_num**, name\_text, slide\_num)

Type\_num    is a number specifying the type of information you want.

These values of type\_num return information about a slide show.

+-----------------+---------------------------------------------------+
| > **Type\_num** | > **Type of information**                         |
+-----------------+---------------------------------------------------+
| > 1             | > Number of slides in the slide show              |
+-----------------+---------------------------------------------------+
| > 2             | > A two-item horizontal array containing the      |
|                 | > numbers of the first and last slides in the     |
|                 | > current selection, or the \#VALUE error value   |
|                 | > if the selection is nonadjacent                 |
+-----------------+---------------------------------------------------+
| > 3             | > Version number of the Slide Show add-in that    |
|                 | > created the slide show sheet                    |
+-----------------+---------------------------------------------------+

These values of type\_num return information about a specific slide in
the slide show.

+-----------------+---------------------------------------------------+
| > **Type\_num** | > **Type of information**                         |
+-----------------+---------------------------------------------------+
| > 4             | > Transition effect number                        |
+-----------------+---------------------------------------------------+
| > 5             | > Transition effect name                          |
+-----------------+---------------------------------------------------+
| > 6             | > Transition effect speed                         |
+-----------------+---------------------------------------------------+
| > 7             | > Number of seconds the slide is displayed before |
|                 | > advancing                                       |
+-----------------+---------------------------------------------------+
| > 8             | > Name of the sound file associated with the      |
|                 | > slide, or empty text (\"\") if none is          |
|                 | > specified (in Microsoft Excel for the           |
|                 | > Macintosh, this includes the number or name of  |
|                 | > the sound resource within the sound file)       |
+-----------------+---------------------------------------------------+

Name\_text    is the name of an open slide show sheet for which you want
information. If name\_text is omitted, it is assumed to be the active
sheet.

Slide\_num    is the number of the slide about which you want
information.

-   If slide\_num is omitted, it is assumed to be the slide associated
    > with the active cell on the sheet specified by name\_text.

-   If type\_num is 1 through 3, slide\_num is ignored.

Return to [top](#Q)

SLIDE.PASTE
