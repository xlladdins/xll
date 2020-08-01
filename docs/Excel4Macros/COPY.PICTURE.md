# COPY.PICTURE

Equivalent to choosing the Copy Picture command from the Edit menu. The
Copy Picture command appears if you hold down SHIFT while choosing the
Edit menu. It copies a chart or range of cells to the Clipboard as a
graphic. Use COPY.PICTURE to create an image of the current selection or
chart for use in another program.

**Syntax**

**COPY.PICTURE**(appearance\_num, size\_num, type\_num)

**COPY.PICTURE**?(appearance\_num, size\_num, type\_num)

**Remarks**

Graphics are created differently on screen and on a printer. Thus, the
printed picture may look different from the one on screen.

Appearance\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number describing how to
copy the picture.

|                     |                                                                                 |
| ------------------- | ------------------------------------------------------------------------------- |
| **Appearance\_num** | **Action**                                                                      |
| 1 or omitted        | Copies a picture as closely as possible to the picture displayed on your screen |
| 2                   | Copies what you would see if you printed the selection                          |

Size\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number describing how to copy the
picture and is only available if the current selection is a chart.

|               |                                                                          |
| ------------- | ------------------------------------------------------------------------ |
| **Size\_num** | **Action**                                                               |
| 1 or omitted  | Copies the chart in the same size as the window on which it is displayed |
| 2             | Copies what you would see if you printed the chart                       |

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the format of
the picture. This argument is available only in Microsoft Excel for
Windows.

|               |                           |
| ------------- | ------------------------- |
| **Type\_num** | **Format of the picture** |
| 1 or omitted  | Picture                   |
| 2             | Bitmap                    |

**Related Functions**

[COPY](COPY.md)&nbsp;&nbsp;&nbsp;Copies and pastes data or objects

[CUT](CUT.md)&nbsp;&nbsp;&nbsp;Cuts or moves data or objects

[PASTE](PASTE.md)&nbsp;&nbsp;&nbsp;Pastes cut or copied data

[PASTE.PICTURE.LINK](PASTE.PICTURE.LINK.md)&nbsp;&nbsp;&nbsp;Pastes a linked picture of the
currently copied area

[PASTE.SPECIAL](PASTE.SPECIAL.md)&nbsp;&nbsp;&nbsp;Pastes specific components of copied data



Return to [README](README.md#C)

# COPY.PICTURE

Equivalent to choosing the Copy Picture command from the Edit menu. The
Copy Picture command appears if you hold down SHIFT while choosing the
Edit menu. It copies a chart or range of cells to the Clipboard as a
graphic. Use COPY.PICTURE to create an image of the current selection or
chart for use in another program.

**Syntax**

**COPY.PICTURE**(appearance\_num, size\_num, type\_num)

**COPY.PICTURE**?(appearance\_num, size\_num, type\_num)

**Remarks**

Graphics are created differently on screen and on a printer. Thus, the
printed picture may look different from the one on screen.

Appearance\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number describing how to
copy the picture.

|                     |                                                                                 |
| ------------------- | ------------------------------------------------------------------------------- |
| **Appearance\_num** | **Action**                                                                      |
| 1 or omitted        | Copies a picture as closely as possible to the picture displayed on your screen |
| 2                   | Copies what you would see if you printed the selection                          |

Size\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number describing how to copy the
picture and is only available if the current selection is a chart.

|               |                                                                          |
| ------------- | ------------------------------------------------------------------------ |
| **Size\_num** | **Action**                                                               |
| 1 or omitted  | Copies the chart in the same size as the window on which it is displayed |
| 2             | Copies what you would see if you printed the chart                       |

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the format of
the picture. This argument is available only in Microsoft Excel for
Windows.

|               |                           |
| ------------- | ------------------------- |
| **Type\_num** | **Format of the picture** |
| 1 or omitted  | Picture                   |
| 2             | Bitmap                    |

**Related Functions**

[COPY](COPY.md)&nbsp;&nbsp;&nbsp;Copies and pastes data or objects

[CUT](CUT.md)&nbsp;&nbsp;&nbsp;Cuts or moves data or objects

[PASTE](PASTE.md)&nbsp;&nbsp;&nbsp;Pastes cut or copied data

[PASTE.PICTURE.LINK](PASTE.PICTURE.LINK.md)&nbsp;&nbsp;&nbsp;Pastes a linked picture of the
currently copied area

[PASTE.SPECIAL](PASTE.SPECIAL.md)&nbsp;&nbsp;&nbsp;Pastes specific components of copied data



Return to [README](README.md#C)

