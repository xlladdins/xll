INPUT

Displays a dialog box for user input. Returns the information entered in
the dialog box. Use INPUT to display a simple dialog box for the user to
enter information to be used in a macro.

The dialog box has an OK and a Cancel button. If you click the OK
button, INPUT returns the default value specified or the value typed in
the edit box. If you click the Cancel button, INPUT returns FALSE.

**Syntax**

**INPUT**(**message\_text**, type\_num, title\_text, default, x\_pos,
y\_pos, help\_ref)

Message\_text&nbsp;&nbsp;&nbsp;&nbsp;is the text to be displayed in the
dialog box. Message\_text must be enclosed in quotation marks.

Type\_num&nbsp;&nbsp;&nbsp;&nbsp;is a number specifying the type of data
to be entered.

|               |               |
| ------------- | ------------- |
| **Type\_num** | **Data type** |
| 0             | Formula       |
| 1             | Number        |
| 2             | Text          |
| 4             | Logical       |
| 8             | Reference     |
| 16            | Error         |
| 64            | Array         |

You can also use the sum of the allowable data types for type\_num. For
example, for an input box that can accept formulas, text, or numbers,
set type\_num equal to 3 (the sum of 0, 1, and 2, which are the type
specifiers for formula, number, and text). If type\_num is omitted, it
is assumed to be 2.

  - > If type\_num is 0, INPUT returns the formula in the form of text,
    > for example, "=2\*PI()/360".

  - > To enter a formula, include an equal sign at the beginning of the
    > formula; otherwise the formula is returned as text.

  - > If the formula contains references, they are returned as
    > R1C1-style references, for example, "=RC\[-1\]\*(1+R1C1)".

  - > If type\_num is 8, INPUT returns an absolute reference to the
    > specified cells.

  - > If you enter a single-cell reference in the dialog box, the value
    > in that cell is returned by the INPUT function.

  - > If the information entered in the dialog box is not of the correct
    > data type, Microsoft Excel attempts to convert it to the specified
    > type. If the information can't be converted, Microsoft Excel
    > displays an error message.


Title\_text&nbsp;&nbsp;&nbsp;&nbsp;is text specifying a title to be
displayed in the title bar of the dialog box. If title\_text is omitted,
it is assumed to be "Input".

Default&nbsp;&nbsp;&nbsp;&nbsp;specifies a value to be shown in the edit
box when the dialog box is initially displayed. If default is omitted,
the edit box is left empty.

X\_pos, y\_pos&nbsp;&nbsp;&nbsp;&nbsp;specify the horizontal and
vertical position, in points, of the dialog box. A point is 1/72nd of an
inch. If either or both arguments are omitted, the dialog box is
centered in the corresponding direction.

Help\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a reference to a custom online Help
topic in a text file, in the form "filename\!topic\_number".

  - > If help\_ref is present, a Help button appears in the lower-right
    > corner of the dialog box. Clicking the Help button starts Help and
    > displays the specified topic.

  - > If help\_ref is omitted, no Help button appears.

  - > Help\_ref must be given as text.


For more information about custom Help topics, see HELP.

**Remarks**

Relative references entered in formulas in the INPUT dialog box are
relative to the active cell at the time the INPUT function is
calculated. If you are using the reference entered into the dialog box
in a cell other than the active cell, it may not refer to the cells you
intend it to. For example, if the active cell is A3 and you enter the
formula "=A1+A2" in an INPUT dialog box, intending to add the values in
those cells, and then use the FORMULA function to enter the formula in
cell B3, the formula in cell B3 will read "=B1+B2" because you gave a
relative reference. You can use FORMULA.CONVERT to solve this problem.

**Examples**

In Microsoft Excel for Windows, the following macro formula displays the
following dialog box:

INPUT("Enter the inflation rate:", 1, "Inflation Rate", , , ,
"CUSTHELP.DOC\!101")

If you then enter 12%, INPUT returns the value 0.12.

In Microsoft Excel for the Macintosh, the following macro formula
displays the following dialog box:

INPUT("Enter the inflation rate:", 1, "Inflation Rate", , , , "CUSTOM
HELP\!101")

If you then enter 12%, INPUT returns the value 0.12.

If the active cell is C2 and you enter the formula =B2\*(1+$A$1) in
response to the following macro formula:

INPUT("Enter your monthly increase formula:", 0)

INPUT returns "=RC\[-1\]\*(1+R1C1)"

If you select the range $A$2:$A$8 in the INPUT dialog box:

REFTEXT|USA|002|001|001|common|UREFTEXT(INPUT("Please make your
selection.", 8)) returns R2C1:R8C1

**Related Functions**

[ALERT](ALERT.md)&nbsp;&nbsp;&nbsp;Displays a dialog box and a message

[DIALOG.BOX](DIALOG.BOX.md)&nbsp;&nbsp;&nbsp;Displays a custom dialog box

[FORMULA.CONVERT](FORMULA.CONVERT.md)&nbsp;&nbsp;&nbsp;Changes the style and type of
references in a formula

[HELP](HELP.md)&nbsp;&nbsp;&nbsp;Displays a custom Help topic



Return to [README](README.md)

