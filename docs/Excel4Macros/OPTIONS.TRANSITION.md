OPTIONS.TRANSITION

Equivalent to clicking the Options command on the Tools menu and then
clicking the Transition tab in the Options dialog box. Sets options
relating to compatibility with other spreadsheets.

**Syntax**

**OPTIONS.TRANSITION**(menu\_key, menu\_key\_action, nav\_keys,
trans\_eval, trans\_entry)

**OPTIONS.TRANSITION**?(menu\_key, menu\_key\_action, nav\_keys,
trans\_eval, trans\_entry)

Menu\_key&nbsp;&nbsp;&nbsp;&nbsp;is text specifying which alternate menu
key to use.

Menu\_key\_action&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 specifying
options for the alternate menu or Help key. In Microsoft Excel for the
Macintosh, menu\_key\_action is ignored.

|                       |                                          |
| --------------------- | ---------------------------------------- |
| **Menu\_key\_action** | **Alternate menu or Help key activates** |
| 1 or omitted          | Microsoft Excel menus                    |
| 2                     | Lotus 1-2-3 Help                         |

Nav\_keys&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds to
the Transition Navigation Keys check box, which if TRUE uses alternate
navigation keys that correspond to the navigation keys for Lotus 1-2-3.
In Microsoft Excel for the Macintosh, nav\_keys is ignored.

Trans\_eval&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds
to the Transition Formula Evaluation check box.

  - > If trans\_eval is TRUE, Microsoft Excel uses a set of rules
    > compatible with that of Lotus 1-2-3 when calculating formulas.
    > Text is treated as 0. TRUE and FALSE are treated as 1 and 0.
    > Certain characters in database criteria ranges are interpreted the
    > same way Lotus 1-2-3 interprets them.

  - > If trans\_eval is FALSE or omitted, Microsoft Excel calculates
    > normally.

Trans\_entry&nbsp;&nbsp;&nbsp;&nbsp;is a logical value that corresponds
to the Transition Formula Entry check box.

  - > This argument is available only in Microsoft Excel for Windows.

  - > If trans\_entry is TRUE, Microsoft Excel accepts formulas entered
    > in  
    > Lotus 1-2-3 style.

  - > If trans\_entry is FALSE or omitted, Microsoft Excel only accepts
    > formulas entered in Microsoft Excel style.

**Related Functions**

[OPTIONS.LISTS.DELETE](OPTIONS.LISTS.DELETE.md)&nbsp;&nbsp;&nbsp;Deletes a custom list

[OPTIONS.LISTS.GET](OPTIONS.LISTS.GET.md)&nbsp;&nbsp;&nbsp;Returns contents of custom AutoFill
lists

[OPTIONS.VIEW](OPTIONS.VIEW.md)&nbsp;&nbsp;&nbsp;Sets various view settings



Return to [README](README.md)

