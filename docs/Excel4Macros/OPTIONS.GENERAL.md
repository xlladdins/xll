OPTIONS.GENERAL

Equivalent to clicking the Options command on the Tools menu and then
clicking the General tab from the Options dialog box. Sets various
general Microsoft Excel settings.

**Syntax**

**OPTIONS.GENERAL**(R1C1\_mode, dde\_on, sum\_info, tips, recent\_files,
old\_menus, user\_info, font\_name, font\_size, default\_location,
alternate\_location, sheet\_num, enable\_under)

**OPTIONS.GENERAL**?(R1C1\_mode, dde\_on, sum\_info, tips,
recent\_files, old\_menus, user\_info, font\_name, font\_size,
default\_location, alternate\_location, sheet\_num, enable\_under)

Arguments correspond to option buttons, check boxes and text boxes on
the General tab of the Options dialog box. Arguments corresponding to
check boxes are logical values. If an argument is TRUE, the check box is
selected; if FALSE, the check box is cleared; if omitted, the current
setting is not changed.

R1C1\_mode**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp; is a number specifying the reference
style. Use 1 for A1 style references; 2 for R1C1 style references.

Dde\_on**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value corresponding to the
Ignore Other Applications check box, which if TRUE ignores DDE request
from other applications. If FALSE, DDE requests from other applications
are allowed to happen.

Sum\_info**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value corresponding to the
Prompt For Workbook Properties check box, which if TRUE displays the
Summary tab of the Properties dialog box when a workbook is initially
saved. If FALSE, the dialog box is not displayed.

Tips**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value corresponding to the
Reset TipWizard check box in Microsoft Excel 95 or earlier versions,
which, if TRUE, resets the TipWizard. In Microsoft Excel 97 or later,
tips corresponds to the Reset My Tips button in the Office Assistant
dialog box. If FALSE, the TipWizard will not be reset.

Recent\_files**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value corresponding to
the Recently Used File List check box, which if TRUE displays the four
last opened files from the File menu. If FALSE, the file list will not
be displayed.

Old\_menus**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a logical value corresponding to
the Microsoft Excel 4.0 Menus check box in Microsoft Excel version 5.0,
which if TRUE replaces the current menu bar with the Microsoft Excel 4.0
menu bar. If FALSE, the menu bar will not be replaced. This argument is
for compatibility with Microsoft Excel version 5.0 only and is ignored
in later versions. User\_info**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the
Name text box, and is the name of the user of this copy of Microsoft
Excel. By default it is the registered user, but can be changed to work
on a network.

Font\_name**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp; corresponds to the Standard Font text
box, and is the name of the default font.

Font\_size**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Size drop-down edit
box, and is the size of the default font.

Default\_location**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp; corresponds to the Default
File Location text box, and is the default location that the File Open
command displays. The Default is where Microsoft Excel is installed.

Alternate\_location**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Alternate
Startup File Location text box, and is the alternate startup directory.

Sheet\_num**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;corresponds to the Sheets In New
Workbook spin control, and is the number of sheets in a new workbook.
Default is 3. Can go up to 255.

Enable\_under**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;enables underlining of the menus.
Used for only Microsoft Excel for the Macintosh. Ignored in Microsoft
Excel for Windows.

**Related Functions**

[OPTIONS.LISTS.DELETE](OPTIONS.LISTS.DELETE.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Deletes a custom list

[OPTIONS.LISTS.GET](OPTIONS.LISTS.GET.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Returns contents of custom AutoFill
lists

[OPTIONS.VIEW](OPTIONS.VIEW.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Sets various view settings



Return to [README](README.md)

