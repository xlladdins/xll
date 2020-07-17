MACRO.OPTIONS

Equivalent to clicking the Options button in the Macro dialog box, which
appears when you click the Macros command (Tools menu, Macro submenu).

**Syntax**

MACRO.OPTIONS(**macro\_name**, description, menu\_on, menu\_text,
shortcut\_on, shortcut\_key, function\_category, status\_bar\_text,
help\_id, help\_file)

Macro\_name    is the name of the macro that you want to set options
for, including the name of the workbook and sheet containing the macro.

Description    is the description of the macro displayed in the Macro
dialog box.

Menu\_on    is a logical value indicating whether a menu item is
automatically added for this macro. If TRUE, menu\_text must be
specified. If FALSE or omitted, no menu item is added. If the macro
already has a menu item, setting this argument to FALSE removes the menu
item.

Menu\_text    is the text of the menu item to be added for the macro.
Ignored unless menu\_on is TRUE.

Shortcut\_on    is a logical value indicating whether a shortcut key is
assigned to the macro. If TRUE, shortcut\_key must be specified. If
FALSE or omitted, no shortcut key is assigned. If the macro already has
a shortcut key, setting this argument to FALSE removes the shortcut key.

Shortcut\_key    is the letter of the shortcut key for the macro.
Ignored if shortcut\_key is FALSE.

Function\_category    is the number of the category in the Paste
Function dialog box that the macro is assigned to. Categories are
numbered starting at 1 for the category at the top of the list in the
Paste Function dialog box.

Status\_bar\_text    the text displayed in the status bar when a menu
item or toolbar button assigned to this macro is clicked on. Be sure to
enclose the text in quotes.

Help\_id    is the numerical ID for the help topic associated with this
macro and any related menu items or toolbar buttons.

Help\_file    is the pathname of the help file for the macro.


