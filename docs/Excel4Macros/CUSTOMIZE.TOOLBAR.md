CUSTOMIZE.TOOLBAR
=================

Equivalent to choosing the Toolbars command from the View menu and
choosing the Customize button in Microsoft Excel 95. Displays the
Customize Toolbars dialog box. In Microsoft Excel 97 or later, this
function displays the Commands tab on the Customize dialog box. The
Customize dialog box appears when you choose the Toolbars command from
the View menu and then choose the Customize command. This function has a
dialog-box syntax only.

**Syntax**

**CUSTOMIZE.TOOLBAR**?(category)

Category    is a number that specifies which category of tools you want
displayed in the dialog box. If omitted, the previous setting is used.
This argument is for compatibility with Microsoft Excel 95.

  -------------- -----------------------
  **Category**   **Category of tools**
  1              File
  2              Edit
  3              Formula
  4              Formatting
  5              Text Formatting
  6              Drawing
  7              Macro
  8              Charting
  9              Utility
  10             Data
  11             TipWizard
  12             Auditing
  13             Forms
  14             Custom
  -------------- -----------------------

**Related Functions**

ADD.TOOLBAR   Creates a new toolbar with the specified tools

SHOW.TOOLBAR   Hides or displays a toolbar

Return to [top](#A)

CUSTOM.REPEAT
