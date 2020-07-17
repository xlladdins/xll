SCENARIO.ADD

Equivalent to clicking the Scenarios command on the Tools menu and then
clicking the Add button. Defines the specified values as a scenario. A
scenario is a set of values to be used as input for a model on your
worksheet.

**Syntax**

**SCENARIO.ADD**(**scen\_name**, value\_array, changing\_ref,
scen\_comment, locked, hidden)

Scen\_name    is the name of the scenario you want to define.

Value\_array    is a horizontal array of values you want to use as input
for the model on your worksheet.

  - > Any entry that would be valid for a cell in your model can be a
    > value in value\_array.

  - > The values must be arranged in the same order as the model's
    > changing cells. The changing cells are listed in the Changing
    > Cells box in the Scenario Manager dialog box.

  - > If value\_array is omitted, it is assumed to contain the current
    > values of the changing cells.

>  

Changing\_ref    is a reference to cells you want to define as changing
cells for a scenario.

  - > If omitted, uses the changing cells for the last scenario defined
    > for the sheet.

  - > If changing\_ref contains nonadjacent references, you must
    > separate the reference areas by commas (or other list separator).
    > If you are using A1-style references, then you must enclose
    > reference in an extra set of parentheses.

>  

Scen\_comment    is text specifying a descriptive comment for the
scenario defined by scen\_name.

Locked    is a logical value that corresponds to the Prevent Changes
check box in the Add or Edit Scenario dialogs boxes. If TRUE or omitted
, prevents users from changing values in a scenario. If FALSE, users are
allowed to make changes to the scenario. The locking will not become
enabled until the sheet is protected with the Protect Sheet command from
the Protection submenu on the Tools menu.

Hidden    is a logical value that corresponds to the Hide check box in
the Add or Edit Scenario dialog boxes. If TRUE, the scenario will be
hidden from view from the users and will not appear in the Scenario
Manager dialog box. If FALSE or omitted, the scenario will remain
unhidden. The scenario will not become hidden until the sheet is
protected with the Protect Sheet command from the Protection submenu on
the Tools menu.

**Related Functions**

[REPORT.DEFINE](REPORT.DEFINE.md)   Creates a report

[SCENARIO.GET](SCENARIO.GET.md)   Returns the specified information about the scenarios
defined on your worksheet



Return to [README.md](README.md)

SCENARIO.ADD

Equivalent to clicking the Scenarios command on the Tools menu and then
clicking the Add button. Defines the specified values as a scenario. A
scenario is a set of values to be used as input for a model on your
worksheet.

**Syntax**

**SCENARIO.ADD**(**scen\_name**, value\_array, changing\_ref,
scen\_comment, locked, hidden)

Scen\_name    is the name of the scenario you want to define.

Value\_array    is a horizontal array of values you want to use as input
for the model on your worksheet.

  - > Any entry that would be valid for a cell in your model can be a
    > value in value\_array.

  - > The values must be arranged in the same order as the model's
    > changing cells. The changing cells are listed in the Changing
    > Cells box in the Scenario Manager dialog box.

  - > If value\_array is omitted, it is assumed to contain the current
    > values of the changing cells.

>  

Changing\_ref    is a reference to cells you want to define as changing
cells for a scenario.

  - > If omitted, uses the changing cells for the last scenario defined
    > for the sheet.

  - > If changing\_ref contains nonadjacent references, you must
    > separate the reference areas by commas (or other list separator).
    > If you are using A1-style references, then you must enclose
    > reference in an extra set of parentheses.

>  

Scen\_comment    is text specifying a descriptive comment for the
scenario defined by scen\_name.

Locked    is a logical value that corresponds to the Prevent Changes
check box in the Add or Edit Scenario dialogs boxes. If TRUE or omitted
, prevents users from changing values in a scenario. If FALSE, users are
allowed to make changes to the scenario. The locking will not become
enabled until the sheet is protected with the Protect Sheet command from
the Protection submenu on the Tools menu.

Hidden    is a logical value that corresponds to the Hide check box in
the Add or Edit Scenario dialog boxes. If TRUE, the scenario will be
hidden from view from the users and will not appear in the Scenario
Manager dialog box. If FALSE or omitted, the scenario will remain
unhidden. The scenario will not become hidden until the sheet is
protected with the Protect Sheet command from the Protection submenu on
the Tools menu.

**Related Functions**

[REPORT.DEFINE](REPORT.DEFINE.md)   Creates a report

[SCENARIO.GET](SCENARIO.GET.md)   Returns the specified information about the scenarios
defined on your worksheet



Return to [README.md](README.md)

