SCENARIO.EDIT

Equivalent to clicking the Scenarios command from the Tools menu and
then clicking the Edit button.

**Syntax**

**SCENARIO.EDIT**(**scen\_name**, new\_scenname, value\_array,
changing\_ref, scen\_comment, locked, hidden)

**SCENARIO.EDIT**?(scen\_name, new\_scenname, value\_array,
changing\_ref, scen\_comment, locked, hidden)

Scen\_name    is the name of the scenario that you want to edit.

New\_scenname    is the new name you want to give to the scenario.

Value\_array    is a horizontal array of values that you want to use for
the scenario.

  - > If value\_array is omitted but changing\_ref is specified,
    > Scenario Manager uses the values in changing\_ref as value\_array.

  - > Value\_array must match the dimensions of changing\_ref for the
    > scenario being edit.

Changing\_ref    is a reference to cells you want to define as changing
cells for a scenario.

Scen\_comment    is text specifying a descriptive comment for the
scenario you want to edit.

Locked    is a logical value that corresponds to the Prevent Changes
check box in the Add or Edit Scenario dialogs boxes. If TRUE or omitted
, prevents users from changing values in a scenario. If FALSE, users are
allowed to make changes to the scenario. The locking will not become
enabled until the sheet is protected with the Protect Sheet command from
the Protection submenu on the Tools menu.

Hidden    is a logical value that corresponds to the Hide check box in
the Add or Edit Scenario dialog boxes. If TRUE, the scenario will be
hidden from view from the users. If FALSE or omitted, the scenario will
remain unhidden. The scenario will not become hidden until the sheet is
hidden with the Hide command from the Window menu.

**Related Functions**

[SCENARIO.GET](SCENARIO.GET.md)   Returns the specified information about the scenarios
defined on your worksheet

[SCENARIO.ADD](SCENARIO.ADD.md)   Equivalent to clicking the Scenario Manager command on
the Tools menu and then clicking the Add button

[SCENARIO.DELETE](SCENARIO.DELETE.md)   Equivalent to clicking the Scenario Manager command on
the Tools menu and then selecting a scenario and clicking the Delete
button



Return to [README.md](README.md)

SCENARIO.EDIT

Equivalent to clicking the Scenarios command from the Tools menu and
then clicking the Edit button.

**Syntax**

**SCENARIO.EDIT**(**scen\_name**, new\_scenname, value\_array,
changing\_ref, scen\_comment, locked, hidden)

**SCENARIO.EDIT**?(scen\_name, new\_scenname, value\_array,
changing\_ref, scen\_comment, locked, hidden)

Scen\_name    is the name of the scenario that you want to edit.

New\_scenname    is the new name you want to give to the scenario.

Value\_array    is a horizontal array of values that you want to use for
the scenario.

  - > If value\_array is omitted but changing\_ref is specified,
    > Scenario Manager uses the values in changing\_ref as value\_array.

  - > Value\_array must match the dimensions of changing\_ref for the
    > scenario being edit.

Changing\_ref    is a reference to cells you want to define as changing
cells for a scenario.

Scen\_comment    is text specifying a descriptive comment for the
scenario you want to edit.

Locked    is a logical value that corresponds to the Prevent Changes
check box in the Add or Edit Scenario dialogs boxes. If TRUE or omitted
, prevents users from changing values in a scenario. If FALSE, users are
allowed to make changes to the scenario. The locking will not become
enabled until the sheet is protected with the Protect Sheet command from
the Protection submenu on the Tools menu.

Hidden    is a logical value that corresponds to the Hide check box in
the Add or Edit Scenario dialog boxes. If TRUE, the scenario will be
hidden from view from the users. If FALSE or omitted, the scenario will
remain unhidden. The scenario will not become hidden until the sheet is
hidden with the Hide command from the Window menu.

**Related Functions**

[SCENARIO.GET](SCENARIO.GET.md)   Returns the specified information about the scenarios
defined on your worksheet

[SCENARIO.ADD](SCENARIO.ADD.md)   Equivalent to clicking the Scenario Manager command on
the Tools menu and then clicking the Add button

[SCENARIO.DELETE](SCENARIO.DELETE.md)   Equivalent to clicking the Scenario Manager command on
the Tools menu and then selecting a scenario and clicking the Delete
button



Return to [README.md](README.md)

