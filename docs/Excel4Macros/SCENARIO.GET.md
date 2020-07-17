SCENARIO.GET
============

Returns the specified information about the scenarios defined on your
worksheet.

**Syntax**

**SCENARIO.GET**(**type\_num**, scen\_name)

Type\_num    is a number from 1 to 8 specifying the type of information
you want.

+-----------------+---------------------------------------------------+
| > **Type\_num** | > **Information returned**                        |
+-----------------+---------------------------------------------------+
| > 1             | > A horizontal array of all scenario names in the |
|                 | > form of text                                    |
+-----------------+---------------------------------------------------+
| > 2             | > A reference to the set of changing cells of     |
|                 | > scen\_name (specified in the Changing Cells box |
|                 | > of the Scenario Manager dialog box). If         |
|                 | > scen\_name is omitted, the first scenario is    |
|                 | > used.                                           |
+-----------------+---------------------------------------------------+
| > 3             | > A reference to the result cells (specified in   |
|                 | > the Result Cells box in the Scenario Summary    |
|                 | > dialog box)                                     |
+-----------------+---------------------------------------------------+
| > 4             | > An array of scenario values for the scenario    |
|                 | > scen\_name . Each scenario is in a separate     |
|                 | > row. If scen\_name is omitted, the first        |
|                 | > scenario is used.                               |
+-----------------+---------------------------------------------------+
| > 5             | > Comment, as text, for the scenario              |
+-----------------+---------------------------------------------------+
| > 6             | > Returns TRUE if the specified scenario is       |
|                 | > locked to prevent changes; FALSE, if unlocked.  |
|                 | > Scen\_name is required.                         |
+-----------------+---------------------------------------------------+
| > 7             | > Returns TRUE if the specified scenario is       |
|                 | > hidden; FALSE, if visible to the user.          |
|                 | > Scen\_name is required.                         |
+-----------------+---------------------------------------------------+
| > 8             | > Returns the user name of the person who last    |
|                 | > modified the scenario by either adding or       |
|                 | > editing a scenario. Scen\_name is required.     |
+-----------------+---------------------------------------------------+

Scen\_name    is the name of the scenario that you want information
about. Ignored if type\_num equals 1 or 3.

**Remarks**

In the returned array of scenario values, the number of rows is the
number of scenarios, and the number of columns is the number of changing
cells.

Return to [top](#Q)

SCENARIO.MERGE
