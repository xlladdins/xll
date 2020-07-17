SOLVER.GET
==========

Returns information about current settings for Solver. The settings are
specified in the Solver Parameters and Solver Options dialog boxes.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.GET**(**type\_num**, sheet\_name)

Type\_num    is a number specifying the type of information you want.

The following settings are specified in the Solver Parameters dialog
box.

+-----------------+---------------------------------------------------+
| > **Type\_Num** | > **Returns**                                     |
+-----------------+---------------------------------------------------+
| > 2             | > A number corresponding to the Equal To option\  |
|                 | > 1 = Max\                                        |
|                 | > 2 = Min\                                        |
|                 | > 3 = Value of                                    |
+-----------------+---------------------------------------------------+
| > 3             | > The value in the Value Of box                   |
+-----------------+---------------------------------------------------+
| > 4             | > The reference (as a multiple reference if       |
|                 | > necessary) in the By Changing Cells box         |
+-----------------+---------------------------------------------------+
| > 5             | > The number of constraints                       |
+-----------------+---------------------------------------------------+
| > 6             | > An array of the left sides of the constraints   |
|                 | > in the form of text                             |
+-----------------+---------------------------------------------------+
| > 7             | > An array of numbers corresponding to the        |
|                 | > relationships between the left and right sides  |
|                 | > of the constraints:\                            |
|                 | > 1 = \<=\                                        |
|                 | > 2 = =\                                          |
|                 | > 3 = \>=\                                        |
|                 | > 4 = int                                         |
+-----------------+---------------------------------------------------+
| > 8             | > An array of the right sides of the constraints  |
|                 | > in the form of text                             |
+-----------------+---------------------------------------------------+

The following settings are specified in the Solver Options dialog box:

+-----------------+---------------------------------------------------+
| > **Type\_Num** | > **Returns**                                     |
+-----------------+---------------------------------------------------+
| > 10            | > The maximum number of iterations                |
+-----------------+---------------------------------------------------+
| > 11            | > The precision                                   |
+-----------------+---------------------------------------------------+
| > 12            | > The integer tolerance value                     |
+-----------------+---------------------------------------------------+
| > 13            | > TRUE if the Assume Linear Model check box is    |
|                 | > selected; FALSE otherwise                       |
+-----------------+---------------------------------------------------+
| > 14            | > TRUE if the Show Iteration Results check box is |
|                 | > selected; FALSE otherwise                       |
+-----------------+---------------------------------------------------+
| > 15            | > TRUE if the Use Automatic Scaling check box is  |
|                 | > selected; FALSE otherwise                       |
+-----------------+---------------------------------------------------+
| > 16            | > A number corresponding to the type of           |
|                 | > estimates:\                                     |
|                 | > 1 = Tangent\                                    |
|                 | > 2 = Quadratic                                   |
+-----------------+---------------------------------------------------+
| > 17            | > A number corresponding to the type of           |
|                 | > derivatives:\                                   |
|                 | > 1 = Forward\                                    |
|                 | > 2 = Central                                     |
+-----------------+---------------------------------------------------+
| > 18            | > A number corresponding to the type of search:\  |
|                 | > 1 = Quasi-Newton\                               |
|                 | > 2 = Conjugate Gradient                          |
+-----------------+---------------------------------------------------+

Sheet\_name    is the name of a sheet that contains the scenario for
which you want information. If sheet\_name is omitted, it is assumed to
be the active sheet.

Return to [top](#Q)

SOLVER.LOAD
