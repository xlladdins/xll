SOLVER.GET

Returns information about current settings for Solver. The settings are
specified in the Solver Parameters and Solver Options dialog boxes.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.GET**(**type\_num**, sheet\_name)

Type\_num    is a number specifying the type of information you want.

The following settings are specified in the Solver Parameters dialog
box.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the Equal To option<br />
1 = Max<br />
2 = Min<br />
3 = Value of</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>The value in the Value Of box</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The reference (as a multiple reference if necessary) in the By Changing Cells box</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>The number of constraints</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>An array of the left sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>An array of numbers corresponding to the relationships between the left and right sides of the constraints:<br />
1 = &lt;=<br />
2 = =<br />
3 = &gt;=<br />
4 = int</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>An array of the right sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
</tbody>
</table>

The following settings are specified in the Solver Options dialog box:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>The maximum number of iterations</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>The precision</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>The integer tolerance value</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Assume Linear Model check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Show Iteration Results check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>15</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Use Automatic Scaling check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of estimates:<br />
1 = Tangent<br />
2 = Quadratic</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>17</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of derivatives:<br />
1 = Forward<br />
2 = Central</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>18</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of search:<br />
1 = Quasi-Newton<br />
2 = Conjugate Gradient</p>
</blockquote></td>
</tr>
</tbody>
</table>

Sheet\_name    is the name of a sheet that contains the scenario for
which you want information. If sheet\_name is omitted, it is assumed to
be the active sheet.


SOLVER.GET

Returns information about current settings for Solver. The settings are
specified in the Solver Parameters and Solver Options dialog boxes.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.GET**(**type\_num**, sheet\_name)

Type\_num    is a number specifying the type of information you want.

The following settings are specified in the Solver Parameters dialog
box.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the Equal To option<br />
1 = Max<br />
2 = Min<br />
3 = Value of</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>The value in the Value Of box</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The reference (as a multiple reference if necessary) in the By Changing Cells box</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>The number of constraints</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>An array of the left sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>An array of numbers corresponding to the relationships between the left and right sides of the constraints:<br />
1 = &lt;=<br />
2 = =<br />
3 = &gt;=<br />
4 = int</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>An array of the right sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
</tbody>
</table>

The following settings are specified in the Solver Options dialog box:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>The maximum number of iterations</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>The precision</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>The integer tolerance value</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Assume Linear Model check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Show Iteration Results check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>15</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Use Automatic Scaling check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of estimates:<br />
1 = Tangent<br />
2 = Quadratic</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>17</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of derivatives:<br />
1 = Forward<br />
2 = Central</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>18</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of search:<br />
1 = Quasi-Newton<br />
2 = Conjugate Gradient</p>
</blockquote></td>
</tr>
</tbody>
</table>

Sheet\_name    is the name of a sheet that contains the scenario for
which you want information. If sheet\_name is omitted, it is assumed to
be the active sheet.


SOLVER.GET

Returns information about current settings for Solver. The settings are
specified in the Solver Parameters and Solver Options dialog boxes.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.GET**(**type\_num**, sheet\_name)

Type\_num    is a number specifying the type of information you want.

The following settings are specified in the Solver Parameters dialog
box.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the Equal To option<br />
1 = Max<br />
2 = Min<br />
3 = Value of</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>The value in the Value Of box</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The reference (as a multiple reference if necessary) in the By Changing Cells box</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>The number of constraints</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>An array of the left sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>An array of numbers corresponding to the relationships between the left and right sides of the constraints:<br />
1 = &lt;=<br />
2 = =<br />
3 = &gt;=<br />
4 = int</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>An array of the right sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
</tbody>
</table>

The following settings are specified in the Solver Options dialog box:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>The maximum number of iterations</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>The precision</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>The integer tolerance value</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Assume Linear Model check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Show Iteration Results check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>15</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Use Automatic Scaling check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of estimates:<br />
1 = Tangent<br />
2 = Quadratic</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>17</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of derivatives:<br />
1 = Forward<br />
2 = Central</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>18</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of search:<br />
1 = Quasi-Newton<br />
2 = Conjugate Gradient</p>
</blockquote></td>
</tr>
</tbody>
</table>

Sheet\_name    is the name of a sheet that contains the scenario for
which you want information. If sheet\_name is omitted, it is assumed to
be the active sheet.


SOLVER.GET

Returns information about current settings for Solver. The settings are
specified in the Solver Parameters and Solver Options dialog boxes.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.GET**(**type\_num**, sheet\_name)

Type\_num    is a number specifying the type of information you want.

The following settings are specified in the Solver Parameters dialog
box.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the Equal To option<br />
1 = Max<br />
2 = Min<br />
3 = Value of</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>The value in the Value Of box</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The reference (as a multiple reference if necessary) in the By Changing Cells box</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>The number of constraints</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>An array of the left sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>An array of numbers corresponding to the relationships between the left and right sides of the constraints:<br />
1 = &lt;=<br />
2 = =<br />
3 = &gt;=<br />
4 = int</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>An array of the right sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
</tbody>
</table>

The following settings are specified in the Solver Options dialog box:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>The maximum number of iterations</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>The precision</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>The integer tolerance value</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Assume Linear Model check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Show Iteration Results check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>15</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Use Automatic Scaling check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of estimates:<br />
1 = Tangent<br />
2 = Quadratic</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>17</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of derivatives:<br />
1 = Forward<br />
2 = Central</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>18</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of search:<br />
1 = Quasi-Newton<br />
2 = Conjugate Gradient</p>
</blockquote></td>
</tr>
</tbody>
</table>

Sheet\_name    is the name of a sheet that contains the scenario for
which you want information. If sheet\_name is omitted, it is assumed to
be the active sheet.


SOLVER.GET

Returns information about current settings for Solver. The settings are
specified in the Solver Parameters and Solver Options dialog boxes.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.GET**(**type\_num**, sheet\_name)

Type\_num    is a number specifying the type of information you want.

The following settings are specified in the Solver Parameters dialog
box.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the Equal To option<br />
1 = Max<br />
2 = Min<br />
3 = Value of</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>The value in the Value Of box</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The reference (as a multiple reference if necessary) in the By Changing Cells box</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>The number of constraints</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>An array of the left sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>An array of numbers corresponding to the relationships between the left and right sides of the constraints:<br />
1 = &lt;=<br />
2 = =<br />
3 = &gt;=<br />
4 = int</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>An array of the right sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
</tbody>
</table>

The following settings are specified in the Solver Options dialog box:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>The maximum number of iterations</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>The precision</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>The integer tolerance value</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Assume Linear Model check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Show Iteration Results check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>15</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Use Automatic Scaling check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of estimates:<br />
1 = Tangent<br />
2 = Quadratic</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>17</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of derivatives:<br />
1 = Forward<br />
2 = Central</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>18</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of search:<br />
1 = Quasi-Newton<br />
2 = Conjugate Gradient</p>
</blockquote></td>
</tr>
</tbody>
</table>

Sheet\_name    is the name of a sheet that contains the scenario for
which you want information. If sheet\_name is omitted, it is assumed to
be the active sheet.


SOLVER.GET

Returns information about current settings for Solver. The settings are
specified in the Solver Parameters and Solver Options dialog boxes.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.GET**(**type\_num**, sheet\_name)

Type\_num    is a number specifying the type of information you want.

The following settings are specified in the Solver Parameters dialog
box.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the Equal To option<br />
1 = Max<br />
2 = Min<br />
3 = Value of</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>The value in the Value Of box</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The reference (as a multiple reference if necessary) in the By Changing Cells box</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>The number of constraints</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>An array of the left sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>An array of numbers corresponding to the relationships between the left and right sides of the constraints:<br />
1 = &lt;=<br />
2 = =<br />
3 = &gt;=<br />
4 = int</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>An array of the right sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
</tbody>
</table>

The following settings are specified in the Solver Options dialog box:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>The maximum number of iterations</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>The precision</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>The integer tolerance value</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Assume Linear Model check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Show Iteration Results check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>15</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Use Automatic Scaling check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of estimates:<br />
1 = Tangent<br />
2 = Quadratic</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>17</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of derivatives:<br />
1 = Forward<br />
2 = Central</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>18</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of search:<br />
1 = Quasi-Newton<br />
2 = Conjugate Gradient</p>
</blockquote></td>
</tr>
</tbody>
</table>

Sheet\_name    is the name of a sheet that contains the scenario for
which you want information. If sheet\_name is omitted, it is assumed to
be the active sheet.


SOLVER.GET

Returns information about current settings for Solver. The settings are
specified in the Solver Parameters and Solver Options dialog boxes.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.GET**(**type\_num**, sheet\_name)

Type\_num    is a number specifying the type of information you want.

The following settings are specified in the Solver Parameters dialog
box.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the Equal To option<br />
1 = Max<br />
2 = Min<br />
3 = Value of</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>The value in the Value Of box</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The reference (as a multiple reference if necessary) in the By Changing Cells box</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>The number of constraints</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>An array of the left sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>An array of numbers corresponding to the relationships between the left and right sides of the constraints:<br />
1 = &lt;=<br />
2 = =<br />
3 = &gt;=<br />
4 = int</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>An array of the right sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
</tbody>
</table>

The following settings are specified in the Solver Options dialog box:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>The maximum number of iterations</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>The precision</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>The integer tolerance value</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Assume Linear Model check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Show Iteration Results check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>15</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Use Automatic Scaling check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of estimates:<br />
1 = Tangent<br />
2 = Quadratic</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>17</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of derivatives:<br />
1 = Forward<br />
2 = Central</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>18</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of search:<br />
1 = Quasi-Newton<br />
2 = Conjugate Gradient</p>
</blockquote></td>
</tr>
</tbody>
</table>

Sheet\_name    is the name of a sheet that contains the scenario for
which you want information. If sheet\_name is omitted, it is assumed to
be the active sheet.


SOLVER.GET

Returns information about current settings for Solver. The settings are
specified in the Solver Parameters and Solver Options dialog boxes.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.GET**(**type\_num**, sheet\_name)

Type\_num    is a number specifying the type of information you want.

The following settings are specified in the Solver Parameters dialog
box.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the Equal To option<br />
1 = Max<br />
2 = Min<br />
3 = Value of</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>The value in the Value Of box</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The reference (as a multiple reference if necessary) in the By Changing Cells box</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>The number of constraints</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>An array of the left sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>An array of numbers corresponding to the relationships between the left and right sides of the constraints:<br />
1 = &lt;=<br />
2 = =<br />
3 = &gt;=<br />
4 = int</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>An array of the right sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
</tbody>
</table>

The following settings are specified in the Solver Options dialog box:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>The maximum number of iterations</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>The precision</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>The integer tolerance value</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Assume Linear Model check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Show Iteration Results check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>15</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Use Automatic Scaling check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of estimates:<br />
1 = Tangent<br />
2 = Quadratic</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>17</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of derivatives:<br />
1 = Forward<br />
2 = Central</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>18</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of search:<br />
1 = Quasi-Newton<br />
2 = Conjugate Gradient</p>
</blockquote></td>
</tr>
</tbody>
</table>

Sheet\_name    is the name of a sheet that contains the scenario for
which you want information. If sheet\_name is omitted, it is assumed to
be the active sheet.


SOLVER.GET

Returns information about current settings for Solver. The settings are
specified in the Solver Parameters and Solver Options dialog boxes.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.GET**(**type\_num**, sheet\_name)

Type\_num    is a number specifying the type of information you want.

The following settings are specified in the Solver Parameters dialog
box.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the Equal To option<br />
1 = Max<br />
2 = Min<br />
3 = Value of</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>The value in the Value Of box</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The reference (as a multiple reference if necessary) in the By Changing Cells box</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>The number of constraints</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>An array of the left sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>An array of numbers corresponding to the relationships between the left and right sides of the constraints:<br />
1 = &lt;=<br />
2 = =<br />
3 = &gt;=<br />
4 = int</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>An array of the right sides of the constraints in the form of text</p>
</blockquote></td>
</tr>
</tbody>
</table>

The following settings are specified in the Solver Options dialog box:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_Num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Returns</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>The maximum number of iterations</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>The precision</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>The integer tolerance value</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Assume Linear Model check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>14</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Show Iteration Results check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>15</p>
</blockquote></td>
<td><blockquote>
<p>TRUE if the Use Automatic Scaling check box is selected; FALSE otherwise</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of estimates:<br />
1 = Tangent<br />
2 = Quadratic</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>17</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of derivatives:<br />
1 = Forward<br />
2 = Central</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>18</p>
</blockquote></td>
<td><blockquote>
<p>A number corresponding to the type of search:<br />
1 = Quasi-Newton<br />
2 = Conjugate Gradient</p>
</blockquote></td>
</tr>
</tbody>
</table>

Sheet\_name    is the name of a sheet that contains the scenario for
which you want information. If sheet\_name is omitted, it is assumed to
be the active sheet.


