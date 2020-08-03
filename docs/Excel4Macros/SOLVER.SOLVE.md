# SOLVER.SOLVE

Equivalent to clicking the Solver command on the Tools menu and clicking
the Solve button in the Solver Parameters dialog box. If successful,
returns an integer value indicating the condition that caused Solver to
stop as described in "Remarks" later in this topic.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.SOLVE**(user\_finish, show\_ref)

User\_finish&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying
whether to display the Solver Results dialog box.

  - > If user\_finish is TRUE, SOLVER.SOLVE returns its integer value
    > without displaying anything. Your macro should decide what action
    > to take (for example, by examining the return value or presenting
    > its own dialog box); it must call SOLVER.FINISH in any case to
    > restore the sheet to its proper state.

  - > If user\_finish is FALSE or omitted, Solver displays the Solver
    > Results dialog box, which allows you to keep or discard the final
    > solution and run reports.


Show\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a macro to be called in place of the
Show Trial Solution dialog box. It is used when you want to regain
control whenever Solver finds a new intermediate solution value.

  - > For this argument to have an effect, the Show Iteration Results
    > check box must be selected in the Solver Options dialog box. This
    > can be done manually by selecting the check box, or automatically
    > by calling SOLVER.OPTIONS in your macro.

  - > The macro you call can inspect the current solution values on the
    > sheet or take other actions such as saving or charting the
    > intermediate values. It must return the value TRUE with a
    > statement such as =RETURN(TRUE) if the solution process is to
    > continue, or FALSE if the solution process should stop at this
    > point.


**Remarks**

If a problem has not been completely defined, SOLVER.SOLVE returns the
\#N/A error value. Otherwise, the Solver application is started and the
problem specifications are passed to it. When the solution process is
complete, SOLVER.SOLVE returns an integer value indicating the stopping
condition:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Value</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Stopping condition</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>0</p>
</blockquote></td>
<td><blockquote>
<p>Solver found a solution. All constraints and optimality conditions are satisfied.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Solver has converged to the current solution. All constraints are satisfied.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Solver cannot improve the current solution. All constraints are satisfied.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Stop chosen when the maximum iteration limit was reached.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The Set Cells values do not converge.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Solver could not find a feasible solution.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Solver stopped at user's request.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>The conditions for Assume Linear Model are not satisfied.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>The problem is too large for Solver to solve.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>Solver encountered an error value in a target or constraint cell.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>Stop chosen when the maximum time limit was reached.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>There is not enough memory available to solve the problem.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>Another Excel instance is using SOLVER.DLL. Try again later.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>Error in model. Please verify that all cells and constraints are valid.</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Related Function**

[SOLVER.FINISH](SOLVER.FINISH.md)&nbsp;&nbsp;&nbsp;Equivalent to clicking OK in the Solver
[R](R.md)esults dialog box that appears when the solution process is complete



Return to [README](README.md#S)

# SOLVER.SOLVE

Equivalent to clicking the Solver command on the Tools menu and clicking
the Solve button in the Solver Parameters dialog box. If successful,
returns an integer value indicating the condition that caused Solver to
stop as described in "Remarks" later in this topic.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.SOLVE**(user\_finish, show\_ref)

User\_finish&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying
whether to display the Solver Results dialog box.

  - > If user\_finish is TRUE, SOLVER.SOLVE returns its integer value
    > without displaying anything. Your macro should decide what action
    > to take (for example, by examining the return value or presenting
    > its own dialog box); it must call SOLVER.FINISH in any case to
    > restore the sheet to its proper state.

  - > If user\_finish is FALSE or omitted, Solver displays the Solver
    > Results dialog box, which allows you to keep or discard the final
    > solution and run reports.


Show\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a macro to be called in place of the
Show Trial Solution dialog box. It is used when you want to regain
control whenever Solver finds a new intermediate solution value.

  - > For this argument to have an effect, the Show Iteration Results
    > check box must be selected in the Solver Options dialog box. This
    > can be done manually by selecting the check box, or automatically
    > by calling SOLVER.OPTIONS in your macro.

  - > The macro you call can inspect the current solution values on the
    > sheet or take other actions such as saving or charting the
    > intermediate values. It must return the value TRUE with a
    > statement such as =RETURN(TRUE) if the solution process is to
    > continue, or FALSE if the solution process should stop at this
    > point.


**Remarks**

If a problem has not been completely defined, SOLVER.SOLVE returns the
\#N/A error value. Otherwise, the Solver application is started and the
problem specifications are passed to it. When the solution process is
complete, SOLVER.SOLVE returns an integer value indicating the stopping
condition:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Value</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Stopping condition</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>0</p>
</blockquote></td>
<td><blockquote>
<p>Solver found a solution. All constraints and optimality conditions are satisfied.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Solver has converged to the current solution. All constraints are satisfied.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Solver cannot improve the current solution. All constraints are satisfied.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Stop chosen when the maximum iteration limit was reached.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The Set Cells values do not converge.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Solver could not find a feasible solution.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Solver stopped at user's request.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>The conditions for Assume Linear Model are not satisfied.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>The problem is too large for Solver to solve.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>Solver encountered an error value in a target or constraint cell.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>Stop chosen when the maximum time limit was reached.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>There is not enough memory available to solve the problem.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>Another Excel instance is using SOLVER.DLL. Try again later.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>Error in model. Please verify that all cells and constraints are valid.</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Related Function**

[SOLVER.FINISH](SOLVER.FINISH.md)&nbsp;&nbsp;&nbsp;Equivalent to clicking OK in the Solver
[R](R.md)esults dialog box that appears when the solution process is complete



Return to [README](README.md#S)

# SOLVER.SOLVE

Equivalent to clicking the Solver command on the Tools menu and clicking
the Solve button in the Solver Parameters dialog box. If successful,
returns an integer value indicating the condition that caused Solver to
stop as described in "Remarks" later in this topic.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.SOLVE**(user\_finish, show\_ref)

User\_finish&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying
whether to display the Solver Results dialog box.

  - > If user\_finish is TRUE, SOLVER.SOLVE returns its integer value
    > without displaying anything. Your macro should decide what action
    > to take (for example, by examining the return value or presenting
    > its own dialog box); it must call SOLVER.FINISH in any case to
    > restore the sheet to its proper state.

  - > If user\_finish is FALSE or omitted, Solver displays the Solver
    > Results dialog box, which allows you to keep or discard the final
    > solution and run reports.


Show\_ref&nbsp;&nbsp;&nbsp;&nbsp;is a macro to be called in place of the
Show Trial Solution dialog box. It is used when you want to regain
control whenever Solver finds a new intermediate solution value.

  - > For this argument to have an effect, the Show Iteration Results
    > check box must be selected in the Solver Options dialog box. This
    > can be done manually by selecting the check box, or automatically
    > by calling SOLVER.OPTIONS in your macro.

  - > The macro you call can inspect the current solution values on the
    > sheet or take other actions such as saving or charting the
    > intermediate values. It must return the value TRUE with a
    > statement such as =RETURN(TRUE) if the solution process is to
    > continue, or FALSE if the solution process should stop at this
    > point.


**Remarks**

If a problem has not been completely defined, SOLVER.SOLVE returns the
\#N/A error value. Otherwise, the Solver application is started and the
problem specifications are passed to it. When the solution process is
complete, SOLVER.SOLVE returns an integer value indicating the stopping
condition:

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Value</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Stopping condition</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>0</p>
</blockquote></td>
<td><blockquote>
<p>Solver found a solution. All constraints and optimality conditions are satisfied.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Solver has converged to the current solution. All constraints are satisfied.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Solver cannot improve the current solution. All constraints are satisfied.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>3</p>
</blockquote></td>
<td><blockquote>
<p>Stop chosen when the maximum iteration limit was reached.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>The Set Cells values do not converge.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>5</p>
</blockquote></td>
<td><blockquote>
<p>Solver could not find a feasible solution.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>6</p>
</blockquote></td>
<td><blockquote>
<p>Solver stopped at user's request.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>7</p>
</blockquote></td>
<td><blockquote>
<p>The conditions for Assume Linear Model are not satisfied.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>The problem is too large for Solver to solve.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>9</p>
</blockquote></td>
<td><blockquote>
<p>Solver encountered an error value in a target or constraint cell.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>10</p>
</blockquote></td>
<td><blockquote>
<p>Stop chosen when the maximum time limit was reached.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>11</p>
</blockquote></td>
<td><blockquote>
<p>There is not enough memory available to solve the problem.</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>12</p>
</blockquote></td>
<td><blockquote>
<p>Another Excel instance is using SOLVER.DLL. Try again later.</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>13</p>
</blockquote></td>
<td><blockquote>
<p>Error in model. Please verify that all cells and constraints are valid.</p>
</blockquote></td>
</tr>
</tbody>
</table>

**Related Function**

[SOLVER.FINISH](SOLVER.FINISH.md)&nbsp;&nbsp;&nbsp;Equivalent to clicking OK in the Solver
[R](R.md)esults dialog box that appears when the solution process is complete



Return to [README](README.md#S)

