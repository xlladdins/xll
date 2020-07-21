# SOLVER.OPTIONS

Equivalent to clicking the Solver command on the Tools menu and then
clicking the Options button in the Solver Parameters dialog box.
Specifies the available options.

If this function is not available, you must install the Solver add-in.

**Syntax**

**SOLVER.OPTIONS**(max\_time, iterations, precision, assume\_linear,
step\_thru, estimates, derivatives, search, int\_tolerance, scaling)

The arguments correspond to the options in the dialog box. If an
argument is omitted, Microsoft Excel uses an appropriate value based on
the current situation. If any of the arguments are the wrong type, the
function returns the \#N/A error value. If an argument has the correct
type but an invalid value, the function returns a positive integer
corresponding to its position. A zero indicates all options were
accepted.

Max\_time&nbsp;&nbsp;&nbsp;&nbsp;must be an integer greater than zero
and less than 32768. It corresponds to the Max Time box.

Iterations&nbsp;&nbsp;&nbsp;&nbsp;must be an integer greater than zero
and less than 32768. It corresponds to the Iterations box.

Precision&nbsp;&nbsp;&nbsp;&nbsp;must be a number between zero and one,
but not equal to zero or one. It corresponds to the Precision box.

Assume\_linear&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding
to the Assume Linear Model check box and allows Solver to arrive at a
solution more quickly. If TRUE, Solver assumes that the underlying model
is linear; if FALSE, it does not.

Step\_thru&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to
the Show Iteration Results check box. If you have supplied SOLVER.SOLVE
with a valid command macro reference, your macro will be called each
time Solver pauses. If TRUE, Solver pauses at each trial solution; if
FALSE, it does not.

Estimates&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and corresponds to
the Estimates options: 1 for the Tangent option and 2 for the Quadratic
option.

Derivatives&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and corresponds
to the Derivatives options: 1 for the Forward option and 2 for the
Central option.

Search&nbsp;&nbsp;&nbsp;&nbsp;is the number 1 or 2 and corresponds to
the Search options: 1 for the Quasi-Newton option and 2 for the
Conjugate Gradient option.

Int\_tolerance&nbsp;&nbsp;&nbsp;&nbsp;is a decimal number corresponding
to the Tolerance box in the Solver Options dialog box, and must be
between zero and 1, inclusively. This argument applies only if integer
constraints have been defined.

Scaling&nbsp;&nbsp;&nbsp;&nbsp;is a logical value corresponding to the
Use Automatic Scaling check box. If scaling is TRUE, then if two or more
constraints differ by several orders of magnitude, Solver scales the
constraints to similar orders of magnitude during computation. If
scaling is FALSE, Solver calculates normally.



Return to [README](README.md#S)

