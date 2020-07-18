# ON.TIME

Runs a macro at a specified time. Use ON.TIME to run a macro at a
specific time of day or after a specified period has passed.

**Syntax**

**ON.TIME**(**time, macro\_text**, tolerance, insert\_logical)

Time&nbsp;&nbsp;&nbsp;&nbsp;is the time and date, given as a serial
number, at which the macro is to be run. If time does not include a date
(that is, if time is a serial number less than 1), the macro is run the
next time this time occurs.

Macro\_text&nbsp;&nbsp;&nbsp;&nbsp;is the name of, or an R1C1-style
reference to, a macro to run at the specified time and every subsequent
day at that time.

Tolerance&nbsp;&nbsp;&nbsp;&nbsp;is the time and date, given as a serial
number, that you are willing to wait until and still have the macro run.
For example, if Microsoft Excel is not in Ready, Copy, Cut, or Find mode
at time, because another macro is running, but is in Ready mode 15
seconds later, and tolerance is set to time plus 30 seconds, the macro
specified by macro\_text will run. If Microsoft Excel was not in Ready
mode within 30 seconds, the macro would not run. If tolerance is
omitted, it is assumed to be infinite.

Insert\_logical&nbsp;&nbsp;&nbsp;&nbsp;is a logical value specifying
whether you want every day macro\_text to run at time. Use
insert\_logical when you want to clear a previously set ON.TIME formula.
If insert\_logical is TRUE or omitted, the macro specified by
macro\_text is carried out at time. If insert\_logical is FALSE and
macro\_text is not set to run at time, ON.TIME returns the \#VALUE error
value.

**Examples**

The following macro formula runs a macro called Test at 5:00:00 P.M.
every day when Microsoft Excel is in Ready mode:

ON.TIME("5:00:00 PM", "Test")

The following macro formula runs a macro called Test 5 seconds after the
formula is evaluated:

ON.TIME(NOW()+"00:00:05", "Test")

The following macro formula runs a macro called Test 10 seconds after
the formula is evaluated. If Microsoft Excel is not in Ready mode at
that time (because it is in Edit mode, for example), the tolerance
argument specifies 5 seconds of additional time to wait to run the
macro. If Microsoft Excel is still not in Ready mode at that time,
macro\_text is not run.

ON.TIME(NOW()+"00:00:10", "Test", NOW()+"00:00:15")



Return to [README](README.md)

