# TERMINATE

Closes a dynamic data exchange (DDE) channel previously opened with the
INITIATE function. Use TERMINATE to close a channel after you have
finished communicating with another application.

**Syntax**

**TERMINATE**(**channel\_num**)

**Important**&nbsp;&nbsp;&nbsp;Microsoft Excel for the Macintosh
requires system software version 7.0 or later for this function.

Channel\_num&nbsp;&nbsp;&nbsp;&nbsp;is the number returned by a
previously run INITIATE function. Channel\_num identifies a DDE channel
to close.

If TERMINATE is not successful, it returns the \#VALUE\! error value.

**Related Functions**

[EXEC](EXEC.md)&nbsp;&nbsp;&nbsp;Starts another application

[INITIATE](INITIATE.md)&nbsp;&nbsp;&nbsp;Opens a channel to another application



Return to [README](README.md)

