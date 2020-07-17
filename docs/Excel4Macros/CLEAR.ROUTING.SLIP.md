CLEAR.ROUTING.SLIP

Equivalent to the Clear button in the Routing Slip dialog box. Clears
the routing slip.

**Syntax**

**CLEAR.ROUTING.SLIP**(reset\_only\_logical)

Reset\_only\_logical    is a logical value that specifies whether the
routing slip should be cleared.

  - > This option is valid only after every recipient on the routing
    > slip has received and forwarded the workbook. Setting
    > reset\_only\_logical to TRUE in this case is equivalent to the
    > Reset button in the routing slip dialog.

  - > If some recipients have not received or routed the workbook,
    > reset\_only\_logical is ignored.

  - > If reset\_only\_logical is FALSE or omitted and the workbook has
    > been routed to all recipients, then the routing slip is removed
    > from the workbook. A new slip can be subsequently added using
    > ROUTING.SLIP.



Return to [README](README.md)

