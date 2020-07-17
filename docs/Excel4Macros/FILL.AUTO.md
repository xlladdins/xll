FILL.AUTO

Equivalent to copying cells or automatically filling a selection by
dragging the fill selection handle with the mouse (the AutoFill
feature).

**Syntax**

**FILL.AUTO**(destination\_ref, copy\_only)

Destination\_ref    is the range of cells into which you want to fill
data. The top, bottom, left, or right end of destination\_ref must
include all of the cells in the source reference (the current
selection).

Copy\_only    is a number specifying whether to copy cells or perform an
AutoFill operation.

|              |                      |
| ------------ | -------------------- |
| **Value**    | **Result**           |
| 0 or omitted | Normal AutoFill      |
| 1 or TRUE    | Copy cells           |
| 2            | Copy formats         |
| 3            | Fill values          |
| 4            | Increment            |
| 5            | Increment by day     |
| 6            | Increment by weekday |
| 7            | Increment by month   |
| 8            | Increment by year    |
| 9            | Linear trend         |
| 10           | Growth trend         |

**Related Functions**

[COPY](COPY.md)   Copies and pastes data or objects

[DATA.SERIES](DATA.SERIES.md)   Fills a range of cells with a series of numbers or dates


