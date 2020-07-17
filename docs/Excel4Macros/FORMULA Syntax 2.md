FORMULA Syntax 2
================

Enters a text label or SERIES formula in a chart. To enter formulas on a
worksheet or macro sheet, use syntax 1 of this function.

**Syntax**

**FORMULA**(**formula\_text**)

Formula\_text    is the text label or SERIES formula you want to enter
into the chart.

  ----------------------------------------------------------------------------------------------------------------------------- -------------------------------------------------------------
  **If**                                                                                                                        **Then**
  Formula\_text can be treated as a text label and the current selection is a text label                                        The selected text label is replaced with formula\_text.
  Formula\_text can be treated as a text label and there is no current selection or the current selection is not a text label   Formula\_text creates a new unattached text label.
  Formula\_text can be treated as a SERIES formula and the current selection is a SERIES formula                                The selected SERIES formula is replaced with formula\_text.
  Formula\_text can be treated as a SERIES formula and the current selection is not a SERIES formula                            Formula\_text creates a new SERIES formula.
  ----------------------------------------------------------------------------------------------------------------------------- -------------------------------------------------------------

**Remarks**

You would normally use the EDIT.SERIES function to create or edit a
chart series. For more information, see EDIT.SERIES.

**Example**

The following macro formula enters a SERIES formula on the chart. If the
current selection is a SERIES formula, it is replaced:

FORMULA(\"=SERIES(\"\"Title\"\", , {1, 2, 3}, 1)\")

**Related Functions**

EDIT.SERIES   Creates or changes a chart series

FORMULA, Syntax 1   Enters numbers, text, references, and formulas in a
worksheet

Return to [top](#E)

FORMULA.ARRAY
