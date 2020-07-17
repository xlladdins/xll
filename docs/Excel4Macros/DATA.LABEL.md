DATA.LABEL
==========

Specifies label contents and position.

**Syntax**

**DATA.LABEL**(show\_option, auto\_text, show\_key)

Show\_option    is a number that specifies what type of labels to
display.

  ------------------ ------------------------
  **Show\_option**   **Type displayed**
  1                  none
  2                  Show value
  3                  Show percent
  4                  Show label
  5                  Show label and percent
  ------------------ ------------------------

Auto\_text    is a logical value that corresponds the Automatic Checkbox
in the Data Labels dialog box. If TRUE, resets a chart\'s data labels
back to their actual values. If FALSE, they are not reset. The Automatic
Text checkbox appears only if the label has been selected and its value
changed.

Show\_key    is a logical value that specified whether to show the
legend key next to the label. If TRUE, displays the legend key. If FALSE
or omitted, does not display the legend key.

Return to [top](#A)

DATA.SERIES
