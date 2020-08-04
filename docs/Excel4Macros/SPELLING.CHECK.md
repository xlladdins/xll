# SPELLING.CHECK

Checks the spelling of a word. Returns TRUE if the word is spelled
correctly; FALSE otherwise.

**Syntax**

**SPELLING.CHECK**(**word\_text**, custom\_dic, ignore\_uppercase)

Word\_text&nbsp;&nbsp;&nbsp;&nbsp;is the word whose spelling you want to
check. It can be text or a reference to text.

Custom\_dic&nbsp;&nbsp;&nbsp;&nbsp;is the filename of a custom
dictionary to examine if the word is not found in the main dictionary.

Ignore\_uppercase&nbsp;&nbsp;&nbsp;&nbsp;is a logical value
corresponding to the Ignore Words In Uppercase check box. If
ignore\_uppercase is TRUE, the check box is selected, and Microsoft
Excel ignores words in all uppercase letters; if FALSE, the check box is
cleared, and Microsoft Excel checks all words; if omitted, the current
setting is used.

**Remarks**

This function does not have a dialog-box form. To display the Spelling
dialog box, use SPELLING.

**Related Function**

[SPELLING](SPELLING.md)&nbsp;&nbsp;&nbsp;Checks the spelling of words in the current
selection



Return to [README](README.md#S)

