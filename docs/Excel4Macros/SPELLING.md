SPELLING
========

Equivalent to clicking the Spelling command on the Tools menu. Checks
the spelling of words in the current selection.

**Syntax**

**SPELLING**(custom\_dic, ignore\_uppercase, always\_suggest)

Custom\_dic    is the filename of the custom dictionary to examine if
words are not found in the main dictionary. If custom\_dic is omitted,
the currently specified dictionary is used.

Ignore\_uppercase    is a logical value corresponding to the Ignore
UPPERCASE check box.

+-------------------------------+-----------------------------------------+
| > **If ignore\_uppercase is** | > **Microsoft Excel will**              |
+-------------------------------+-----------------------------------------+
| > TRUE                        | > Ignore words in all uppercase letters |
+-------------------------------+-----------------------------------------+
| > FALSE                       | > Check words in all uppercase letters  |
+-------------------------------+-----------------------------------------+
| > Omitted                     | > Use the current setting               |
+-------------------------------+-----------------------------------------+

Always\_suggest    is a logical value corresponding to the Always
Suggest check box.

+-----------------------------+---------------------------------------+
| > **If always\_suggest is** | > **Microsoft Excel will**            |
+-----------------------------+---------------------------------------+
| > TRUE                      | > Display a list of suggested         |
|                             | > alternate spellings when an         |
|                             | > incorrect spelling is found         |
+-----------------------------+---------------------------------------+
| > FALSE                     | > Wait for user to input the correct  |
|                             | > spelling                            |
+-----------------------------+---------------------------------------+
| > Omitted                   | > Use the current setting             |
+-----------------------------+---------------------------------------+

**Related Function**

SPELLING.CHECK   Checks the spelling of a word

Return to [top](#Q)

SPELLING.CHECK
