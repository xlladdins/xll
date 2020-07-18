RESULT

Specifies the type of data a macro or custom function returns. Use
RESULT to make sure your macros, custom functions, or subroutines return
values of the correct data type.

**Syntax**

**RESULT**(type\_num)

Type\_num**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;is a number specifying the data type.

<table>
<tbody>
<tr class="odd">
<td><blockquote>
<p><strong>Type_num</strong></p>
</blockquote></td>
<td><blockquote>
<p><strong>Type of returned data</strong></p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>1</p>
</blockquote></td>
<td><blockquote>
<p>Number</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>2</p>
</blockquote></td>
<td><blockquote>
<p>Text</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>4</p>
</blockquote></td>
<td><blockquote>
<p>Logical</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>8</p>
</blockquote></td>
<td><blockquote>
<p>Reference</p>
</blockquote></td>
</tr>
<tr class="even">
<td><blockquote>
<p>16</p>
</blockquote></td>
<td><blockquote>
<p>Error</p>
</blockquote></td>
</tr>
<tr class="odd">
<td><blockquote>
<p>64</p>
</blockquote></td>
<td><blockquote>
<p>Array</p>
</blockquote></td>
</tr>
</tbody>
</table>

&nbsp;

  - > Type\_num can be the sum of the numbers in the preceding table to
    > allow for more than one possible result type. For example, if
    > type\_num is 12, which equals 4 + 8, the result can be a logical
    > or a reference value.

  - > If you omit type\_num, it is assumed to be 7. Since 7 equals 1 + 2
    > + 4, the value returned can be a number (1), text (2), or logical
    > value (4).


**Examples**

The following function specifies that a custom function's return value
can be a number or a logical value (4+1=5):

RESULT(5)

**Related Functions**

[ARGUMENT](ARGUMENT.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Passes an argument to a macro

[RETURN](RETURN.md)**&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;&nbsp;&nbsp;&nbsp;nbsp;Ends the currently running macro



Return to [README](README.md)

