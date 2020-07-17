GET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the Bcc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>The subject of the PowerTalk mailer, as text.</td>
</tr>
<tr class="even">
<td>29</td>
<td>The enclosures of the PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>TRUE, if the PowerTalk mailer has been received from another user (as opposed to just being added but not sent). FALSE, if the mailer has not been received from another user.</td>
</tr>
<tr class="even">
<td>31</td>
<td>The date and time the PowerTalk mailer was sent, as a serial number. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The sender name of the PowerTalk mailer, as text. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="even">
<td>33</td>
<td>The title of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>The subject of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>35</td>
<td>The author of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>The keywords for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>37</td>
<td>The comments for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>The name of the active sheet.</td>
</tr>
</tbody>
</table>

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

**Example**

The following macro formula returns the name of the active sheet in the
workbook named SALES.XLS:

GET.WORKBOOK(38, "SALES.XLS")

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

WORKBOOK.SELECT   Selects the specified documents in a workbook


GET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the Bcc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>The subject of the PowerTalk mailer, as text.</td>
</tr>
<tr class="even">
<td>29</td>
<td>The enclosures of the PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>TRUE, if the PowerTalk mailer has been received from another user (as opposed to just being added but not sent). FALSE, if the mailer has not been received from another user.</td>
</tr>
<tr class="even">
<td>31</td>
<td>The date and time the PowerTalk mailer was sent, as a serial number. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The sender name of the PowerTalk mailer, as text. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="even">
<td>33</td>
<td>The title of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>The subject of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>35</td>
<td>The author of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>The keywords for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>37</td>
<td>The comments for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>The name of the active sheet.</td>
</tr>
</tbody>
</table>

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

**Example**

The following macro formula returns the name of the active sheet in the
workbook named SALES.XLS:

GET.WORKBOOK(38, "SALES.XLS")

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

WORKBOOK.SELECT   Selects the specified documents in a workbook


GET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the Bcc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>The subject of the PowerTalk mailer, as text.</td>
</tr>
<tr class="even">
<td>29</td>
<td>The enclosures of the PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>TRUE, if the PowerTalk mailer has been received from another user (as opposed to just being added but not sent). FALSE, if the mailer has not been received from another user.</td>
</tr>
<tr class="even">
<td>31</td>
<td>The date and time the PowerTalk mailer was sent, as a serial number. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The sender name of the PowerTalk mailer, as text. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="even">
<td>33</td>
<td>The title of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>The subject of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>35</td>
<td>The author of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>The keywords for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>37</td>
<td>The comments for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>The name of the active sheet.</td>
</tr>
</tbody>
</table>

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

**Example**

The following macro formula returns the name of the active sheet in the
workbook named SALES.XLS:

GET.WORKBOOK(38, "SALES.XLS")

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

WORKBOOK.SELECT   Selects the specified documents in a workbook


GET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the Bcc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>The subject of the PowerTalk mailer, as text.</td>
</tr>
<tr class="even">
<td>29</td>
<td>The enclosures of the PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>TRUE, if the PowerTalk mailer has been received from another user (as opposed to just being added but not sent). FALSE, if the mailer has not been received from another user.</td>
</tr>
<tr class="even">
<td>31</td>
<td>The date and time the PowerTalk mailer was sent, as a serial number. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The sender name of the PowerTalk mailer, as text. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="even">
<td>33</td>
<td>The title of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>The subject of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>35</td>
<td>The author of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>The keywords for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>37</td>
<td>The comments for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>The name of the active sheet.</td>
</tr>
</tbody>
</table>

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

**Example**

The following macro formula returns the name of the active sheet in the
workbook named SALES.XLS:

GET.WORKBOOK(38, "SALES.XLS")

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

WORKBOOK.SELECT   Selects the specified documents in a workbook


GET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the Bcc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>The subject of the PowerTalk mailer, as text.</td>
</tr>
<tr class="even">
<td>29</td>
<td>The enclosures of the PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>TRUE, if the PowerTalk mailer has been received from another user (as opposed to just being added but not sent). FALSE, if the mailer has not been received from another user.</td>
</tr>
<tr class="even">
<td>31</td>
<td>The date and time the PowerTalk mailer was sent, as a serial number. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The sender name of the PowerTalk mailer, as text. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="even">
<td>33</td>
<td>The title of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>The subject of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>35</td>
<td>The author of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>The keywords for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>37</td>
<td>The comments for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>The name of the active sheet.</td>
</tr>
</tbody>
</table>

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

**Example**

The following macro formula returns the name of the active sheet in the
workbook named SALES.XLS:

GET.WORKBOOK(38, "SALES.XLS")

**Related Functions**
**Related Functions**

GET.DOCUMENT   Returns information about a workbook

WORKBOOK.SELECT   Selects the specified documents in a workbook


GET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the Bcc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>The subject of the PowerTalk mailer, as text.</td>
</tr>
<tr class="even">
<td>29</td>
<td>The enclosures of the PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>TRUE, if the PowerTalk mailer has been received from another user (as opposed to just being added but not sent). FALSE, if the mailer has not been received from another user.</td>
</tr>
<tr class="even">
<td>31</td>
<td>The date and time the PowerTalk mailer was sent, as a serial number. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The sender name of the PowerTalk mailer, as text. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="even">
<td>33</td>
<td>The title of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>The subject of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>35</td>
<td>The author of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>The keywords for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>37</td>
<td>The comments for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>The name of the active sheet.</td>
</tr>
</tbody>
</table>

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

**Example**

The following macro formula returns the name of the active sheet in the
workbook named SALES.XLS:

GET.WORKBOOK(38, "SALES.XLS")

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

WORKBOOK.SELECT   Selects the specified documents in a workbook


GET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the Bcc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>The subject of the PowerTalk mailer, as text.</td>
</tr>
<tr class="even">
<td>29</td>
<td>The enclosures of the PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>TRUE, if the PowerTalk mailer has been received from another user (as opposed to just being added but not sent). FALSE, if the mailer has not been received from another user.</td>
</tr>
<tr class="even">
<td>31</td>
<td>The date and time the PowerTalk mailer was sent, as a serial number. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The sender name of the PowerTalk mailer, as text. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="even">
<td>33</td>
<td>The title of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>The subject of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>35</td>
<td>The author of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>The keywords for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>37</td>
<td>The comments for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>The name of the active sheet.</td>
</tr>
</tbody>
</table>

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

**Example**

The following macro formula returns the name of the active sheet in the
workbook named SALES.XLS:

GET.WORKBOOK(38, "SALES.XLS")

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

WORKBOOK.SELECT   Selects the specified documents in a workbook


GET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the BccGET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the Bcc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>The subject of the PowerTalk mailer, as text.</td>
</tr>
<tr class="even">
<td>29</td>
<td>The enclosures of the PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>TRUE, if the PowerTalk mailer has been received from another user (as opposed to just being added but not sent). FALSE, if the mailer has not been received from another user.</td>
</tr>
<tr class="even">
<td>31</td>
<td>The date and time the PowerTalk mailer was sent, as a serial number. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The sender name of the PowerTalk mailer, as text. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="even">
<td>33</td>
<td>The title of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>The subject of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>35</td>
<td>The author of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>The keywords for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>37</td>
<td>The comments for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>The name of the active sheet.</td>
</tr>
</tbody>
</table>

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

**Example**

The following macro formula returns the name of the active sheet in the
workbook named SALES.XLS:

GET.WORKBOOK(38, "SALES.XLS")

**Related Functions**

GET.DOCUMENT   Returns information about a workbook

WORKBOOK.SELECT   Selects the specified documents in a workbook


GET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the BccGET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the BccGET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the BccGET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the BccGET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the BccGET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the BccGET.WORKBOOK

Returns information about a workbook.

**Syntax**

**GET.WORKBOOK**(**type\_num**, name\_text)

Type\_num    is a number that specifies what type of workbook
information you want.

<table>
<tbody>
<tr class="odd">
<td><strong>Type_num</strong></td>
<td><strong>Returns</strong></td>
</tr>
<tr class="even">
<td>1</td>
<td>The names of all sheets in the workbook, as a horizontal array of text values. Names are returned as [book]sheet.</td>
</tr>
<tr class="odd">
<td>2</td>
<td>This will always return the #N/A error value.</td>
</tr>
<tr class="even">
<td>3</td>
<td>The names of the currently selected sheets in the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="odd">
<td>4</td>
<td>The number of sheets in the workbook.</td>
</tr>
<tr class="even">
<td>5</td>
<td>TRUE if the workbook has a routing slip; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>6</td>
<td>The names of all of the workbook routing recipients who have not received the workbook, as a horizontal array of text values.</td>
</tr>
<tr class="even">
<td>7</td>
<td>The subject line for the current routing slip, as text.</td>
</tr>
<tr class="odd">
<td>8</td>
<td>The message text for the routing slip, as text.</td>
</tr>
<tr class="even">
<td>9</td>
<td>If the workbook is to be routed to recipients one after another, returns 1. If it is to be routed all at once, returns 2.</td>
</tr>
<tr class="odd">
<td>10</td>
<td>TRUE, if the Return When Done check box in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>11</td>
<td>TRUE, if the current recipient has already forwarded the current workbook; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>12</td>
<td>TRUE, if the Track Status checkbox in the Routing Slip dialog box is selected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>13</td>
<td><p>Status of the workbook routing slip:</p>
<p>0 = Unrouted</p>
<p>1 = Routing in progress, or the workbook has been routed to a user</p>
<p>2 = Routing is finished</p></td>
</tr>
<tr class="odd">
<td>14</td>
<td>TRUE, if the workbook structure is protected; otherwise, FALSE.</td>
</tr>
<tr class="even">
<td>15</td>
<td>TRUE, if the workbook windows are protected; otherwise, FALSE.</td>
</tr>
<tr class="odd">
<td>16</td>
<td>Name of the workbook as text. The workbook name does not include the drive, directory or folder, or window number.</td>
</tr>
<tr class="even">
<td>17</td>
<td>TRUE if the workbook is read only; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(34).</td>
</tr>
<tr class="odd">
<td>18</td>
<td>TRUE if sheet is write-reserved; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(35).</td>
</tr>
<tr class="even">
<td>19</td>
<td>Name of the user with current write permission for the workbook. This is the equivalent of GET.DOCUMENT(36).</td>
</tr>
<tr class="odd">
<td>20</td>
<td>Number corresponding to the file type of the document as displayed in the Save As dialog box. This is the equivalent of GET.DOCUMENT(37).</td>
</tr>
<tr class="even">
<td>21</td>
<td>TRUE if the Always Create Backup check box is selected in the Save Options dialog box; otherwise, FALSE. This is the equivalent of GET.DOCUMENT(40).</td>
</tr>
<tr class="odd">
<td>22</td>
<td>TRUE if the Save External Link Values check box is selected in the Calculation tab of the Options dialog box. This is the equivalent of GET.DOCUMENT(43).</td>
</tr>
<tr class="even">
<td>23</td>
<td>TRUE if the workbook has a PowerTalk mailer; otherwise, FALSE. Returns #N/A if no OCE mailer is installed.</td>
</tr>
<tr class="odd">
<td>24</td>
<td>TRUE if changes have been made to the workbook since the last time it was saved; FALSE if book is unchanged (or when closed, will not prompt to be saved).</td>
</tr>
<tr class="even">
<td>25</td>
<td>The recipients on the To line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>26</td>
<td>The recipients on the Cc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="even">
<td>27</td>
<td>The recipients on the Bcc line of a PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>28</td>
<td>The subject of the PowerTalk mailer, as text.</td>
</tr>
<tr class="even">
<td>29</td>
<td>The enclosures of the PowerTalk mailer, as a horizontal array of text.</td>
</tr>
<tr class="odd">
<td>30</td>
<td>TRUE, if the PowerTalk mailer has been received from another user (as opposed to just being added but not sent). FALSE, if the mailer has not been received from another user.</td>
</tr>
<tr class="even">
<td>31</td>
<td>The date and time the PowerTalk mailer was sent, as a serial number. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="odd">
<td>32</td>
<td>The sender name of the PowerTalk mailer, as text. Returns the #N/A error value if the mailer has not yet been sent.</td>
</tr>
<tr class="even">
<td>33</td>
<td>The title of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>34</td>
<td>The subject of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>35</td>
<td>The author of the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>36</td>
<td>The keywords for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="even">
<td>37</td>
<td>The comments for the workbook as displayed on the Summary tab of the Properties dialog box, as text.</td>
</tr>
<tr class="odd">
<td>38</td>
<td>The name of the active sheet.</td>
</tr>
</tbody>
</table>

Name\_text    is the name of an open workbook. If name\_text is omitted,
it is assumed to be the active workbook.

**Example**

The following macro formula returns the name of the active sheet in the
workbook named SALES.XLS:

GET.WORKBOOK(38, "SALES.XLS")

**Related Functions**

[GET.DOCUMENT](GET.DOCUMENT.md)   Returns information about a workbook

[WORKBOOK.SELECT](WORKBOOK.SELECT.md)   Selects the specified documents in a workbook


