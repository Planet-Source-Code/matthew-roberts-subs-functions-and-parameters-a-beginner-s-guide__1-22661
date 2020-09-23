<div align="center">

## Subs, Functions, and Parameters\.\.\.A beginner's guide\.


</div>

### Description

Explains the concepts of Subs, Functions, and Parameters. If you are a little fuzzy on the difference, take a look.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-roberts.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-roberts-subs-functions-and-parameters-a-beginner-s-guide__1-22661/archive/master.zip)





### Source Code

```
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 9">
<meta name=Originator content="Microsoft Word 9">
<link rel=File-List
href="./Sometimes%20we%20miss%20the%20obvious_files/filelist.xml">
<title>Sometimes we miss the obvious</title>
<style>
<!--
 /* Font Definitions */
@font-face
	{font-family:Wingdings;
	panose-1:5 0 0 0 0 0 0 0 0 0;
	mso-font-charset:2;
	mso-generic-font-family:auto;
	mso-font-pitch:variable;
	mso-font-signature:0 268435456 0 0 -2147483648 0;}
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:9.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Arial;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-weight:bold;}
p.MsoTitle, li.MsoTitle, div.MsoTitle
	{margin:0in;
	margin-bottom:.0001pt;
	text-align:center;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:Arial;
	mso-fareast-font-family:"Times New Roman";
	font-weight:bold;
	mso-bidi-font-weight:normal;}
p.MsoBodyTextIndent, li.MsoBodyTextIndent, div.MsoBodyTextIndent
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:.25in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:9.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Arial;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-weight:bold;}
@page Section1
	{size:8.5in 11.0in;
	margin:1.0in 1.25in 1.0in 1.25in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
 /* List Definitions */
@list l0
	{mso-list-id:236324076;
	mso-list-type:hybrid;
	mso-list-template-ids:655414104 67698703 67698713 67698715 67698703 67698713 67698715 67698703 67698713 67698715;}
@list l0:level1
	{mso-level-tab-stop:.75in;
	mso-level-number-position:left;
	margin-left:.75in;
	text-indent:-.25in;}
@list l1
	{mso-list-id:711737014;
	mso-list-type:hybrid;
	mso-list-template-ids:614887728 67698703 67698713 67698715 67698703 67698713 67698715 67698703 67698713 67698715;}
@list l1:level1
	{mso-level-tab-stop:.75in;
	mso-level-number-position:left;
	margin-left:.75in;
	text-indent:-.25in;}
@list l2
	{mso-list-id:1211503777;
	mso-list-type:hybrid;
	mso-list-template-ids:200678504 67698689 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l2:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:.5in;
	mso-level-number-position:left;
	text-indent:-.25in;
	font-family:Symbol;}
@list l3
	{mso-list-id:1282806297;
	mso-list-type:hybrid;
	mso-list-template-ids:-79284828 67698689 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}
@list l3:level1
	{mso-level-number-format:bullet;
	mso-level-text:\F0B7;
	mso-level-tab-stop:.5in;
	mso-level-number-position:left;
	text-indent:-.25in;
	font-family:Symbol;}
ol
	{margin-bottom:0in;}
ul
	{margin-bottom:0in;}
-->
</style>
</head>
<body lang=EN-US style='tab-interval:.5in'>
<div class=Section1>
<p class=MsoTitle>Using Subs and Functions in VB</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Sometimes we miss the obvious. I know there are a lot of
things I was doing in Visual Basic that I thought were not only right, but
extremely clever. What I eventually discovered was that the reason I was having
to dream up such clever &#8220;work abounds&#8221; was that I was basing my code on
incorrect assumptions about how VB works. </p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>One area that I had trouble with, as have several people I
have taught coding to is parameters and return values. This article is in the
Beginner category, but<span style="mso-spacerun: yes">  </span>I didn&#8217;t learn
some of this stuff until long after I was working as a VB programmer and
considered myself a &#8220;professional&#8221;. If there is one thing I have learned about
programming, it is that you NEVER know everything, and there is always another
way to do something. My goal is to show programmers the correct way to use
these features before they go off and invent ways that will cause problems
later. Nothing is worse than realizing that you have been doing something the
wrong way for the last three years.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Is this tutorial for you? Let&#8217;s find out. Consider the
following questions:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.5in;text-indent:-.25in;mso-list:l2 level1 lfo1;
tab-stops:list .5in'><![if !supportLists]><span style='font-family:Symbol'>·<span
style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><![endif]>Can you explain the difference between a Sub and a
Function in one sentence?</p>
<p class=MsoNormal style='margin-left:.5in;text-indent:-.25in;mso-list:l2 level1 lfo1;
tab-stops:list .5in'><![if !supportLists]><span style='font-family:Symbol'>·<span
style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><![endif]>Can you explain the difference between a PUBLIC Sub and
a PRIVATE sub?</p>
<p class=MsoNormal style='margin-left:.5in;text-indent:-.25in;mso-list:l2 level1 lfo1;
tab-stops:list .5in'><![if !supportLists]><span style='font-family:Symbol'>·<span
style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><![endif]>Do you know how to get a return value from a procedure without
using global variables?</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>If you answered &#8220;Yes&#8221; to all of
the questions above, you probably don&#8217;t need this tutorial. But before you
close it out, you should be sure you <i>really</i> understand how to do these
things. Why? Because these skills are fundamental to good coding. You simply
cannot write good complex code without them. </p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>If the questions above left you
scratching your head in confusion, hang in there. You are not alone. If you read
this entire tutorial, I promise they will be answered in a way that you can
understand them (or your money back!)</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'><b style='mso-bidi-font-weight:
normal'>Subs and Functions- What is the difference?</b></p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>To understand what a sub or
function is and why we need them, we should go back a few years. For the
moment, we will concentrate on Subs, since they are easier to understand. I
will then show you how Functions extend the capability of Subs.</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>If there are any old line
programmers out there, you may remember what we now call &#8220;spaghetti code&#8221;. Of course,
when we were writing it, we didn&#8217;t call it that. We called it sheer genius. The
term &#8220;spaghetti code&#8221; refers to the logic path of an application written in a
language such as BASIC or BASICA. In these languages, each line had a line
number, and you controlled program flow by referring to the associated line
number of the command you wanted to execute. (Boy, I am suddenly feeling old
here!)</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>For example:</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>10 CLS<span style="mso-spacerun:
yes">  </span><span style='mso-tab-count:5'>                                                                        </span><span
style="mso-spacerun: yes">   </span>&#8216;Clears the screen&#8230;we are talking DOS here.</p>
<p class=MsoNormal style='margin-left:.25in'>20 LN$=INPUT$ &#8220;What is your last
name?:&#8221;</p>
<p class=MsoNormal style='margin-left:.25in'>30 FN$=INPUT$ &#8220;What is your first
name?</p>
<p class=MsoNormal style='margin-left:.25in'>40 IF LN$=&#8221;&#8221; Then <b
style='mso-bidi-font-weight:normal'>GOTO</b> <b style='mso-bidi-font-weight:
normal'>20<o:p></o:p></b></p>
<p class=MsoNormal style='margin-left:.25in'>50 IF FN$=&#8221;&#8221; THEN <b
style='mso-bidi-font-weight:normal'>GOTO 30</b></p>
<p class=MsoNormal style='margin-left:.25in'>60 PRINT &#8220;You entered &#8220; + LN$ + &#8220;
&#8220; + FN$</p>
<p class=MsoNormal style='margin-left:.25in'>70 END</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>This is a complete (albeit simple)
program for BASIC. It prompts for a last name and a first name, then forces you
to re-enter it if you didn&#8217;t enter a value for one or the other.<span
style="mso-spacerun: yes">  </span>The part I would like you to notice is the
GOTO statement. You may have seen this command used in VB in error handling,
and if you have been around any VB programmers very long, you have heard them
harp on how evil the GOTO command is. </p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>In BASIC, GOTO redirected program
flow to a specific line number. Otherwise, programs started at line 0 (or 10 in
most cases) and executed sequentially until they met an END statement, which
terminated the program. The problem with this concept was that once you got
about 2000 lines of code, it became difficult to track where the program would
jump to in any situation. For example, the statement GOTO 20150 meant that to
track execution, you had to scroll all the way down to line 20150. Of course,
it may contain an IF THEN statement that sent it back to line number 1220. It
isn&#8217;t hard to imagine why this type of code earned its nickname. Eventually it
reached &#8220;critical mass&#8221; and became unmanageable.</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>Most BASIC programmers eventually
discovered the GOSUB statement. This was a pretty big leap in program control.
Like GOTO, it redirected program execution to a line number somewhere in the
thousands of lines of code, but unlike GOTO, GOSUB knew where it was redirected
<i>from</i> and could return to that point in the program when it completed its
task. It told BASIC &#8220;GOTO A BLOCK OF SUB-CODE&#8221;. Here is how it was used:<br
style='mso-special-character:line-break'>
<![if !supportLineBreakNewLine]><br style='mso-special-character:line-break'>
<![endif]></p>
<p class=MsoNormal style='margin-left:.25in'>10 CLS</p>
<p class=MsoNormal style='margin-left:.25in'>20 X$=INPUT$ &#8220;Choose a Menu Item&#8221;</p>
<p class=MsoNormal style='margin-left:.25in'>30 IF X$=&#8221;A&#8221; THEN <b
style='mso-bidi-font-weight:normal'>GOSUB</b> <b style='mso-bidi-font-weight:
normal'>2000<o:p></o:p></b></p>
<p class=MsoNormal style='margin-left:.25in'>40 IF X$=&#8221;B&#8221; THEN <b
style='mso-bidi-font-weight:normal'>GOSUB 3000<o:p></o:p></b></p>
<p class=MsoNormal style='margin-left:.25in'>50 IF X$=&#8221;C&#8221; THEN END</p>
<p class=MsoNormal style='margin-left:.25in'>60 &#8217; <span style='mso-tab-count:
5'>                                                                                </span><span
style="mso-spacerun: yes">  </span>END OF MENU CODE</p>
<p class=MsoNormal style='margin-left:.25in'><b style='mso-bidi-font-weight:
normal'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></b></p>
<p class=MsoNormal style='margin-left:.25in'><b style='mso-bidi-font-weight:
normal'>&#8230;<o:p></o:p></b></p>
<p class=MsoNormal style='margin-left:.25in'><b style='mso-bidi-font-weight:
normal'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></b></p>
<p class=MsoNormal style='margin-left:.25in'>2000 CLS</p>
<p class=MsoNormal style='margin-left:.25in'>2010 &#8230;..DO STUFF&#8230;.</p>
<p class=MsoNormal style='margin-left:.25in'>&#8230;</p>
<p class=MsoNormal style='margin-left:.25in'>2520 <b style='mso-bidi-font-weight:
normal'>RETURN<o:p></o:p></b></p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>3000 CLS</p>
<p class=MsoNormal style='margin-left:.25in'>3010 &#8230;..DO STUFF&#8230;.</p>
<p class=MsoNormal style='margin-left:.25in'>&#8230;</p>
<p class=MsoNormal style='margin-left:.25in'>3520 <b style='mso-bidi-font-weight:
normal'>RETURN<o:p></o:p></b></p>
<p class=MsoNormal style='margin-left:.25in'><b style='mso-bidi-font-weight:
normal'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></b></p>
<p class=MsoNormal style='margin-left:.25in'><b style='mso-bidi-font-weight:
normal'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></b></p>
<p class=MsoNormal style='margin-left:.25in'>In this example, we not only use
the <b style='mso-bidi-font-weight:normal'>GOSUB</b> keyword, but we also use
the <b style='mso-bidi-font-weight:normal'>RETURN </b>keyword. What this told
BASIC was <i>&#8220;Go back to the last GOSUB statement that you executed. &#8220;</i> To
programmers, this was pure gold. It allowed recursion (calling a piece of code
within the same piece of code) and much better program flow control. In
essence, this is the origin of the Visual Basic &#8220;Sub&#8221; procedure. </p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>OK&#8230;enough nostalgia for now. Back
to VB. What this history lesson has shown you is that things could be much
worse. But with Visual Basic, some obvious improvements have been made. Line
numbers have been dropped (a tough concept for us old timers&#8230;QuickBasic helped
ease us into this concept). This allowed the programmer to name blocks of code
with a name instead of a number. Now instead of GOSUB 2000, we can say &#8220;Call
GetCustomerID&#8221; to call a piece of code that will find a customer ID for us.</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'><b style='mso-bidi-font-weight:
normal'>Here is how the VB Sub works:<o:p></o:p></b></p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>First, you need to know what your
sub will be doing. There are some basic criteria to determine whether or not
you need to place code into a sub. They are:</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in;text-indent:0in;mso-list:l0 level1 lfo2;
tab-stops:list .75in'><![if !supportLists]>1.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><![endif]>Is this code used in more than one place (are you writing
duplicate code in your application)?</p>
<p class=MsoNormal style='margin-left:.25in;text-indent:0in;mso-list:l0 level1 lfo2;
tab-stops:list .75in'><![if !supportLists]>2.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><![endif]>Does this code perform a specialized function independent of
the rest of the code?</p>
<p class=MsoNormal style='margin-left:.25in;text-indent:0in;mso-list:l0 level1 lfo2;
tab-stops:list .75in'><![if !supportLists]>3.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><![endif]>Can you effectively create this as a &#8220;stand-alone&#8221; piece of
code?</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>Of the three criteria above,
number three is the toughest. For example, if you have a piece of code to find
state that a customer is located in based on their zip code, do you write that
same piece of code for every customer? Obviously, that would be impractical if
not impossible. What you need is a way to find the state <i>any </i>zip code.
Here is the way most beginners handle this problem:</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>First, create public variables to
hold the values we will be working with:</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'><span style='color:navy'>Public
ZIP As String<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='color:navy'>Public
State As String<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='color:navy'>Public
Sub GetState()<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='color:navy'><span
style="mso-spacerun: yes">      </span>If ZIP &gt; 32501 AND ZIP<span
style="mso-spacerun: yes">  </span>&lt;34205 Then<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='color:navy'><span
style="mso-spacerun: yes">          </span><span style="mso-spacerun:
yes"> </span>State = &#8220;MS&#8221;<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='color:navy'><span
style="mso-spacerun: yes">      </span>ElseIf ZIP &gt;45102 AND ZIP &lt; 53210
Then<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='color:navy'><span
style="mso-spacerun: yes">            </span>State = &#8220;TN&#8221;<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='color:navy'><span
style="mso-spacerun: yes">      </span>&#8230;&#8230;<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='color:navy'>End Sub</span></p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>This code works. Using this
method, you can get the state of any of the zip codes contained in the GetState
sub. But there are problems with this code as well. The main one is that you
are now relying on public variables. This is because you have to be able to
access the values in from your form or calling code <b style='mso-bidi-font-weight:
normal'>and</b> within your sub. This can get real messy real quick when you
consider that you may need to use the name &#8220;state&#8221; for a variable many times in
an application. You are then forced to create MANY public variables with odd
names like GetStateFromZip_State and GetStateFromZip_Zip to insure that you
don&#8217;t accidentally overwrite your values from other places in your program.
This is just a really bad way to code. The solution? <b style='mso-bidi-font-weight:
normal'>Parameters (</b>finally!).</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>In order to get your values safely
to your Sub without having to create public variables, you can instead create
Sub Parameters. These are really just variables that only your calling code and
your sub can see.<span style="mso-spacerun: yes">  </span>They look like this:</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'>Public Sub GetState(ZIP As String)<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><span style="mso-spacerun: yes">     
</span>If ZIP &gt; 32501 AND ZIP<span style="mso-spacerun: yes"> 
</span>&lt;34205 Then<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><span style="mso-spacerun:
yes">           </span>State = &#8220;MS&#8221;<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><span style="mso-spacerun: yes">     
</span>ElseIf ZIP &gt;45102 AND ZIP &lt; 53210 Then<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><span style="mso-spacerun:
yes">            </span>State = &#8220;TN&#8221;<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><span style="mso-spacerun: yes">     
</span>&#8230;&#8230;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'><span style="mso-spacerun: yes">              </span>End If<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'>End Sub</span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyTextIndent>Now this code does the exact same thing, but without
having to rely on the public variable ZIP. You can also pass multiple
parameters:</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'>Public Sub AverageNumbers(Number1 As
Integer, Number2 As Integer, Number3 As Integer)<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><span style='mso-tab-count:1'>        </span>&#8230;<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'>End Sub</span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>All three of these values will be
available from within your sub, but will not exist outside of it. Getting
pretty neat, isn&#8217;t it?</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyTextIndent>But we still have a problem. We passed the variable
IN , but how do we get a value back OUT of a sub? I mean, it&#8217;s nice that we
averaged these numbers, but we still have to use a public variable to get the
return value, right? Wrong. </p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'>There are three ways to get return
values from a piece of code:</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.75in;text-indent:-.25in;mso-list:l1 level1 lfo3;
tab-stops:list .75in'><![if !supportLists]>1.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><![endif]>Public Variables (ugly!)</p>
<p class=MsoNormal style='margin-left:.75in;text-indent:-.25in;mso-list:l1 level1 lfo3;
tab-stops:list .75in'><![if !supportLists]>2.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><![endif]>By making a &#8220;return value&#8221; parameter.</p>
<p class=MsoNormal style='margin-left:.75in;text-indent:-.25in;mso-list:l1 level1 lfo3;
tab-stops:list .75in'><![if !supportLists]>3.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><![endif]>By making your sub into a function.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyTextIndent>We have already established that public variables
are not the answer we are seeking, so lets examine option #2. This is more of a
&#8220;hack&#8221; than a feature of VB. It takes advantage of the fact that both the
calling code and the sub code have access to parameter values. You could do
this to get a return value:</p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'>Public Sub AverageNumbers(Number1 As
Integer, Number2 As Integer, Number3 As Integer, ReturnValue As Integer)<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><span style='mso-tab-count:1'>        </span><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><span style="mso-spacerun: yes">    
</span>ReturnValue = (Number1 + Number2 + Number3) /3<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'>End Sub</span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes"> </span>Again, this would
work. But it creates problems on the calling side now. In order to use it, you
have to use code similar to this:<br style='mso-special-character:line-break'>
<![if !supportLineBreakNewLine]><br style='mso-special-character:line-break'>
<![endif]></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'><span style="mso-spacerun: yes">      </span>Call
AverageNumbers(10, 20, 50,0)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'><span style="mso-spacerun: yes">      </span>Msgbox &#8220;Average =
&#8220;<span style="mso-spacerun: yes">  </span>&amp; ReturnValue</span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>It is technically returning a value, but it looks (and is) rather
clunky. You have to pass in a &#8220;0&#8221; for the return value or you will get an
error.<span style="mso-spacerun: yes">  </span>Then, to make matters worse, you
have to then refer to that variable when the sub completes. What would be nice
is if you could use it just like a VB command.<span style="mso-spacerun: yes"> 
</span>Consider the <i>Ucase()</i> command in Visual Basic. It is
straightforward:</p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'><span style="mso-spacerun: yes">       </span>MsgBox Ucase(&#8220;this is
a test&#8221;)<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes">      </span>Results is a
message box being displayed with the words &#8220;THIS IS A TEST&#8221; in it.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>So how does Microsoft do that? How do they get a return
value from the Ucase command? Surely it is some secret code that you don&#8217;t have
the power to duplicate. <b style='mso-bidi-font-weight:normal'>Wrong</b>! The
way they do it is by making the command Ucase() into a <b style='mso-bidi-font-weight:
normal'>Function</b> instead of a Sub. What is the difference? Simple:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>A function can return
a value through its name variable. </b><span style="mso-spacerun:
yes"> </span>A sub cannot. </p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>What does that mean &#8220;through its name variable&#8221;?</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Let&#8217;s look at a function declaration:</p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'>Public Function AverageNumbers(Number1 As Integer, Number2 As
Integer, Number3 As Integer) </span><b style='mso-bidi-font-weight:normal'><span
style='font-size:8.0pt;mso-bidi-font-size:12.0pt;color:navy'>As Integer</span></b><span
style='font-size:8.0pt;mso-bidi-font-size:12.0pt;color:navy'><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'>End Function</span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>You should notice two things different about this
declaration compared to the sub declaration:</p>
<p class=MsoNormal style='margin-left:.5in;text-indent:-.25in;mso-list:l3 level1 lfo4;
tab-stops:list .5in'><![if !supportLists]><span style='font-family:Symbol'>·<span
style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><![endif]>It has no return parameter.</p>
<p class=MsoNormal style='margin-left:.5in;text-indent:-.25in;mso-list:l3 level1 lfo4;
tab-stops:list .5in'><![if !supportLists]><span style='font-family:Symbol'>·<span
style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><![endif]>It has &#8220;As Integer&#8221; stuck on the end. What is <b
style='mso-bidi-font-weight:normal'><i>that</i></b> all about?</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>So how in the world does the value of Ucase() (or
AverageNumbers for that matter), find its way back to the calling code &#8220;MsgBox<span
style="mso-spacerun: yes">  </span>= &#8220;?</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Well, <u>that</u> is the difference between Subs and
Functions. Functions act as a <i>return value </i><b style='mso-bidi-font-weight:
normal'><i>variable</i></b><i>. </i>That is why it was declared &#8220;<b
style='mso-bidi-font-weight:normal'>As Integer&#8221;</b>. You can now call the
AverageNumber function exactly as you would the Ucase() function:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes">          </span><span
style='color:navy'>MsgBox &#8220;Average =&#8221; &amp; AverageNumbers(10,20,50)<o:p></o:p></span></p>
<p class=MsoNormal><span style='color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style="mso-spacerun: yes"> </span>Here is the complete
function:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'>Public Function AverageNumbers(Number1 As Integer, Number2 As
Integer, Number3 As Integer) As Integer<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><span style="mso-spacerun: yes">   
</span>AverageNumbers = (Number1 + Number2 + Number3) /3<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'>End Function<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Notice that the line:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><span style="mso-spacerun: yes">    
</span>ReturnValue = (Number1 + Number2 + Number3) /3<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.5in;text-indent:.5in'>has now been
modified to read:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><span style="mso-spacerun: yes">   
</span>AverageNumbers = (Number1 + Number2 + Number3) /3<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>This is where the &#8220;magic&#8221; happens. VB performs the
calculations, and when it gets the results, it pipes it back into the
AverageNumbers <i>variable</i> which was created when this function was
declared. From there, it can be assigned back to a variable in the calling
code. So to see the whole picture:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes">   </span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'>Private Sub Command1_Click()<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'><span style="mso-spacerun: yes">      </span>MsgBox &#8220;Average =&#8221; &amp;
AverageNumbers(10,20,50)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'>End Sub<o:p></o:p></span></p>
<p class=MsoNormal><span style="mso-spacerun: yes"> </span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'>Public Function AverageNumbers(Number1 As Integer, Number2 As
Integer, Number3 As Integer) As Integer<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.25in'><span style='font-size:8.0pt;
mso-bidi-font-size:12.0pt;color:navy'><span style="mso-spacerun: yes">   
</span>AverageNumbers = (Number1 + Number2 + Number3) /3<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:8.0pt;mso-bidi-font-size:12.0pt;
color:navy'>End Function<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>So now you can start to see how return values work with
functions. Once you fully grasp this concept, you will begin to realize that VB
commands are nothing but functions written by the Microsoft VB team.
Conceptually, yours are no different. You can actually <b style='mso-bidi-font-weight:
normal'>create your own commands</b> to us in your application, just by
creating them as functions!</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>This is big!</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>There is one more concept I want to touch on before wrapping
up:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes"> </span><b style='mso-bidi-font-weight:
normal'>The difference between Public and Private Subs and Functions<o:p></o:p></b></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>To fully understand the implications of making a function or
sub public or private, you will need to study up on the topic of SCOPE in VB.
This basically is the level at which something is &#8220;visible&#8221; to the rest of the
application. As you have seen through earlier examples, Public variables are
not a good thing. Public Functions, on the other hand, are a VERY good thing.
This makes them accessible from anywhere in your application&#8230;so if you want to
average a number from three different forms, you can still call the same
AverageNumbers function. </p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>With Private subs and functions, you will only be able to
call them from <i>within the object that they are declared in.</i> Why would
you want to do this? Well, you may have noticed that VB creates all Form Subs
as Private. This is because if you created them as Public, you would have many
Form_Load() subs, many<span style="mso-spacerun: yes">  </span>Command1_Click
() subs, etc. This would make your application crash instantly, so by using
private scope, you effectively &#8220;hide&#8221; these subs from other forms.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>ONE NOTE: You cannot declare a Public variable, sub, or
function from within a form. <b style='mso-bidi-font-weight:normal'>YOU MUST
DECLARE ALL PUBLIC ITEMS FROM WITHIN A MODULE. </b>You can add a module to your
project by going to the Project menu and clicking &#8220;Add Module&#8221;.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>I sincerely hope this tutorial has helped you in grasping
how and why to use subs and functions. This is such a vital topic to good
programming and so little is published about it. Please let me know if you need
more information on any of the topics covered in this tutorial.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>M@</p>
</div>
</body>
</html>
```

