<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/12/1999

dim strSubject, intParent, strBody, strType, intTopic, strReturn
dim strTemp, intLength, intBreak1, intBreak2, intLine

if Request.QueryString("msg_id") <> "" then
%>
<!--#include file="data/webGab_data.inc"-->
<%
	strQuery = "SELECT * FROM gab_items WHERE (msg_id)=" _
		& Request.QueryString("msg_id")
	Set rsMsg = db.Execute(strQuery,,&H0001)

' add a "re:" to replies but make sure not to double
' them up
	
	if InStr(rsMsg("msg_subject"), "RE:") = 0 then
		strSubject = "RE: " & rsMsg("msg_subject")
	else
		strSubject = rsMsg("msg_subject")
	end if

' whatever message we're replying to becomes the parent
	
	intParent = rsMsg("msg_id")

' determine which page to show if cancel is hit
	
	strTemp = rsMsg("msg_body")
	
	if strTemp <> "" then
		strReturn = "detail.asp?msg_id=" & rsMsg("msg_id")
	else
		strReturn = "intro.asp"
	end if

' if including original message in reply, break it
' into lines that fit nicely in the form
	
	if Request.Form("include") = "on" then
		intTotal = Len(strTemp)
		intBreak2 = 0

		do while (intTotal - intBreak2) > 48
			intBreak1 = intBreak2
			intBreak2 = InStrRev(strTemp, " ", intBreak1 + 48)
			intLine = intBreak2 - intBreak1
			strBody = strBody & "> " & Mid(strTemp, intBreak1 + 1, intLine) & VbCrLf
		loop
		
		if intBreak2 < intTotal then
			strBody = strBody & "> " & Right(strTemp, intTotal - intBreak2) & VbCrLf
		end if
	end if

' if the author we're responding to wanted e-mail replies
' then let's keep record that address

	if rsMsg("msg_reply") then
		strEmail = rsMsg("msg_email")
	else
		strEmail = ""
	end if
	
	intTopic = rsMsg("msg_topic")
	
	rsMsg.Close
	Set rsMsg = nothing
	db.Close
	Set db = nothing
else
	strSubject = ""
	intParent = 0
	strReturn = "intro.asp"
	strBody = ""
	strEmail = ""
	intTopic = Request.Form("topic_id")
end if
%>

<html>
<head>
<SCRIPT LANGUAGE="javascript">
<!--
function Validate() {
	if (document.postForm.subject.value.length <= 0) {
		alert("You must enter a subject");
		document.postForm.subject.select();
		document.postForm.subject.focus();
		return false;
	}
	if (document.postForm.author.value.length <= 0) {
		alert("You must enter your name");
		document.postForm.author.select();
		document.postForm.author.focus();
		return false;
	}
	if (document.postForm.reply.checked) {
		if (document.postForm.email.value.length <= 0) {
			alert("If you wish to receive e-mailed replies you must supply your address");
			document.postForm.email.select();
			document.postForm.email.focus();
			return false;
		}
	}
	if (document.postForm.email.value.length > 0) {
		var email = document.postForm.email.value;
		if (email.indexOf("@") < 3) {
		   alert("The e-mail address you entered seems incorrect.  Please check it")
			document.postForm.email.select();
			document.postForm.email.focus();
			return false;
		}
	}
}
//-->
</SCRIPT>

<!--#include file="webGab_themes.inc"-->
</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table width="100%" cellspacing=0 cellpadding=2 border=0>
<tr>
	<form name="postForm" action="webGab_updated.asp" method="post">
	<td bgcolor="#<%=arColor(12)%>" align="right">
	<font face="Tahoma, Arial, Helvetica" size=1 color="#<%=arColor(14)%>">
	subject:</font>
	<input type="text" name="subject" value="<%=strSubject%>" size="50" maxlength="50">
	</td>
	
	<td bgcolor="#<%=arColor(12)%>" align="right">
		<input type="submit" name="cancel" value="Cancel">
		<input type="submit" name="post" value="Post" onClick="return Validate();">
	</td>
<tr>
	<td bgcolor="#<%=arColor(12)%>" valign="top" align="right">
	<textarea name="body" cols="50" rows="9" wrap="virtual"><%=strBody%></textarea></td>
	
	<td bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=1 color="#<%=arColor(14)%>">
	name:</font><br>
	<input type="text" name="author" maxlength="20"><br>
	<font face="Tahoma, Arial, Helvetica" size=1 color="#<%=arColor(14)%>">
	e-mail:</font><br>
	<input type="text" name="email" maxlength="20"><br>
	<font face="Tahoma, Arial, Helvetica" size=2>
	<input type="checkbox" name="reply">e-mail replies to me</font>
	</td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="parent_id" value="<%=intParent%>">
<input type="hidden" name="topic_id" value="<%=intTopic%>">
<input type="hidden" name="return" value="<%=strReturn%>">
<input type="hidden" name="reply_to" value="<%=strEmail%>">
</form>

</center>
</body>
</html>