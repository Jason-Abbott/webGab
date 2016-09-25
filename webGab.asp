<!-- 
Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
Last updated 08/11/1999
-->

<html>
<head>
<title>webGab Message Board</title>

<!--#include file="webGab_themes.inc"-->
</head>
<body bgcolor="#<%=arColor(1)%>" text="#<%=arColor(7)%>" link="#<%=arColor(8)%>" vlink="#<%=arColor(8)%>" alink="#<%=arColor(9)%>">

<font face="Tahoma, Arial, Helvetica" size=5><b>
Message Boards</b></font>
<br>

<table width="100%" cellspacing=0 cellpadding=2 border=0>
<tr>
	<td width="75%"></td>
	<td align="center" valign="bottom"><font face="Tahoma, Arial, Helvetica" size=1>
	Moderator</font></td>
	<td align="center"><font face="Tahoma, Arial, Helvetica" size=1>
	Total<br>Messages</font></td>
	<td align="center"><font face="Tahoma, Arial, Helvetica" size=1>
	Today's<br>Messages</font></td>

<!--#include file="data/webGab_data.inc"-->
<!--#include file="show_status.inc"-->
<%
dim strQuery, rsTopics, rsMsgs, intTotal, intNew, strMod

strQuery = "SELECT * FROM gab_topics T LEFT OUTER JOIN gab_users U " _
	& "ON (T.topic_moderator = U.user_id) ORDER BY topic_name"

Set rsTopics = db.Execute(strQuery,,&H0001)
do while not rsTopics.EOF
	response.write "<tr>" & VbCrLf & "<td bgcolor=""#" _
		& arColor(2) & """><font face=""Verdana, Arial, Helvetica"" size=2><b>" _
		& "<a href=""webGab_go.asp?topic_id=" & rsTopics("topic_id") & """ " _
		& showStatus("View '" & rsTopics("topic_name") & "' messages") & ">" _
		& rsTopics("topic_name") & "</a></b></td>" & VbCrLf

	strQuery = "SELECT msg_date, msg_topic FROM gab_items " _
		& "WHERE msg_topic=" & rsTopics("topic_id")

	intTotal = 0
	intNew = 0	
		
	Set rsMsgs = db.Execute(strQuery,,&H0001)
	do while not rsMsgs.EOF
		if DateDiff("h", rsMsgs("msg_date"), Now) < 24 then
			intNew = intNew + 1
		end if
		intTotal = intTotal + 1
		rsMsgs.MoveNext
	loop
	
	if rsTopics("user_first") <> "" then
		strMod = rsTopics("user_first")
	else
		strMod = "none"
	end if

	response.write "<td align=""center"" bgcolor=""#" & arColor(2) _
		& """><font face=""Tahoma, Arial, Helvetica"" size=2>" _
		& strMod & "</font></td>" & VbCrLf _
		& "<td align=""center"" bgcolor=""#" & arColor(2) _
		& """><font face=""Tahoma, Arial, Helvetica"" size=2>" _
		& intTotal & "</font></td>" & VbCrLf _
		& "<td align=""center"" bgcolor=""#" & arColor(2) _
		& """><font face=""Tahoma, Arial, Helvetica"" size=2>" _
		& intNew & "</font></td>" & VbCrLf _
		& "<tr><td colspan=3><font face=""Tahoma, Arial, Helvetica"" size=2>" _
		& "<blockquote>" & rsTopics("topic_description") _
		& "</font></td></blockquote>" & VbCrLf
		

	rsTopics.MoveNext
loop

Set rsMsgs = nothing

rsTopics.Close
Set rsTopics = nothing

db.Close
Set db = nothing
%>

</table>

</body>
</html>