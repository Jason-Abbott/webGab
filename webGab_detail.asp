<html>
<head>
<!--#include file="webGab_themes.inc"-->
<!--#include file="data/webGab_data.inc"-->
<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/13/1999

dim strQuery, rsMsg, strTime, strAuthor

strQuery = "SELECT * FROM gab_items WHERE (msg_id)=" _
	& Request.QueryString("msg_id")
Set rsMsg = db.Execute(strQuery,,&H0001)

strTime = Left(rsMsg("msg_date"), Len(rsMsg("msg_date"))-6) _
	& Right(rsMsg("msg_date"), 3)

if rsMsg("msg_email") <> "" then
	strAuthor = "<a href=""mailto:" & rsMsg("msg_email") _
		& """>" & rsMsg("msg_author") & "</a>"
else
	strAuthor = rsMsg("msg_author")
end if

' if a new message was posted then generate the JavaScript
' that will refresh the message list

if Request.QuerySTring("refresh") then
	strScript = "<script language=""javascript""><!--" & VbCrLf _
		& "parent.frames[1].location.href = ""webGab_list.asp?topic_id=" _
		& rsMsg("msg_topic") & """" & VbCrLF & "--></script>"
else
	strScript = ""
end if
%>

</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" border=0 cellpadding=2 cellspacing=0 width="100%"><tr><td>
<!-- end framing table -->



<table width="100%" cellspacing=0 cellpadding=2 border=0>
<tr>
	<form action="webGab_post.asp?msg_id=<%=rsMsg("msg_id")%>" method="post">
	<td bgcolor="#<%=arColor(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2>
	<%=rsMsg("msg_subject")%></font></td>
	
<!-- 	<td bgcolor="#<%=arColor(12)%>" align="center">
	<font face="Tahoma, Arial, Helvetica" size=2>
	<%=strAuthor%></font></td>
	
	<td bgcolor="#<%=arColor(12)%>" align="center">
	<font face="Tahoma, Arial, Helvetica" size=2>
	<%=strTime%></font></td> -->
	
	<td bgcolor="#<%=arColor(12)%>" align="right" valign="center">
	<font face="Tahoma, Arial, Helvetica" size=1 color="#<%=arColor(14)%>">
	<input type="checkbox" name="include">include original message in</font>
	<input type="submit" name="reply" value="Reply">
	</td>
	</form>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<tt>
<%=Replace(rsMsg("msg_body"), VbCrLf, "<br>")%>
</tt>
<p align="right">
<font face="Tahoma, Arial, Helvetica" size=2>
<%=strAuthor%><br>
<%=strTime%>
</font>

<%
rsMsg.Close
Set rsMsg = nothing
db.Close
Set db = nothing
%>

</body>
<%=strScript%>
</html>