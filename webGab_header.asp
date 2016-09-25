<html>
<head>
<!--#include file="webGab_themes.inc"-->
<!--#include file="data/webGab_data.inc"-->
</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">

<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/12/1999

dim intTopic, strQuery, rsTopic, strTopic

intTopic = Request.QueryString("topic_id")

strQuery = "SELECT topic_name FROM gab_topics " _
	& "WHERE (topic_id)=" & intTopic

Set rsTopic = db.Execute(strQuery,,&H0001)
strTopic = rsTopic("topic_name")
rsTopic.Close
Set rsTopic = nothing
db.Close
Set db = nothing
%>

<!-- framing table -->
<table bgcolor="#<%=arColor(6)%>" border=0 cellpadding=2 cellspacing=0 width="100%"><tr><td>
<!-- end framing table -->

<table width="100%" cellspacing=0 cellpadding=2 border=0>
<tr>
	<form action="webGab.asp" target="_top">
	<td bgcolor="#<%=arColor(12)%>">
		<input type="submit" value="Home">
	</td>
	</form>
	
	<form action="webGab_list.asp?topic_id=<%=intTopic%>" method="post" target="messages">
	<td bgcolor="#<%=arColor(12)%>">
		<input type="submit" name="expandall" value="Expand all">
		<input type="submit" name="refresh" value="Refresh">
	</td>
	</form>

	<form action="webGab_find.asp?topic_id=<%=intTopic%>" method="post" target="detail">
	<td bgcolor="#<%=arColor(12)%>">
		<input type="submit" name="find" value="Find">
	</td>
	</form>
	
	<form action="webGab_post.asp" method="post" target="detail">
	<td bgcolor="#<%=arColor(12)%>" align="right">
		<input type="submit" name="new" value="New Post">
		<input type="hidden" name="topic_id" value="<%=intTopic%>">
	</td>
	</form>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<!-- <%=strTopic%> -->

<table width="100%" cellspacing=0 cellpadding=0 border=0>
<tr>
	<td width="57%" align="center"><font face="Tahoma, Arial, Helvetica" size=2>
	<b>Subject</b></font></td>
	<td align="center"><font face="Tahoma, Arial, Helvetica" size=2>
	<b>Author</b></font></td>
	<td align="center"><font face="Tahoma, Arial, Helvetica" size=2>
	<b>Time</b></font></td>
	<td align="center"><font face="Tahoma, Arial, Helvetica" size=2>
	<b>Re's</b></font></td>
</table>

</body>
</html>