<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/12/1999

dim strScript

' if a new message was posted then generate the JavaScript
' that will refresh the message list

if Request.QuerySTring("refresh") then
	strScript = "<script language=""javascript""><!--" & VbCrLf _
		& "parent.frames[1].location.href = ""webGab_list.asp?topic_id=" _
		& Request.QueryString("topic_id") & """" & VbCrLf & "--></script>"
else
	strScript = ""
end if
%>

<html>
<head>
<!--#include file="webGab_themes.inc"-->
</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">

<font face="Tahoma, Arial, Helvetica"><font color="#990000"><b>
Beta version</b></font>
<br>
<font size=2>
Welcome to the pre-release version of webGab.  Please <a href="mailto:jabbott@uidaho.edu">let me know</a> if you encounter any error messages.  Here are some usage tips:
<li>There is no need to use "n/m" or variations thereof to indicate a posting without a message.  webGab will detect such messages and highlight them with an asterisk (*).
<li>Messages posted in the last 24 hours have their time highlighted in another color.
<li>Messages with replies posted in the last 24 hours have their reply count highlighted in another color.

</body>
<%=strScript%>
</html>