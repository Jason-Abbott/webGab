<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/13/1999

dim intTopic
intTopic = Request.QueryString("topic_id")
%>

<SCRIPT LANGUAGE="javascript">
<!--
/*
if the client browser supports the "all" property
for the document object then load the javascript
version of board, otherwise load the server-
side version
*/

if (document.all) {
//	this.location = "webGab_js-list.asp?topic_id=<%=intTopic%>";
	this.location = "webGab_list.asp?topic_id=<%=intTopic%>";
} else {
	this.location = "webGab_list.asp?topic_id=<%=intTopic%>";
}
	
//-->
</SCRIPT>