<HTML>
<HEAD>
	<TITLE>webGab Discussion Forum</TITLE>
</HEAD>

<%
dim intTopic
intTopic = Request.QueryString("topic_id")
%>

<frameset rows="60,*,35%" border=0 framespacing=0>
	<frame name="header" src="webGab_header.asp?topic_id=<%=intTopic%>" marginwidth="8" marginheight="1" scrolling="no" frameborder="0" noresize>
	<frame name="messages" src="webGab_choose.asp?topic_id=<%=intTopic%>" scrolling="auto" marginwidth="8" marginheight="0" frameborder="0">
	<frame name="detail" src="webGab_intro.asp?topic_id=<%=intTopic%>" scrolling="auto" frameborder=no>
</frameset>

</HTML>