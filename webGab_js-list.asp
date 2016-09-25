<html>
<head>

<!-- 
Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
Last updated 08/13/1999
-->

<!--#include file="webGab_themes.inc"-->
<SCRIPT LANGUAGE="JavaScript">
<!--
// cache graphics for fast display

plus = new Image(); plus.src = "./images/plus_<%=strStyle%>.gif";
minus = new Image(); minus.src = "./images/minus_<%=strStyle%>.gif";
blank = new Image(); blank.src = "./images/blank_<%=strStyle%>.gif";

// change style visibility and image

function ToggleDisplay(oButton, oItems)
{

	if ((oItems.style.display == "") || (oItems.style.display == "none"))	{
		oItems.style.display = "block";
		oButton.src = "/msdn-online/start/images/minus.gif";
	}	else {
		oItems.style.display = "none";
		oButton.src = "/msdn-online/start/images/plus.gif";
	}
}


function toggleIt(id,reply) {
	var allTRs = document.all.tags("TR");
	if (eval("id" + reply + ".style.display") == "none") {
		for (i=0; i<allTRs.length; i++) {
			if (allTRs(i).className == "reply_" + id) {
				allTRs(i).style.display = "block";
				eval("document." + id + "_img.src=minus.src");
			}
		}
	} else {
		for (i=0; i<allTRs.length; i++) {
			if (allTRs(i).className == "reply_" + id) {
				allTRs(i).style.display = "none";
				eval("document." + id + "_img.src=plus.src");
			}
		}
	}
}
//-->
</SCRIPT>
<STYLE TYPE="text/css">
td	{
	font-family : Tahoma, Arial, Helvetica;
	font-size : 10pt;
   }

.root  {	background-color : #<%=arColor(2)%>; }
</STYLE>
</head>
<body bgcolor="#<%=arColor(1)%>" text="#<%=arColor(7)%>" link="#<%=arColor(8)%>" vlink="#<%=arColor(8)%>" alink="#<%=arColor(9)%>">

<table width="100%" cellspacing=0 cellpadding=0 border=0>

<%
dim item, divList, intTopic

intTopic = Request.QueryString("topic_id")
%>
<!--#include file="data/webGab_array.inc"-->
<%
' 0 id
' 1 subject
' 2 author
' 3 date
' 4 parent id
' 5 depth in tree
' 6 number of replies
' 7 new replies?

for x = 0 to UBound(arMsgs,1)
	response.write "<tr id=""id" & arMsgs(x,0) & """ class="""

	if arMsgs(x,4) <> 0 then
		response.write "reply_id" & arMsgs(x,4) & """>"
	else
		response.write "root"">"
	end if
		
	response.write VbCrLf & "<td><nobr>"
	for d = 1 to (arMsgs(x,5))
		response.write "<img src=""./images/blank_" & strStyle & ".gif"">"
	next

' JavaScript doesn't accept variable names that begin
' with a number so prepend "id"

	if arMsgs(x,6) > 0 then
		response.write "<a href='#' onClick=" _
			& """toggleIt('id" & arMsgs(x,0) & "'," & arMsgs(x+1,0) & "); return true;"">" _
			& "<img name=""id" & arMsgs(x,0) & "_img"" src=""./images/plus_" _
			& strStyle & ".gif"" border=0></a>" & VbCrLf _
			& arMsgs(x,1) & VbCrLf & "</nobr></td>" & VbCrLf
	else
		response.write "<img src=""./images/blank_" _
			& strStyle & ".gif"">" & VbCrLf & arMsgs(x,1) _
			& "</nobr></td>" & VbCrLf
	end if
	
	response.write "<td align=""center"">" & arMsgs(x,2) & "</td>" & VbCrLf _
		& "<td>" & arMsgs(x,3) & "</td>" & VbCrLf _
		& "<td align=""right"">" & arMsgs(x,6) & "</td>" & VbCrLf & VbCrLf
	
next
%>
</div>
</table>
</body>
<SCRIPT LANGUAGE="JavaScript">
<!--
var allTRs = document.all.tags("TR");
for (i=0; i<allTRs.length; i++) {
	if (allTRs(i).className.indexOf("reply") != -1) {
		allTRs(i).style.display = "none";
	}
}
//-->
</SCRIPT>
</html>