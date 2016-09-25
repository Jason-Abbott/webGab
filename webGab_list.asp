<!-- 
Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
Last updated 08/12/1999
-->

<html>
<head>
<!--#include file="webGab_themes.inc"-->
</head>
<body bgcolor="#<%=arColor(1)%>" text="#<%=arColor(7)%>" link="#<%=arColor(8)%>" vlink="#<%=arColor(8)%>" alink="#<%=arColor(9)%>">

<font face="Tahoma, Arial, Helvetica" size=2>
<table width="100%" cellspacing=0 cellpadding=0 border=0>
	
<%
dim strExpand, intTopic, strFont

intTopic = Request.QueryString("topic_id")
strFont = "<font face=""Tahoma, Arial, Helvetica"" size=2"
%>
<!--#include file="data/webGab_array.inc"-->
<%
' 0 id
' 1 subject
' 2 author
' 3 date
' 4 parent id
' 5 depth in tree
' 6 number of children
' 7 new replies?

' lists expanded folders by index number

if Request.QueryString("expand") <> "" then
	strExpand = Request.QueryString("expand")
else
	if Request.Form("expandall") = "Expand all" then
		for x = 0 to UBound(arMsgs,1)
			if arMsgs(x,6) > 0 then
				strExpand = strExpand & "(" & x & ")"
			end if
		next
	else
		strExpand = ""	
	end if
end if

' now cycle through every post

intRow = 1
for x = 0 to UBound(arMsgs,1)
	dim strColor

' thisx keeps track of current item even after x is
' possibly incremented to skip children
	
	thisx = x
	if intRow mod 2 <> 0 then
		strColor = arColor(2)
	else
		strColor = arColor(1)
	end if
	
	response.write "<tr><td bgcolor=""" & strColor _
		& """><font face=""Tahoma, Arial, Helvetica"" size=2>"

' insert a blank space for every unit of tree depth		

	for d = 1 to (arMsgs(x,5))
		response.write "<img src=""./images/blank_" _
			& strStyle & ".gif"">"
	next

' if there are children then display plus/minus graphic

	if arMsgs(x,6) > 0 then
		response.write "<a href=""webGab_list.asp?topic_id=" _
			& intTopic & "&expand="
 		if InStr(strExpand, "(" & x & ")") then
			
' if the current item is selected for expansion then
' display option to collapse
			
			response.write Replace(strExpand, "(" & x & ")", "") _
				& """><img src=""./images/minus_"
		else
			
' otherwise show option to expand and skip all children

			response.write strExpand & "(" & x & ")" _
				& """><img src=""./images/plus_"
			x = x + arMsgs(x,6)	
		end if
		response.write strStyle & ".gif"" border=0></a>"
	else
		response.write "<img src=""./images/blank_" _
			& strStyle & ".gif"">"
	end if
	
	response.write arMsgs(thisx,1) & "</font></td>" & VbCrLf _
		& "<td bgcolor=""#" & strColor & """ align=""center"">" _
		& strFont & ">" & VbCrLf & "<nobr>" _
		& arMsgs(thisx,2) & "</nobr></font></td>" & VbCrLf _
		& "<td bgcolor=""#" & strColor & """>" _
		& strFont & ">" & VbCrLf & "<nobr>" _
		& arMsgs(thisx,3) & "</nobr></font></td>" & VbCrLf _
		& "<td bgcolor=""#" & strColor & """ align=""right"">" _
		& strFont

	if arMsgs(thisx,7) then
		response.write " color=""#" & arColor(10) & """><b>"
	else
		response.write ">"
	end if
	
	if arMsgs(thisx,6) > 0 then
		response.write arMsgs(thisx,6)
	else
		response.write "&nbsp;"
	end if
	
	response.write "</font></td>" & VbCrLf & VbCrLf
	intRow = intRow + 1
next
%>

<tr>
	<td width="60%"></td>
	<td></td>
	<td></td>
	<td></td>

</table>
</font>
</body>
</html>