<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/12/1999

' this include assumes that intTopic has already been
' defined and populated

' To prevent conflicting Session variables with multiple
' copies of this program, specify a unique data name for
' this copy:

dataName = "newGab"

' if the menu array doesn't already exist then
' create it

if Not IsArray(Application(dataName & intTopic & "Msgs")) then
'if 1 = 1 then
	dim DSN, strQuery, rsMsgs, intCount, x, intDepth, intRow, strUsedRow

	DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; " _
		& "DBQ=" & Server.Mappath("data/" & dataName & ".mdb")

	Set rsMsgs = Server.CreateObject("ADODB.RecordSet")

	strQuery = "SELECT * FROM gab_items WHERE msg_topic=" _
		& intTopic & " ORDER BY msg_date DESC"

' adOpenStatic = 3
' adLockReadOnly = 1
' adCmdText = &H0001

	rsMsgs.Open strQuery, DSN, 3, 1, &H0001

	intCount = CInt(rsMsgs.Recordcount - 1)
	Redim arTemp(intCount,4)
	Redim arMsgs(intCount,7)

' go through each record and put the values
' into a temporary two dimensional array

	for x = 0 to intCount
		arTemp(x,0) = rsMsgs("msg_id")

' if the post has no message then add an asterisk and use
' different link for the subject
			
		if rsMsgs("msg_body") <> "" then
			arTemp(x,1) = "<a href=""webGab_detail.asp?msg_id=" _
				& rsMsgs("msg_id") & """ target=""detail"">" _
				& rsMsgs("msg_subject") & "</a>"
		else
			arTemp(x,1) = "<a href=""webGab_post.asp?msg_id=" _
				& rsMsgs("msg_id") & """ target=""detail"">" _
				& rsMsgs("msg_subject") & "</a>*"
		end if
		
		arTemp(x,2) = rsMsgs("msg_author")

		arTemp(x,3) = Left(rsMsgs("msg_date"), Len(rsMsgs("msg_date"))-6) _
			& Right(rsMsgs("msg_date"), 3)

' if the message was posted in the last 24 hours then highlight
' the time			

		if DateDiff("h", rsMsgs("msg_date"), Now) < 24 then
			arTemp(x,3) = "<font color=""#" & arColor(10) & """>" _
				& arTemp(x,3) & "</font>"
		end if

		arTemp(x,4) = rsMsgs("msg_parent")

		rsMsgs.MoveNext
	next
	rsMsgs.Close
	Set rsMsgs = nothing
	
	intDepth = 0
	intRow = -1
	strUsedRow = ""
	
' with the temporary array created, now go
' through it to order the items and calculate
' the items' depth in the tree and number
' of children, producing these fields:
' 0 id
' 1 subject
' 2 author
' 3 date
' 4 parent id
' 5 depth in tree (added)
' 6 number of replies (added)
' 7 new replies? (added)

	Sub orderItems(intParentID,intCallingRow)
	
' declare these values within the subroutine
' so that they remain strictly local and can
' be used recursively

		dim i, j, intParentRow, intOldRow
		intParentRow = intCallingRow
		for i = 0 to UBound(arTemp,1)
			if intParentID = arTemp(i,4) then
			
' intRow keeps global track of the current
' position in the menu
				
				intRow = intRow + 1

' intOldRow keeps local track of the current
' row even after the subroutine has been
' called recursively to display the children
' (which increments row and updates parentRow)

				intOldRow = intRow		
	
' transfer the contents of the temporary
' array to the new one

				for j = 0 to 4
					arMsgs(intRow,j) = arTemp(i,j)
				next

				arMsgs(intRow,5) = intDepth
				arMsgs(intRow,6) = 0

' here we call this same subroutine, recursively,
' adding to the tree depth each time and sub-
' tracting that depth when we return

				intDepth = intDepth + 1
				Call orderItems(arTemp(i,0),intRow)
				intDepth = intDepth - 1

' each item must indicate not only the number of
' immediate children but all progeny, so we add
' one to the immediate parent row with each run
' through and then we add the progeny of the current
' row to the parent row, so it has a total

				if intParentRow <> "" then
					arMsgs(intParentRow,6) = arMsgs(intParentRow,6) + arMsgs(intOldRow,6) + 1

' if the time has been highlighted with another font color
' then it must be a post within the last 24 hours, thus
' also highlight the parents' reply count
					
					if Left(arMsgs(intOldRow,3), 2) = "<f" OR arMsgs(intOldRow,7) then
						arMsgs(intParentRow,7) = True
					end if
				end if
			end if
		next
	end Sub

' we prime the subroutine by displaying all
' root items
	
	Call orderItems(0,"")
	Application(dataName & intTopic & "Msgs") = arMsgs
else

' if the Session array already exists then
' assign it to the arMsgs array

	arMsgs = Application(dataName & intTopic & "Msgs")
end if
%>