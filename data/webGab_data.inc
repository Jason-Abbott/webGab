<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 08/06/1999

' To prevent conflicting Session variables with multiple
' copies of this program, specify a unique data name for
' this copy:

dataName = "newGab"

' This opens the database webNav.mdb stored in the
' the directory /data under the main webNav files

dim db
Set db = Server.CreateObject("ADODB.Connection")
DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; " _
	& "DBQ=" & Server.Mappath("data/" & dataName & ".mdb")
db.Open DSN

' make sure the containing file does
' db.Close
' Set db = nothing
%>