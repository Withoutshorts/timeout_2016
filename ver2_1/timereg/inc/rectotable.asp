 <%
 Function RecToTable(objRec)
 
 Dim strT
 Dim fldF
 
 'table header
 strT = "<table border=0 cellpadding=3 cellspacing=2 width=600>" & "<tr align=center>"
 
 strT = strT & "<td width=55>"
 strT = strT & "<b>Dato</b>"
 strT = strT & "</td>"
 
 strT = strT & "<td width=55>"
 strT = strT & "<b>Taste dato</b>"
 strT = strT & "</td>"
 
 strT = strT & "<td>"
 strT = strT & "<b>Jobnr</b>"
 strT = strT & "</td>"
 
 strT = strT & "<td width=40>"
 strT = strT & "<b>Jobnavn</b>"
 strT = strT & "</td>"
 
 strT = strT & "<td width=40>"
 strT = strT & "<b>Aktivitet</b>"
 strT = strT & "</td>"
 
 strT = strT & "<td width=50>"
 strT = strT & "<b>Kunde</b>"
 strT = strT & "</td>"
 
 strT = strT & "<td>"
 strT = strT & "<b>Timer</b>"
 strT = strT & "</td>"
 
 strT = strT & "<td>"
 strT = strT & "<b>Id</b>"
 strT = strT & "</td>"
 
 strT = strT & "<td width=55>"
 strT = strT & "<b>Fakbar?</b>"
 strT = strT & "</td>"

strF = strF & "</TR>"
 
 
 'now build the rows
 intFields = objRec.fields.count - 1
   
 
 'Response.write intRows 
 
 
		 While not objRec.EOF
		 
		 strT = strT & "<TR Align=center>"
		 
		 For each fldF in objRec.Fields
		 	strT = strT & "<td>" & fldF.value & "</td>"
		 Next	
		 
		 strT = strT & "<td>" & "<a href='rediger_tastede_dage.asp?id="& objRec("Tid").value &"'>rediger</a>" & "</td>"
		 
		 strT = strT & "</tr>"
 
	
 		objRec.Movenext
 		Wend
	 
 strT = strT & "</table>"
 
 RecToTable = strT
 
 End function
 
 		
  %>
