<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>TimeOut Email adr.</title>
</head>

<body>
<%
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRec = Server.CreateObject("ADODB.Recordset")


x = 1
numberoflicens = 14
For x = 1 To 14 'numberoflicens
Select Case x
Case 1
strConnect = "driver={MySQL};server=192.168.1.33;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;" 
Case 2
strConnect = "driver={MySQL};server=192.168.1.33;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_sthaus;"  
Case 3
strConnect = "driver={MySQL};server=192.168.1.33;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_buying;"
Case 4
strConnect = "driver={MySQL};server=192.168.1.33;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_inlead;"
Case 5
strConnect = "driver={MySQL};server=192.168.1.33;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_netstrategen;"
Case 6
strConnect = "driver={MySQL};server=192.168.1.33;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ravnit;"
Case 7
strConnect = "driver={MySQL};server=192.168.1.33;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_titoonic;"
Case 8
strConnect = "driver={MySQL};server=192.168.1.33;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_userneeds;"
Case 9
strConnect = "driver={MySQL};server=192.168.1.33;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_demo;"
case 10
strConnect = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mezzo;"
case 11
strConnect = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_skybrud;"
case 12
strConnect = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_gramtech;"
case 13
strConnect = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cybervision;"
case 14
strConnect = "driver={MySQL};server=192.168.1.33; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ferro;"


'Case 
'strConnect = "driver={MySQL};server=192.168.1.33;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_margin;"
'* Margin Media fejler
End Select

oConn.open strConnect
			
			strSQL = "SELECT licens FROM licens"
 			oRec.open strSQL, oConn, 3
            
			if Not oRec.EOF then
			Response.write "<br><br><b>" & oRec("licens")& "</b><hr>"
			end if
			oRec.close
			
 			strSQL = "SELECT mnavn, email FROM medarbejdere"
 			oRec.open strSQL, oConn, 3
            
			While Not oRec.EOF
				
				'Response.write oRec("mnavn") & ", " & oRec("email") & "<br>"
				Response.write oRec("email") & "; "
				
			oRec.movenext
			wend
			
			
			oRec.close
           oConn.close

 Next

 Set oRec = Nothing
 Set oConn = Nothing
%>


</body>
</html>
