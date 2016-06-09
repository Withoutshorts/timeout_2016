<%response.buffer = true%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<%

 ' Username and password of appointment creator.
 strUserName = "arusha\administrator"
 strPassWord = "Sok!2637"
 
 strExchSvrName = "outzource-747" '"localhost/exchange" '"www.danexchange.dk" '"Exchangeservername"

 ' Mailbox folder name.
 strMailbox = "administrator" '"administrator@arusha.outzource.dk" '"user"

 ' Appointment item.
 strApptItem = "testappointment10.eml" '* Skal ændres hver gang ***

' URL of the appointment item.
 strApptURL = "http://" & strExchSvrName & "/exchange/" & _
 strMailbox & "/Calendar/" & strApptItem


uri = strApptURL
 

'// Initialize the XMLHTTP request object.
 Set xmlReq = server.CreateObject("Microsoft.XMLHTTP")
 'xmlReq.open "PROPPATCH", strApptURL, False, strUserName, strPassWord
 xmlReq.open "DELETE", uri, false, strUserName, strPassWord
 xmlReq.setRequestHeader "Content-Type", "text/xml"
 xmlReq.send()

'// An error occurred on the server.
' if xmlReq.status >= 500 then
 
'  Response.write "Status: " & xmlReq.status & "<br>"
'  Response.write "Status text: An error occurred on the server."
 

' else

'   Response.write "Status: " + xmlReq.status
'   Response.write "Status text: " + xmlReq.statustext
  
' end if


%>


</body>
</html>
