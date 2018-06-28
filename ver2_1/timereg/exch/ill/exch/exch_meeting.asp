<%response.buffer = true%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<%

Set objSession = server.createobject("MAPI.session")

strServer = "outzource-747"
strMailbox = "administrator@arusha.outzource.dk" '"Administrator"
strProfileInfo = strServer & vbLf & strMailbox


'profileName = "51935@demo.danexchange.dk"
'profilePW = "mjlnduqzw"
'strProfileInfo = ""


'*** Test ***
'objSession.logon profileName, profilePW, false, True, 0, True, strProfileInfo  
'*** Virker **
objSession.logon "", "", false, True, 0, True, strProfileInfo
	
	
	If objSession Is Nothing Then
        Response.write ("Must log on first")
    End If
    
	Set objAddrEntry = objSession.CurrentUser
   	If objAddrEntry Is Nothing Then
        Response.write "Could not set the address entry object"
    Else
        Response.write "Full address = " & objAddrEntry.Type & ":" _
                                 & objAddrEntry.Address
    End If


'Dim objSess 
'Dim objCalFold As Folder ' calendar folder to create appointments in
'Dim colAppts As Messages ' appointments collection in calendar folder
'Dim objNewAppt As AppointmentItem
'Dim colRecips As Recipients ' users to be invited to the meeting

Set objCalFold = objSession.GetDefaultFolder(0) 'CdoDefaultFolderCalendar
Set colAppts = objCalFold.Messages
Set objNewAppt = colAppts.Add ' no parameters when adding appointments

With objNewAppt
   .MeetingStatus = 1 'CdoMeeting ' required before calling .Send
   .Importance = 1 'CdoHigh
   .Location = "My office"
   .Subject = "Extremely important meeting tomorrow at 9:00 AM!"
   .Text = "We will discuss stock distribution."
   .StartTime = DateAdd("d", 1, DateAdd("h", 12, Date))
   .EndTime = DateAdd("h", 1, .StartTime) ' meeting to last one hour
   Set colRecips = .Recipients ' empty collection of users to invite
End With

With colRecips ' populate collection with invitees
   '.Add "Søren Karlsen", "SMTP:sk@outzource.dk" ' can't use parentheses!
   '.Add "SK - dandomain", "SMTP:51935@demo.danexchange.dk"
   .Add "TimeOut", "SMTP:timeout@kringitsolutions.dk"
   .Resolve ' default to dialog in case of name ambiguities
End With

objNewAppt.Send ' appointment becomes a meeting at this moment
%>


</body>
</html>
