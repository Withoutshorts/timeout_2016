<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<script>

Sub CreateOtherUserAppointment()
    Dim objApp As Outlook.Application
    Dim objNS As Outlook.NameSpace
    Dim objFolder As Outlook.MAPIFolder
    Dim objDummy As Outlook.MailItem
    Dim objRecip As Outlook.Recipient
    Dim objAppt As Outlook.AppointmentItem
    Dim strMsg As String
    Dim strName As String
    On Error Resume Next
    
    //### name of person whose Calendar you want to use ### '
    strName = "FlaviusJ"
    
    Set objApp = CreateObject("Outlook.Application")
    Set objNS = objApp.GetNamespace("MAPI")
    Set objDummy = objApp.CreateItem(olMailItem)
    Set objRecip = objDummy.Recipients.Add(strName)
    objRecip.Resolve
    If objRecip.Resolved Then
        On Error Resume Next
        Set objFolder = _
          objNS.GetSharedDefaultFolder(objRecip, _
            olFolderCalendar)
        If Not objFolder Is Nothing Then
            Set objAppt = objFolder.Items.Add
            If Not objAppt Is Nothing Then
                With objAppt
                    .Subject = "Test Appointment"
                    .Start = Date + 14
                    .AllDayEvent = True
                    .Save
                End With
            End If
        End If
    Else
        MsgBox "Could not find " & Chr(34) & strName & Chr(34), , _
               "User not found"
    End If

    Set objApp = Nothing
    Set objNS = Nothing
    Set objFolder = Nothing
    Set objDummy = Nothing
    Set objRecip = Nothing
    Set objAppt = Nothing
End Sub


</script>








</body>
</html>
