
<html>

<head>

<script language='VBScript'>

Dim objXMLHTTP

Sub cmdListProperties_OnClick()

  strMailboxItemURL = document.all.mailboxItemURL.Value
  strNameSpace = document.all.nameSpace.Value
  strPropertyName = document.all.propertyName.Value

  If strMailboxItemURL = "" Then
    MsgBox "Mailbox Item URL can't be empty."
    Exit Sub
  End If

  If UCase(Left(strMailboxItemURL, 4)) <> "HTTP" Then
    MsgBox "Mailbox Item URL must begin with http:// or https://"
    Exit Sub
  End If

  document.all.XMLStatus.Value = "Please Wait ..."
  document.all.XMLResponse.Value = ""

  Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
  objXMLHTTP.Open "PROPFIND", strMailboxItemURL, True
  objXMLHTTP.setRequestHeader "Content-type:", "text/xml"
  objXMLHTTP.setRequestHeader "Depth", "1"
  objXMLHTTP.onReadyStateChange = getRef("checkXMLHTTPState")

  If strNameSpace <> "" And strPropertyName <> "" Then
    strXML = "<a:propfind xmlns:a='DAV:'" & _
     " xmlns:x='" & strNameSpace & "'" & _
     ">" & _
     "<a:prop><x:" & strPropertyName & "/></a:prop></a:propfind>"
  End If

  objXMLHTTP.Send(strXML)

End Sub

Sub checkXMLHTTPState

  If (objXMLHTTP.readyState <> 4) Then Exit Sub
  document.all.XMLResponse.value = objXMLHTTP.ResponseText
  document.all.XMLStatus.value = objXMLHTTP.Status & " - " & objXMLHTTP.StatusText
  Set objXMLHTTP = Nothing

End Sub

</script>

</head>

<body>

<font face='Verdana' size='2'>

<p>
<b>Mailbox Item URL:</b><br>
<input type='text' name='mailboxItemURL' size='75' value='http://servername/exchange/username/inbox/example subject.eml'>

<p>
<b>Property Namespace:</b><br>
<input type='text' name='nameSpace' size='75' value='http://schemas.microsoft.com/mapi/proptag/'><br>
<b>Property Name:</b><br>
<input type='text' name='propertyName' size='75' value='x007d001e'><br>
<b>Leave either Namespace or Property empty to see all properties.</b>

<p>
<input type='button' name='cmdListProperties' value='    OK    '>

<p>
<table border='1' cellpadding='10' cellspacing='0' bgcolor='#E0E0E0'>
  <tr>
    <td><font size='2'>
      <b>Status:</b>
      <input name='XMLStatus' readonly>
    </td>
  </tr>
  <tr>
    <td><font size='2'>
      <b>Response</b><br>
      <textarea name='XMLResponse' rows='15' cols='75' readonly></textarea>
    </td>
  </tr>
</table>

</body>

</html>
