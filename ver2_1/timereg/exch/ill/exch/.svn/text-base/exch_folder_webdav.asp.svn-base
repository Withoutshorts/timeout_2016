<%response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/header_hvd_inc.asp"-->


<%
svr =  "mail.kringitsolutions.dk" '"Servername"
mailbox = "Timeout" '"Mailbox"
NewFLd = "Spamfolder"
strURL = "https://" & svr & "/exchange/" & mailbox & "/inbox/"
strQuery = ""
strQuery = strQuery & "SELECT ""http://schemas.microsoft.com/mapi/proptag/x3001001E"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = True AND "
strQuery = strQuery & """http://schemas.microsoft.com/mapi/proptag/x3001001E"" = '" & NewFLd & "'"
set req = server.createobject("microsoft.xmlhttp")
req.open "SEARCH", strURL, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.send strQuery
	
	If req.status >= 500 Then
		ElseIf req.status = 207 Then
		set oResponseDoc = req.responseXML
		set oNodeList = oResponseDoc.getElementsByTagName("d:x3001001E")
		wscript.echo oNodeList.length 
		if oNodeList.length <> 0 then
		wscript.echo "Folder Already Exists"
		else
			call Createfolder(NewFld,strURL)
		end if 
	Else
		Response.write req.status
	End If

Sub Createfolder(fldname,parentfolder)

nfolderURL = parentfolder & "/" & fldname & "/"
req.open "MKCOL", nfolderURL, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.send
if req.status = 201 then
Wscript.echo "Folder created sucessfully"
else
wscript.echo req.status
wscript.echo req.statustext
end if

end sub

%>


<!--#include file="../inc/regular/footer_inc.asp"-->

