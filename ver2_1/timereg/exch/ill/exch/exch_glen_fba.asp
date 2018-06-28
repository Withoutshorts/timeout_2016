<%
Set req = server.CreateObject("Microsoft.XMLhttp")
servername = "mail.kringitsolutions.dk"
mailbox = "timeout"
domain = "KRING"
strpassword = "bgt567ujmko0"
strusername =  domain & "\" & mailbox
szXml = "destination=https://" & servername & "/exchange&flags=0&username=" & strusername 
szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"
req.Open "post", "https://" & servername & "/exchweb/bin/auth/owaauth.dll", False
req.send szXml
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for i = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(i)),"set-cookie: sessionid=") then reqsessionID = right(reqhedrarry(i),len(reqhedrarry(i))-12)
	if instr(lcase(reqhedrarry(i)),"set-cookie: cadata=") then reqcadata= right(reqhedrarry(i),len(reqhedrarry(i))-12)
next

stmsgbody ="Hello Fred <BR<BR>"
stmsgbody = stmsgbody & "What to get some coffee Mate, Meet me downstairs in ten<BR><BR>"
stmsgbody = stmsgbody & "Cheers<BR>"
stmsgbody = stmsgbody & "Barney<BR>"

szXml = ""
szXml = szXml & "Cmd=send" & vbLf
szXml = szXml & "MsgTo=timeout@kringitsolutions.dk" & vbLf
szXml = szXml & "MsgCc=" & vbLf
szXml = szXml & "MsgBcc=" & vbLf
szXml = szXml & "urn:schemas:httpmail:importance=1" & vbLf
szXml = szXml & "http://schemas.microsoft.com/exchange/readreceiptrequested=1" & vbLf
szXml = szXml & "http://schemas.microsoft.com/exchange/deliveryreportrequested=1" & vbLf
szXml = szXml & "http://schemas.microsoft.com/exchange/sensitivity-long=" & vbLf
szXml = szXml & "urn:schemas:httpmail:subject=Coffee ???" & vbLf
szXml = szXml & "urn:schemas:httpmail:htmldescription=<!DOCTYPE HTML PUBLIC " _
& """-//W3C//DTD HTML 4.0 Transitional//EN""><HTML DIR=ltr><HEAD><META HTTP-EQUIV" _
& "=""Content-Type"" CONTENT=""text/html; charset=utf-8""></HEAD><BODY><DIV>" _
& "<FONT face='Arial' color=#000000 size=2>" & stmsgbody & "</font>" _
& "</DIV></BODY></HTML>" & vbLf


req.Open "POST", "https://" & servername & "/exchange/" & mailbox & "/", False, "", ""
req.setRequestHeader "Accept-Language:", "en-us"
req.setRequestHeader "Content-type:", "application/x-www-UTF8-encoded"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Content-Length:", Len(szXml)
req.Send szXml
'Wscript.echo req.responseText 
Response.write req.responseText 
Response.write "her"

Set req = Nothing

Response.write "<hr>"
Response.flush




'*** Appointment **************


tidszoneTilpasningStart = dateadd("h", -2, "2005-11-17 13:00:00") 
tidszoneTilpasningSlut = dateadd("h", -2, "2005-11-17 15:00:00") 

tzSt = tidszoneTilpasningStart
tzSl = tidszoneTilpasningSlut


'*** Dage funktioner ***
if day(tzSt) < 10 then
tzStDag = "0"&day(tzSt)
else
tzStDag = day(tzSt)
end if

if day(tzSl) < 10 then
tzSlDag = "0"&day(tzSl)
else
tzSlDag = day(tzSl)
end if

'*** Måned funktioner ***
if month(tzSt) < 10 then
tzStMonth = "0"&month(tzSt)
else
tzStMonth = month(tzSt)
end if

if month(tzSl) < 10 then
tzSlMonth = "0"&month(tzSl)
else
tzSlMonth = month(tzSl)
end if

strTilPasStart = year(tzSt)&"-"&tzStMonth&"-"&tzStDag&"T"&right(tidszoneTilpasningStart, 8)&".000Z"
strTilPasSlut = year(tzSl)&"-"&tzSlMonth&"-"&tzSlDag&"T"&right(tidszoneTilpasningSlut, 8)&".000Z"


 ' Appointment item. ???????????
 strApptItem = "testappointment16.eml" '* Skal ændres hver gang ***

 
 
  '*** Nye Cookies ***
Set req = server.CreateObject("Microsoft.XMLhttp")
servername = "mail.kringitsolutions.dk"
mailbox = "timeout"
domain = "KRING"
strpassword = "bgt567ujmko0"
strusername =  domain & "\" & mailbox
 strApptRequest = "destination=https://" & servername & "/exchange&flags=0&username=" & strusername 
 strApptRequest =  strApptRequest & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"
req.Open "post", "https://" & servername & "/exchweb/bin/auth/owaauth.dll", False
req.send strApptRequest
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for i = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(i)),"set-cookie: sessionid=") then reqsessionID = right(reqhedrarry(i),len(reqhedrarry(i))-12)
	if instr(lcase(reqhedrarry(i)),"set-cookie: cadata=") then reqcadata= right(reqhedrarry(i),len(reqhedrarry(i))-12)
next
 
strApptRequest = ""

' XML namespace info of the WebDAV request.
 strXMLNSInfo = "xmlns:g=""DAV:"" " & _
    "xmlns:e=""http://schemas.microsoft.com/exchange/"" " & _
    "xmlns:mapi=""http://schemas.microsoft.com/mapi/"" " & _
    "xmlns:mapit=""http://schemas.microsoft.com/mapi/proptag/"" " & _
    "xmlns:x=""xml:"" xmlns:cal=""urn:schemas:calendar:"" " & _
    "xmlns:dt=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"" " & _
	"xmlns:header=""urn:schemas:mailheader:"" " & _
    "xmlns:mail=""urn:schemas:httpmail:"">"

 ' Set the appointment item properties.  The reminder time is set in seconds.
 ' To create an all-day meeting, set the dtstart/dtend range for 24 hours
 ' or more and set the alldayevent property to 1.  See the documentation
 ' on the properties in the urn:schemas:calendar: namespace for more information.
 strCalInfo = "<cal:location>Møde med OutZourCE hos KringIT</cal:location>" & _
    "<cal:dtstart dt:dt=""dateTime.tz"">"& strTilPasStart &"</cal:dtstart>" & _
    "<cal:dtend dt:dt=""dateTime.tz"">"& strTilPasSlut &"</cal:dtend>" & _
    "<cal:instancetype dt:dt=""int"">0</cal:instancetype>" & _
    "<cal:busystatus>BUSY</cal:busystatus>" & _
    "<cal:meetingstatus>CONFIRMED</cal:meetingstatus>" & _
    "<cal:alldayevent dt:dt=""boolean"">0</cal:alldayevent>" & _
	"<cal:responserequested dt:dt=""boolean"">1</cal:responserequested>" & _
    "<cal:reminderoffset dt:dt=""int"">900</cal:reminderoffset>"& _
	"<cal:timezoneid dt:dt=""int"">15</cal:timezoneid>"  & _
	"<mapi:0x00008223 dt:dt=""boolean"">1</mapi:0x00008223>"   'Recorrence

 ' Set the required attendee of the appointment.
  strHeaderInfo = "<header:to>" & strMailbox & "</header:to>"

 ' Set the subject of the appointment.
 strMailInfo = "<mail:subject>13.00</mail:subject>" & _
    "<mail:htmldescription>Yderligere info:</mail:htmldescription>"

 ' Build the XML body of the PROPPATCH request.
 strApptRequest = "<?xml version=""1.0""?>" & _
 "<g:propertyupdate " & strXMLNSInfo & _
  "<g:set><g:prop>" & _
    "<g:contentclass>urn:content-classes:appointment</g:contentclass>" & _
    "<e:outlookmessageclass>IPM.Appointment</e:outlookmessageclass>" & _
    strMailInfo & _
    strCalInfo & _
    strHeaderInfo & _
    "<mapi:finvited dt:dt=""boolean"">1</mapi:finvited>" & _
   "</g:prop></g:set>" & _
 "</g:propertyupdate>"

 
 

req.Open "POST", "https://" & servername & "/exchange/" & mailbox & "/", False, "", ""
req.setRequestHeader "Accept-Language:", "en-us"
req.setRequestHeader "Content-type:", "application/x-www-UTF8-encoded"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Content-Length:", Len(strApptRequest)
req.Send strApptRequest

 ' The PROPPATCH request was successful.
 If (req.Status >= 200 And req.Status < 300) Then
  Response.write "Appointment created successfully."

 ' Display error info.
 Else

  Response.write "PROPPATCH status: " & req.Status & vbCrLf & _
    "Status text: " & req.statusText
 End If

' Clean up.
Set req = Nothing

%>