<%
Dim objXMLHTTP, objXMLDOM, objModul, objDECS, i

'Opretter en instans af Microsoft.XMLHTTP, så det er muligt at få fat på dokumentet
Set xmlHTTP = Server.CreateObject("Microsoft.XMLHTTP")

'Opretter en instans af Microsofts XML-parser, XMLDOM
Set objXMLDOM = Server.CreateObject("Microsoft.XMLDOM")

'Opretter forbindelse til xml-dokumentet
'Call objXMLHTTP.Open("GET", "http://ktv/timeout_xp/inc/xml/conf4.xml", False)
'objXMLHTTP.Send
set xmlHTTP = CreateObject("Microsoft.XMLHTTP") 
xmlHttp.open "GET", "conf4.xml", false
xmlHttp.send
xmlDoc=xmlHttp.responseText
document.write "<xmp>" + xmlDoc + "</xmp>"

set xmldoc = Server.CreateObject("Microsoft.XMLDOM")
xmldoc.async=false
xmldoc.load(request)

for each x in xmldoc.documentElement.childNodes
   if x.NodeName = "modul" then name=x.text
next

response.write(name) 

'Lægger indholdet af dokumentet over i vores XML-parser objekt
'Set objXMLDOM = objXMLHTTP.ResponseXML

'Henter indholdet af alle tags med navnet 'modul'
Set objModul = objXMLDOM.getElementsByTagName("modulname")

'Henter indholdet af alle tags med navnet 'url'
Set objDESC = objXMLDOM.getElementsByTagName("desc")

'Henter indholdet af alle tags med navnet 'dsn'
Set objDSN = objXMLDOM.getElementsByTagName("dbconn")

Response.Write("<table width=""100%"" cellspacing=""5"" cellpadding=""5"">")

'Løber igennem alle tags, og udskriver dem i en tabel
For i = 0 To objModul.length - 1

  Response.Write("<tr>")
  Response.Write("<td width='200'>")
  Response.Write objModul(i).text &"&nbsp;&nbsp;" & objDESC(i).text 
  Response.Write "</td></tr>"
	
Next

Response.Write("</table>")

Response.write objDSN.text


Set objXMLHTTP = Nothing
Set objXMLDOM = Nothing
%>


