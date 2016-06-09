<%
Dim objXMLHTTP, objXMLDOM, objTitles, objURLs, objCategories, i

'Opretter en instans af Microsoft.XMLHTTP, så det er muligt at få fat på dokumentet
Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")

'Opretter en instans af Microsofts XML-parser, XMLDOM
Set objXMLDOM = Server.CreateObject("Microsoft.XMLDOM")

'Opretter forbindelse til xml-dokumentet
Call objXMLHTTP.Open("GET", "http://www.magik.dk/links.asp", False)

objXMLHTTP.Send

'Lægger indholdet af dokumentet over i vores XML-parser objekt
Set objXMLDOM = objXMLHTTP.ResponseXML

'Henter indholdet af alle tags med navnet 'titel'
Set objTitles = objXMLDOM.getElementsByTagName("titel")

'Henter indholdet af alle tags med navnet 'url'
Set objURLs = objXMLDOM.getElementsByTagName("url")

'Henter indholdet af alle tags med navnet 'kategori'
Set objCategories = objXMLDOM.getElementsByTagName("kategori")

Response.Write("<table width=""100%"" cellspacing=""5"" cellpadding=""5"">")

'Løber igennem alle tags, og udskriver dem i en tabel
For i = 0 To objTitles.length - 1

  Response.Write("<tr>")
  Response.Write("<td width='200'>")
  Response.Write("<a href='" & objURLs(i).text & "' target='_blank'>")
  Response.Write(objTitles(i).text & "</a></td>")
  Response.Write("<td>" & objCategories(i).text & "</td>")
  Response.Write("</tr>")
	
Next

Response.Write("</table>")

Set objXMLHTTP = Nothing
Set objXMLDOM = Nothing
%>


