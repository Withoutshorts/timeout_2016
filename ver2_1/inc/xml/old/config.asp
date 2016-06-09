
<%
Dim objXMLHTTP, objXMLDOM, i

'Opretter en instans af Microsoft.XMLHTTP, så det er muligt at få fat på dokumentet
Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")

'Opretter en instans af Microsofts XML-parser, XMLDOM
Set objXMLDOM = Server.CreateObject("Microsoft.XMLDOM")

'Opretter forbindelse til xml-dokumentet
Call objXMLHTTP.Open("GET", "http://www.outzource.dk/outz/conf_outz.xml", False)

objXMLHTTP.Send

'Lægger indholdet af dokumentet over i vores XML-parser objekt
Set objXMLDOM = objXMLHTTP.ResponseXML

'Henter indholdet af alle tags med navnet 'modul'
Set objModuler = objXMLDOM.getElementsByTagName("modul")

'Henter indholdet af alle tags med navnet 'onoff'
'Set objOnoff = objXMLDOM.getElementsByTagName("onoff")

'Henter indholdet af alle tags med navnet 'url'
Set objDBConn = objXMLDOM.getElementsByTagName("dbconn")

'Henter indholdet af alle tags med navnet 'kategori'
'Set objCategories = objXMLDOM.getElementsByTagName("kategori")

Response.Write("<table width=""100%"" cellspacing=""5"" cellpadding=""5"">")

'Løber igennem alle tags, og udskriver dem i en tabel
For i = 0 To objModuler.length - 1

  Response.Write("<tr>")
  Response.Write("<td width='200'>")
  Response.Write(objModuler(i).text &"</td>")
  Response.Write("</tr>")

Next

  Response.Write("<tr>")
  Response.Write("<td width='200'>")
  Response.Write(objDBConn(0).text & "</a></td>")
  Response.Write("</tr>")

Response.Write("</table>")



Set objXMLHTTP = Nothing
Set objXMLDOM = Nothing
%>



<!--script type="text/vbscript">
set xmlDoc=CreateObject("Microsoft.XMLDOM")
xmlDoc.async = "false"
xmlDoc.load("conf4.xml")

//set x=xmlDoc.getElementsByTagName("modul")
//for i = 1 to x.length
//document.write(x.item(i-1).text)
//document.write("<br>")
//next

set dsn=xmlDoc.getElementsByTagName("dbconn")
document.write(dsn.item(0).text)
</script-->
