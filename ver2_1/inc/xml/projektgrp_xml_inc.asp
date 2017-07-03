
<% 
Dim objXMLHTTP_progrp, objXMLDOM_progrp, i_progrp, strHTML_progrp

Set objXMLDom_progrp = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_progrp = Server.CreateObject("Msxml2.ServerXMLHTTP")
objXmlHttp_progrp.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/projektgrp_sprog.xml", False
'objXmlHttp_progrp.open "GET", "http://localhost/inc/xml/progrp_sprog.xml", False
'objXmlHttp_progrp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/projektgrp_sprog.xml", False
'objXmlHttp_progrp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/projektgrp_sprog.xml", False
'objXmlHttp_progrp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/projektgrp_sprog.xml", False
'objXmlHttp_progrp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/projektgrp_sprog.xml", False

objXmlHttp_progrp.send


Set objXmlDom_progrp = objXmlHttp_progrp.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_progrp = Nothing



Dim Address_progrp, Latitude_progrp, Longitude_progrp
Dim oNode_progrp, oNodes_progrp
Dim sXPathQuery_progrp

sprog = 1 'DK
if len(trim(session("mid"))) <> 0 then
strSQL = "SELECT sprog FROM medarbejdere WHERE mid = " & session("mid")

oRec.open strSQL, oConn, 3
if not oRec.EOF then
sprog = oRec("sprog")
end if
oRec.close
end if

select case sprog
case 1
sXPathQuery_progrp = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_progrp = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_progrp = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_progrp = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_progrp = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_progrp = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_progrp = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_progrp = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_progrp = objXmlDom_progrp.documentElement.selectSingleNode(sXPathQuery_progrp)
Address_progrp = oNode_progrp.Text

Set oNodes_progrp = objXmlDom_progrp.documentElement.selectNodes(sXPathQuery_progrp)

    For Each oNode_progrp in oNodes_progrp

        'progrp_txt_001 = oNode_progrp.selectSingleNode("txt_1").Text

  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>