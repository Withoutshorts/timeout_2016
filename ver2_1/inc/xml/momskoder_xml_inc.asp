
<% 
Dim objXMLHTTP_momskoder, objXMLDOM_momskoder, i_momskoder, strHTML_momskoder

Set objXMLDom_momskoder = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_momskoder = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_momskoder.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/momskoder_sprog.xml", False
'objXmlHttp_momskoder.open "GET", "http://localhost/inc/xml/momskoder_sprog.xml", False
'objXmlHttp_momskoder.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/momskoder_sprog.xml", False
'objXmlHttp_momskoder.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/momskoder_sprog.xml", False
'objXmlHttp_momskoder.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/momskoder_sprog.xml", False
'objXmlHttp_momskoder.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/momskoder_sprog.xml", False
objXmlHttp_momskoder.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/momskoder_sprog.xml", False

objXmlHttp_momskoder.send


Set objXmlDom_momskoder = objXmlHttp_momskoder.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_momskoder = Nothing



Dim Address_momskoder, Latitude_momskoder, Longitude_momskoder
Dim oNode_momskoder, oNodes_momskoder
Dim sXPathQuery_momskoder

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
sXPathQuery_momskoder = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_momskoder = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_momskoder = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_momskoder = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_momskoder = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_momskoder = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_momskoder = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_momskoder = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_momskoder = objXmlDom_momskoder.documentElement.selectSingleNode(sXPathQuery_momskoder)
Address_momskoder = oNode_momskoder.Text

Set oNodes_momskoder = objXmlDom_momskoder.documentElement.selectNodes(sXPathQuery_momskoder)

    For Each oNode_momskoder in oNodes_momskoder

        momskoder_txt_001 = oNode_momskoder.selectSingleNode("txt_1").Text
        momskoder_txt_002 = oNode_momskoder.selectSingleNode("txt_2").Text
        momskoder_txt_003 = oNode_momskoder.selectSingleNode("txt_3").Text
        momskoder_txt_004 = oNode_momskoder.selectSingleNode("txt_4").Text
        momskoder_txt_005 = oNode_momskoder.selectSingleNode("txt_5").Text
        momskoder_txt_006 = oNode_momskoder.selectSingleNode("txt_6").Text
        momskoder_txt_007 = oNode_momskoder.selectSingleNode("txt_7").Text
        momskoder_txt_008 = oNode_momskoder.selectSingleNode("txt_8").Text
        momskoder_txt_009 = oNode_momskoder.selectSingleNode("txt_9").Text
        momskoder_txt_010 = oNode_momskoder.selectSingleNode("txt_10").Text
        momskoder_txt_011 = oNode_momskoder.selectSingleNode("txt_11").Text
  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>