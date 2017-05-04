
<% 
Dim objXMLHTTP_week, objXMLDOM_week, i_week, strHTML_week

Set objXMLDom_week = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_week = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_week.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/weeklynote_sprog.xml", False
objXmlHttp_week.open "GET", "http://localhost/inc/xml/weeklynote_sprog.xml", False
'objXmlHttp_week.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/weeklynote_sprog.xml", False
'objXmlHttp_week.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/weeklynote_sprog.xml", False
'objXmlHttp_week.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/weeklynote_sprog.xml", False
'objXmlHttp_week.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/weeklynote_sprog.xml", False

objXmlHttp_week.send


Set objXmlDom_week = objXmlHttp_week.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_week = Nothing



Dim Address_week, Latitude_week, Longitude_week
Dim oNode_week, oNodes_week
Dim sXPathQuery_week

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
sXPathQuery_week = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_week = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_week = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_week = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_week = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_week = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_week = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_week = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
'Session.LCID = 1030



Set oNode_week = objXmlDom_week.documentElement.selectSingleNode(sXPathQuery_week)
Address_week = oNode_week.Text

Set oNodes_week = objXmlDom_week.documentElement.selectNodes(sXPathQuery_week)

    For Each oNode_week in oNodes_week

          
        week_txt_001 = oNode_week.selectSingleNode("txt_1").Text
        week_txt_002 = oNode_week.selectSingleNode("txt_2").Text
        week_txt_003 = oNode_week.selectSingleNode("txt_3").Text
        week_txt_003 = oNode_week.selectSingleNode("txt_3").Text
        week_txt_004 = oNode_week.selectSingleNode("txt_4").Text
        week_txt_005 = oNode_week.selectSingleNode("txt_5").Text

        week_txt_006 = oNode_week.selectSingleNode("txt_6").Text
        week_txt_007 = oNode_week.selectSingleNode("txt_7").Text
        week_txt_008 = oNode_week.selectSingleNode("txt_8").Text
        week_txt_009 = oNode_week.selectSingleNode("txt_9").Text
        week_txt_010 = oNode_week.selectSingleNode("txt_10").Text
        week_txt_011 = oNode_week.selectSingleNode("txt_11").Text
          
    next




'Response.Write "week_txt_001: " & week_txt_001 & "<br>"
'Response.Write "week_txt_002: " & week_txt_002 & "<br>"


%>