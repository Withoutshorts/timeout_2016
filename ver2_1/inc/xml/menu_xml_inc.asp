
<% 

Dim objXMLHTTP_menu, objXMLDOM_menu, i_menu, strHTML_menu

Set objXMLDom = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/menu_sprog.xml", False
'objXmlHttp.open "GET", "http://localhost/inc/xml/menu_sprog.xml", False
'objXmlHttp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_1/inc/xml/menu_sprog.xml", False
objXmlHttp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/menu_sprog.xml", False
'objXmlHttp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/menu_sprog.xml", False
'objXmlHttp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/menu_sprog.xml", False
'objXmlHttp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver4_22/inc/xml/menu_sprog.xml", False

objXmlHttp.send


Set objXmlDom = objXmlHttp.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp = Nothing



Dim Address_menu, Latitude_menu, Longitude_menu
Dim oNode_menu, oNodes_menu
Dim sXPathQuery_menu


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
sXPathQuery_menu = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_menu = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_menu = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_menu = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_menu = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_menu = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_menu = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_menu = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030

'Response.Write "Session.LCID" &  Session.LCID

Set oNode_menu = objXmlDom.documentElement.selectSingleNode(sXPathQuery_menu)
Address_menu = oNode_menu.Text

Set oNodes_menu = objXmlDom.documentElement.selectNodes(sXPathQuery_menu)

For Each oNode_menu in oNodes_menu

     menu_txt_001 = oNode_menu.selectSingleNode("txt_1").Text
     menu_txt_002 = oNode_menu.selectSingleNode("txt_2").Text
     menu_txt_003 = oNode_menu.selectSingleNode("txt_3").Text
     menu_txt_003 = oNode_menu.selectSingleNode("txt_3").Text
    menu_txt_004 = oNode_menu.selectSingleNode("txt_4").Text
    menu_txt_005 = oNode_menu.selectSingleNode("txt_5").Text

    menu_txt_006 = oNode_menu.selectSingleNode("txt_6").Text
    menu_txt_007 = oNode_menu.selectSingleNode("txt_7").Text
    menu_txt_008 = oNode_menu.selectSingleNode("txt_8").Text
    menu_txt_009 = oNode_menu.selectSingleNode("txt_9").Text
    menu_txt_010 = oNode_menu.selectSingleNode("txt_10").Text
    menu_txt_011 = oNode_menu.selectSingleNode("txt_11").Text
    
next




%>