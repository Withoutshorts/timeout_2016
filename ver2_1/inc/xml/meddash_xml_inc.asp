
<% 

Dim objXMLHTTP_meddash, objXMLDOM_meddash, i_meddash, strHTML_meddash

Set objXMLDom = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp.open "GET", "http://localhost/inc/xml/meddash_sprog.xml", False
'objXmlHttp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_1/inc/xml/meddash_sprog.xml", False
objXmlHttp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver3_99/inc/xml/meddash_sprog.xml", False
'objXmlHttp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/meddash_sprog.xml", False
'objXmlHttp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver4_22/inc/xml/meddash_sprog.xml", False

objXmlHttp.send


Set objXmlDom = objXmlHttp.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp = Nothing



Dim Address_meddash, Latitude_meddash, Longitude_meddash
Dim oNode_meddash, oNodes_meddash
Dim sXPathQuery_meddash


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
sXPathQuery_meddash = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_meddash = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_meddash = "//sprog/se"
Session.LCID = 1053
case else
sXPathQuery_meddash = "//sprog/dk"
Session.LCID = 1030
end select

'Response.Write "Session.LCID" &  Session.LCID

Set oNode_meddash = objXmlDom.documentElement.selectSingleNode(sXPathQuery_meddash)
Address_meddash = oNode_meddash.Text

Set oNodes_meddash = objXmlDom.documentElement.selectNodes(sXPathQuery_meddash)

For Each oNode_meddash in oNodes_meddash

     meddash_txt_001 = oNode_meddash.selectSingleNode("txt_1").Text
     meddash_txt_002 = oNode_meddash.selectSingleNode("txt_2").Text
     meddash_txt_003 = oNode_meddash.selectSingleNode("txt_3").Text
     meddash_txt_003 = oNode_meddash.selectSingleNode("txt_3").Text
     meddash_txt_004 = oNode_meddash.selectSingleNode("txt_4").Text
     meddash_txt_005 = oNode_meddash.selectSingleNode("txt_5").Text

    meddash_txt_006 = oNode_meddash.selectSingleNode("txt_6").Text
    meddash_txt_007 = oNode_meddash.selectSingleNode("txt_7").Text
    meddash_txt_008 = oNode_meddash.selectSingleNode("txt_8").Text
    meddash_txt_009 = oNode_meddash.selectSingleNode("txt_9").Text

    
next




%>