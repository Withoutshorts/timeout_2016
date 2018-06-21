
<% 
Dim objXMLHTTP_godkend, objXMLDOM_godkend, i_godkend, strHTML_godkend

Set objXMLDom_godkend = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_godkend = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_godkend.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/godkend_sprog.xml", False
'objXmlHttp_godkend.open "GET", "http://localhost/inc/xml/godkend_sprog.xml", False
'objXmlHttp_godkend.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/godkend_sprog.xml", False
'objXmlHttp_godkend.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/godkend_sprog.xml", False
'objXmlHttp_godkend.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/godkend_sprog.xml", False
objXmlHttp_godkend.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/godkend_sprog.xml", False

objXmlHttp_godkend.send


Set objXmlDom_godkend = objXmlHttp_godkend.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_godkend = Nothing



Dim Address_godkend, Latitude_godkend, Longitude_godkend
Dim oNode_godkend, oNodes_godkend
Dim sXPathQuery_godkend

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
sXPathQuery_godkend = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_godkend = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_godkend = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_godkend = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_godkend = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_godkend = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_godkend = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_godkend = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_godkend = objXmlDom_godkend.documentElement.selectSingleNode(sXPathQuery_godkend)
Address_godkend = oNode_godkend.Text

Set oNodes_godkend = objXmlDom_godkend.documentElement.selectNodes(sXPathQuery_godkend)

    For Each oNode_godkend in oNodes_godkend

        godkend_txt_001 = oNode_godkend.selectSingleNode("txt_1").Text
        godkend_txt_002 = oNode_godkend.selectSingleNode("txt_2").Text
        godkend_txt_003 = oNode_godkend.selectSingleNode("txt_3").Text
        godkend_txt_003 = oNode_godkend.selectSingleNode("txt_3").Text
        godkend_txt_004 = oNode_godkend.selectSingleNode("txt_4").Text
        godkend_txt_005 = oNode_godkend.selectSingleNode("txt_5").Text

        godkend_txt_006 = oNode_godkend.selectSingleNode("txt_6").Text
        godkend_txt_007 = oNode_godkend.selectSingleNode("txt_7").Text
        godkend_txt_008 = oNode_godkend.selectSingleNode("txt_8").Text
        godkend_txt_009 = oNode_godkend.selectSingleNode("txt_9").Text
        godkend_txt_010 = oNode_godkend.selectSingleNode("txt_10").Text

        godkend_txt_011 = oNode_godkend.selectSingleNode("txt_11").Text
        godkend_txt_012 = oNode_godkend.selectSingleNode("txt_12").Text
        godkend_txt_013 = oNode_godkend.selectSingleNode("txt_13").Text
        godkend_txt_014 = oNode_godkend.selectSingleNode("txt_14").Text
        godkend_txt_015 = oNode_godkend.selectSingleNode("txt_15").Text
        godkend_txt_016 = oNode_godkend.selectSingleNode("txt_16").Text
        godkend_txt_017 = oNode_godkend.selectSingleNode("txt_17").Text
        godkend_txt_018 = oNode_godkend.selectSingleNode("txt_18").Text
        godkend_txt_019 = oNode_godkend.selectSingleNode("txt_19").Text
        godkend_txt_020 = oNode_godkend.selectSingleNode("txt_20").Text
        godkend_txt_021 = oNode_godkend.selectSingleNode("txt_21").Text
        godkend_txt_022 = oNode_godkend.selectSingleNode("txt_22").Text
        godkend_txt_023 = oNode_godkend.selectSingleNode("txt_23").Text

                
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>