
<% 
Dim objXMLHTTP_medarb_protid, objXMLDOM_medarb_protid, i_medarb_protid, strHTML_medarb_protid

Set objXMLDom_medarb_protid = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_medarb_protid = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_medarb_protid.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/medarb_protid_sprog.xml", False
'objXmlHttp_medarb_protid.open "GET", "http://localhost/inc/xml/medarb_protid_sprog.xml", False
'objXmlHttp_medarb_protid.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/medarb_protid_sprog.xml", False
'objXmlHttp_medarb_protid.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/medarb_protid_sprog.xml", False
'objXmlHttp_medarb_protid.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/medarb_protid_sprog.xml", False
objXmlHttp_medarb_protid.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/medarb_protid_sprog.xml", False
'objXmlHttp_medarb_protid.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/medarb_protid_sprog.xml", False

objXmlHttp_medarb_protid.send 


Set objXmlDom_medarb_protid = objXmlHttp_medarb_protid.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_medarb_protid = Nothing



Dim Address_medarb_protid, Latitude_medarb_protid, Longitude_medarb_protid
Dim oNode_medarb_protid, oNodes_medarb_protid
Dim sXPathQuery_medarb_protid

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
sXPathQuery_medarb_protid = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_medarb_protid = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_medarb_protid = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_medarb_protid = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_medarb_protid = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_medarb_protid = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_medarb_protid = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_medarb_protid = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_medarb_protid = objXmlDom_medarb_protid.documentElement.selectSingleNode(sXPathQuery_medarb_protid)
Address_medarb_protid = oNode_medarb_protid.Text

Set oNodes_medarb_protid = objXmlDom_medarb_protid.documentElement.selectNodes(sXPathQuery_medarb_protid)

    For Each oNode_medarb_protid in oNodes_medarb_protid

        medarb_protid_txt_001 = oNode_medarb_protid.selectSingleNode("txt_1").Text
        medarb_protid_txt_002 = oNode_medarb_protid.selectSingleNode("txt_2").Text
        medarb_protid_txt_003 = oNode_medarb_protid.selectSingleNode("txt_3").Text
        medarb_protid_txt_004 = oNode_medarb_protid.selectSingleNode("txt_4").Text
        medarb_protid_txt_005 = oNode_medarb_protid.selectSingleNode("txt_5").Text
        medarb_protid_txt_006 = oNode_medarb_protid.selectSingleNode("txt_6").Text
        medarb_protid_txt_007 = oNode_medarb_protid.selectSingleNode("txt_7").Text
        medarb_protid_txt_008 = oNode_medarb_protid.selectSingleNode("txt_8").Text
        medarb_protid_txt_009 = oNode_medarb_protid.selectSingleNode("txt_9").Text
        medarb_protid_txt_010 = oNode_medarb_protid.selectSingleNode("txt_10").Text

        medarb_protid_txt_011 = oNode_medarb_protid.selectSingleNode("txt_11").Text
        medarb_protid_txt_012 = oNode_medarb_protid.selectSingleNode("txt_12").Text
        medarb_protid_txt_013 = oNode_medarb_protid.selectSingleNode("txt_13").Text
        medarb_protid_txt_014 = oNode_medarb_protid.selectSingleNode("txt_14").Text
        medarb_protid_txt_015 = oNode_medarb_protid.selectSingleNode("txt_15").Text
        medarb_protid_txt_016 = oNode_medarb_protid.selectSingleNode("txt_16").Text
        medarb_protid_txt_017 = oNode_medarb_protid.selectSingleNode("txt_17").Text
        medarb_protid_txt_018 = oNode_medarb_protid.selectSingleNode("txt_18").Text
        medarb_protid_txt_019 = oNode_medarb_protid.selectSingleNode("txt_19").Text
        medarb_protid_txt_020 = oNode_medarb_protid.selectSingleNode("txt_20").Text
        medarb_protid_txt_021 = oNode_medarb_protid.selectSingleNode("txt_21").Text
        medarb_protid_txt_022 = oNode_medarb_protid.selectSingleNode("txt_22").Text

          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>