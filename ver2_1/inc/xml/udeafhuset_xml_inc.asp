
<% 
Dim objXMLHTTP_udeafhuset, objXMLDOM_udeafhuset, i_udeafhuset, strHTML_udeafhuset

Set objXMLDom_udeafhuset = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_udeafhuset = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_udeafhuset.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/udeafhuset.xml", False
'objXmlHttp_udeafhuset.open "GET", "http://localhost/inc/xml/udeafhuset.xml", False
'objXmlHttp_udeafhuset.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/udeafhuset.xml", False
'objXmlHttp_udeafhuset.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/udeafhuset.xml", False
'objXmlHttp_udeafhuset.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/udeafhuset.xml", False
'objXmlHttp_udeafhuset.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/udeafhuset.xml", False
objXmlHttp_udeafhuset.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/udeafhuset.xml", False

objXmlHttp_udeafhuset.send


Set objXmlDom_udeafhuset = objXmlHttp_udeafhuset.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_udeafhuset = Nothing



Dim Address_udeafhuset, Latitude_udeafhuset, Longitude_udeafhuset
Dim oNode_udeafhuset, oNodes_udeafhuset
Dim sXPathQuery_udeafhuset

sprog = 1 'DK
if len(trim(session("mid"))) <> 0 then
strSQL = "SELECT sprog FROM medarbejdere WHERE mid = " & session("mid")

oRec.open strSQL, oConn, 3
if not oRec.EOF then
sprog = oRec("sprog")
end if
oRec.close
end if

if lto = "cool" then
    sprog = 2
end if

select case sprog
case 1
sXPathQuery_udeafhuset = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_udeafhuset = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_udeafhuset = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_udeafhuset = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_udeafhuset = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_udeafhuset = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_udeafhuset = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_udeafhuset = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_udeafhuset = objXmlDom_udeafhuset.documentElement.selectSingleNode(sXPathQuery_udeafhuset)
Address_udeafhuset = oNode_udeafhuset.Text

Set oNodes_udeafhuset = objXmlDom_udeafhuset.documentElement.selectNodes(sXPathQuery_udeafhuset)

    For Each oNode_udeafhuset in oNodes_udeafhuset

        udeafhuset_txt_001 = oNode_udeafhuset.selectSingleNode("txt_1").Text
        udeafhuset_txt_002 = oNode_udeafhuset.selectSingleNode("txt_2").Text
        udeafhuset_txt_003 = oNode_udeafhuset.selectSingleNode("txt_3").Text
        udeafhuset_txt_004 = oNode_udeafhuset.selectSingleNode("txt_4").Text
        udeafhuset_txt_005 = oNode_udeafhuset.selectSingleNode("txt_5").Text
        udeafhuset_txt_006 = oNode_udeafhuset.selectSingleNode("txt_6").Text
        udeafhuset_txt_007 = oNode_udeafhuset.selectSingleNode("txt_7").Text
        udeafhuset_txt_008 = oNode_udeafhuset.selectSingleNode("txt_8").Text
        udeafhuset_txt_009 = oNode_udeafhuset.selectSingleNode("txt_9").Text
        udeafhuset_txt_010 = oNode_udeafhuset.selectSingleNode("txt_10").Text

        udeafhuset_txt_011 = oNode_udeafhuset.selectSingleNode("txt_11").Text
        udeafhuset_txt_012 = oNode_udeafhuset.selectSingleNode("txt_12").Text
        udeafhuset_txt_013 = oNode_udeafhuset.selectSingleNode("txt_13").Text
        udeafhuset_txt_014 = oNode_udeafhuset.selectSingleNode("txt_14").Text
        udeafhuset_txt_015 = oNode_udeafhuset.selectSingleNode("txt_15").Text
        udeafhuset_txt_016 = oNode_udeafhuset.selectSingleNode("txt_16").Text
        udeafhuset_txt_017 = oNode_udeafhuset.selectSingleNode("txt_17").Text
        udeafhuset_txt_018 = oNode_udeafhuset.selectSingleNode("txt_18").Text
        udeafhuset_txt_019 = oNode_udeafhuset.selectSingleNode("txt_19").Text
        udeafhuset_txt_020 = oNode_udeafhuset.selectSingleNode("txt_20").Text
        udeafhuset_txt_021 = oNode_udeafhuset.selectSingleNode("txt_21").Text

    
        
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>