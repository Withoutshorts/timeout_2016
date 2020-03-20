
<% 
Dim objXMLHTTP_leverand, objXMLDOM_leverand, i_leverand, strHTML_leverand

Set objXMLDom_leverand = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_leverand = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_leverand.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/leverand.xml", False
'objXmlHttp_leverand.open "GET", "http://localhost/inc/xml/leverand.xml", False
'objXmlHttp_leverand.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/leverand.xml", False
'objXmlHttp_leverand.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/leverand.xml", False
'objXmlHttp_leverand.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/leverand.xml", False
'objXmlHttp_leverand.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/leverand.xml", False
objXmlHttp_leverand.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/leverand.xml", False

objXmlHttp_leverand.send


Set objXmlDom_leverand = objXmlHttp_leverand.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_leverand = Nothing



Dim Address_leverand, Latitude_leverand, Longitude_leverand
Dim oNode_leverand, oNodes_leverand
Dim sXPathQuery_leverand

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
sXPathQuery_leverand = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_leverand = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_leverand = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_leverand = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_leverand = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_leverand = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_leverand = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_leverand = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_leverand = objXmlDom_leverand.documentElement.selectSingleNode(sXPathQuery_leverand)
Address_leverand = oNode_leverand.Text

Set oNodes_leverand = objXmlDom_leverand.documentElement.selectNodes(sXPathQuery_leverand)

    For Each oNode_leverand in oNodes_leverand

        leverand_txt_001 = oNode_leverand.selectSingleNode("txt_1").Text
        leverand_txt_002 = oNode_leverand.selectSingleNode("txt_2").Text
        leverand_txt_003 = oNode_leverand.selectSingleNode("txt_3").Text
        leverand_txt_004 = oNode_leverand.selectSingleNode("txt_4").Text
        leverand_txt_005 = oNode_leverand.selectSingleNode("txt_5").Text
        leverand_txt_006 = oNode_leverand.selectSingleNode("txt_6").Text
        leverand_txt_007 = oNode_leverand.selectSingleNode("txt_7").Text
        leverand_txt_008 = oNode_leverand.selectSingleNode("txt_8").Text
        leverand_txt_009 = oNode_leverand.selectSingleNode("txt_9").Text
        leverand_txt_010 = oNode_leverand.selectSingleNode("txt_10").Text

        leverand_txt_011 = oNode_leverand.selectSingleNode("txt_11").Text
        leverand_txt_012 = oNode_leverand.selectSingleNode("txt_12").Text
        leverand_txt_013 = oNode_leverand.selectSingleNode("txt_13").Text
        leverand_txt_014 = oNode_leverand.selectSingleNode("txt_14").Text
        leverand_txt_015 = oNode_leverand.selectSingleNode("txt_15").Text
        leverand_txt_016 = oNode_leverand.selectSingleNode("txt_16").Text
        leverand_txt_017 = oNode_leverand.selectSingleNode("txt_17").Text
        leverand_txt_018 = oNode_leverand.selectSingleNode("txt_18").Text
        leverand_txt_019 = oNode_leverand.selectSingleNode("txt_19").Text
        leverand_txt_020 = oNode_leverand.selectSingleNode("txt_20").Text

        leverand_txt_021 = oNode_leverand.selectSingleNode("txt_21").Text
        leverand_txt_022 = oNode_leverand.selectSingleNode("txt_22").Text
        leverand_txt_023 = oNode_leverand.selectSingleNode("txt_23").Text
        leverand_txt_024 = oNode_leverand.selectSingleNode("txt_24").Text
        leverand_txt_025 = oNode_leverand.selectSingleNode("txt_25").Text
        leverand_txt_026 = oNode_leverand.selectSingleNode("txt_26").Text


          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>