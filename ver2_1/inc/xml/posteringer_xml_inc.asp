
<% 
Dim objXMLHTTP_posteringer, objXMLDOM_posteringer, i_posteringer, strHTML_posteringer

Set objXMLDom_posteringer = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_posteringer = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_posteringer.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/posteringer_sprog.xml", False
'objXmlHttp_posteringer.open "GET", "http://localhost/inc/xml/posteringer_sprog.xml", False
'objXmlHttp_posteringer.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/posteringer_sprog.xml", False
'objXmlHttp_posteringer.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/posteringer_sprog.xml", False
'objXmlHttp_posteringer.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/posteringer_sprog.xml", False
'objXmlHttp_posteringer.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/posteringer_sprog.xml", False
objXmlHttp_posteringer.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/posteringer_sprog.xml", False

objXmlHttp_posteringer.send


Set objXmlDom_posteringer = objXmlHttp_posteringer.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_posteringer = Nothing



Dim Address_posteringer, Latitude_posteringer, Longitude_posteringer
Dim oNode_posteringer, oNodes_posteringer
Dim sXPathQuery_posteringer

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
sXPathQuery_posteringer = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_posteringer = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_posteringer = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_posteringer = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_posteringer = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_posteringer = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_posteringer = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_posteringer = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_posteringer = objXmlDom_posteringer.documentElement.selectSingleNode(sXPathQuery_posteringer)
Address_posteringer = oNode_posteringer.Text

Set oNodes_posteringer = objXmlDom_posteringer.documentElement.selectNodes(sXPathQuery_posteringer)

    For Each oNode_posteringer in oNodes_posteringer

        posteringer_txt_001 = oNode_posteringer.selectSingleNode("txt_1").Text
        posteringer_txt_002 = oNode_posteringer.selectSingleNode("txt_2").Text
        posteringer_txt_003 = oNode_posteringer.selectSingleNode("txt_3").Text
        posteringer_txt_004 = oNode_posteringer.selectSingleNode("txt_4").Text
        posteringer_txt_005 = oNode_posteringer.selectSingleNode("txt_5").Text
        posteringer_txt_006 = oNode_posteringer.selectSingleNode("txt_6").Text
        posteringer_txt_007 = oNode_posteringer.selectSingleNode("txt_7").Text
        posteringer_txt_008 = oNode_posteringer.selectSingleNode("txt_8").Text
        posteringer_txt_009 = oNode_posteringer.selectSingleNode("txt_9").Text
        posteringer_txt_010 = oNode_posteringer.selectSingleNode("txt_10").Text

        posteringer_txt_011 = oNode_posteringer.selectSingleNode("txt_11").Text
        posteringer_txt_012 = oNode_posteringer.selectSingleNode("txt_12").Text
        posteringer_txt_013 = oNode_posteringer.selectSingleNode("txt_13").Text
        posteringer_txt_014 = oNode_posteringer.selectSingleNode("txt_14").Text
        posteringer_txt_015 = oNode_posteringer.selectSingleNode("txt_15").Text
        posteringer_txt_016 = oNode_posteringer.selectSingleNode("txt_16").Text
        posteringer_txt_017 = oNode_posteringer.selectSingleNode("txt_17").Text
        posteringer_txt_018 = oNode_posteringer.selectSingleNode("txt_18").Text
        posteringer_txt_019 = oNode_posteringer.selectSingleNode("txt_19").Text
        posteringer_txt_020 = oNode_posteringer.selectSingleNode("txt_20").Text

        posteringer_txt_021 = oNode_posteringer.selectSingleNode("txt_21").Text
        posteringer_txt_022 = oNode_posteringer.selectSingleNode("txt_22").Text
        posteringer_txt_023 = oNode_posteringer.selectSingleNode("txt_23").Text
        posteringer_txt_024 = oNode_posteringer.selectSingleNode("txt_24").Text
        posteringer_txt_025 = oNode_posteringer.selectSingleNode("txt_25").Text
        posteringer_txt_026 = oNode_posteringer.selectSingleNode("txt_26").Text
        posteringer_txt_027 = oNode_posteringer.selectSingleNode("txt_27").Text
        posteringer_txt_028 = oNode_posteringer.selectSingleNode("txt_28").Text
        posteringer_txt_029 = oNode_posteringer.selectSingleNode("txt_29").Text
        posteringer_txt_030 = oNode_posteringer.selectSingleNode("txt_30").Text

        posteringer_txt_031 = oNode_posteringer.selectSingleNode("txt_31").Text
        posteringer_txt_032 = oNode_posteringer.selectSingleNode("txt_32").Text
        posteringer_txt_033 = oNode_posteringer.selectSingleNode("txt_33").Text
        posteringer_txt_034 = oNode_posteringer.selectSingleNode("txt_34").Text
        posteringer_txt_035 = oNode_posteringer.selectSingleNode("txt_35").Text
        posteringer_txt_036 = oNode_posteringer.selectSingleNode("txt_36").Text
        posteringer_txt_037 = oNode_posteringer.selectSingleNode("txt_37").Text
        posteringer_txt_038 = oNode_posteringer.selectSingleNode("txt_38").Text
        posteringer_txt_039 = oNode_posteringer.selectSingleNode("txt_39").Text
        posteringer_txt_040 = oNode_posteringer.selectSingleNode("txt_40").Text


    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>