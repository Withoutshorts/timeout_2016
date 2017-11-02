
<% 
Dim objXMLHTTP_abonner, objXMLDOM_abonner, i_abonner, strHTML_abonner

Set objXMLDom_abonner = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_abonner = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_abonner.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/abonner_sprog.xml", False
'objXmlHttp_abonner.open "GET", "http://localhost/inc/xml/favorit_sprog.xml", False
'objXmlHttp_abonner.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/abonner_sprog.xml", False
objXmlHttp_abonner.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/abonner_sprog.xml", False
'objXmlHttp_abonner.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/abonner_sprog.xml", False
'objXmlHttp_abonner.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/abonner_sprog.xml", False

objXmlHttp_abonner.send


Set objXmlDom_abonner = objXmlHttp_abonner.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_abonner = Nothing



Dim Address_abonner, Latitude_abonner, Longitude_abonner
Dim oNode_abonner, oNodes_abonner
Dim sXPathQuery_abonner

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
sXPathQuery_abonner = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_abonner = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_abonner = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_abonner = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_abonner = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_abonner = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_abonner = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_abonner = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_abonner = objXmlDom_abonner.documentElement.selectSingleNode(sXPathQuery_abonner)
Address_abonner = oNode_abonner.Text

Set oNodes_abonner = objXmlDom_abonner.documentElement.selectNodes(sXPathQuery_abonner)

    For Each oNode_abonner in oNodes_abonner

    abonner_txt_001 = oNode_abonner.selectSingleNode("txt_1").Text
    abonner_txt_002 = oNode_abonner.selectSingleNode("txt_2").Text
    abonner_txt_003 = oNode_abonner.selectSingleNode("txt_3").Text
    abonner_txt_004 = oNode_abonner.selectSingleNode("txt_4").Text
    abonner_txt_005 = oNode_abonner.selectSingleNode("txt_5").Text
    abonner_txt_006 = oNode_abonner.selectSingleNode("txt_6").Text
    abonner_txt_007 = oNode_abonner.selectSingleNode("txt_7").Text
    abonner_txt_008 = oNode_abonner.selectSingleNode("txt_8").Text
    abonner_txt_009 = oNode_abonner.selectSingleNode("txt_9").Text
    abonner_txt_010 = oNode_abonner.selectSingleNode("txt_10").Text

    abonner_txt_011 = oNode_abonner.selectSingleNode("txt_11").Text
    abonner_txt_012 = oNode_abonner.selectSingleNode("txt_12").Text
    abonner_txt_013 = oNode_abonner.selectSingleNode("txt_13").Text
    abonner_txt_014 = oNode_abonner.selectSingleNode("txt_14").Text
    abonner_txt_015 = oNode_abonner.selectSingleNode("txt_15").Text
    abonner_txt_016 = oNode_abonner.selectSingleNode("txt_16").Text
    abonner_txt_017 = oNode_abonner.selectSingleNode("txt_17").Text
    abonner_txt_018 = oNode_abonner.selectSingleNode("txt_18").Text
    abonner_txt_019 = oNode_abonner.selectSingleNode("txt_19").Text
    abonner_txt_020 = oNode_abonner.selectSingleNode("txt_20").Text

    abonner_txt_021 = oNode_abonner.selectSingleNode("txt_21").Text
    abonner_txt_022 = oNode_abonner.selectSingleNode("txt_22").Text
    abonner_txt_023 = oNode_abonner.selectSingleNode("txt_23").Text
    abonner_txt_024 = oNode_abonner.selectSingleNode("txt_24").Text
    abonner_txt_025 = oNode_abonner.selectSingleNode("txt_25").Text
    abonner_txt_026 = oNode_abonner.selectSingleNode("txt_26").Text
    abonner_txt_027 = oNode_abonner.selectSingleNode("txt_27").Text
    abonner_txt_028 = oNode_abonner.selectSingleNode("txt_28").Text
    abonner_txt_029 = oNode_abonner.selectSingleNode("txt_29").Text
    abonner_txt_030 = oNode_abonner.selectSingleNode("txt_30").Text

    abonner_txt_031 = oNode_abonner.selectSingleNode("txt_31").Text
    abonner_txt_032 = oNode_abonner.selectSingleNode("txt_32").Text
    abonner_txt_033 = oNode_abonner.selectSingleNode("txt_33").Text
    abonner_txt_034 = oNode_abonner.selectSingleNode("txt_34").Text


       
  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>