
<% 
Dim objXMLHTTP_favorit, objXMLDOM_favorit, i_favorit, strHTML_favorit

Set objXMLDom_favorit = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_favorit = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_favorit.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/favorit_sprog.xml", False
'objXmlHttp_favorit.open "GET", "http://localhost/inc/xml/favorit_sprog.xml", False
'objXmlHttp_favorit.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/favorit_sprog.xml", False
objXmlHttp_favorit.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/favorit_sprog.xml", False
'objXmlHttp_favorit.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/favorit_sprog.xml", False
'objXmlHttp_favorit.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/favorit_sprog.xml", False

objXmlHttp_favorit.send


Set objXmlDom_favorit = objXmlHttp_favorit.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_favorit = Nothing



Dim Address_favorit, Latitude_favorit, Longitude_favorit
Dim oNode_favorit, oNodes_favorit
Dim sXPathQuery_favorit

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
sXPathQuery_favorit = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_favorit = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_favorit = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_favorit = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_favorit = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_favorit = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_favorit = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_favorit = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_favorit = objXmlDom_favorit.documentElement.selectSingleNode(sXPathQuery_favorit)
Address_favorit = oNode_favorit.Text

Set oNodes_favorit = objXmlDom_favorit.documentElement.selectNodes(sXPathQuery_favorit)

    For Each oNode_favorit in oNodes_favorit

        favorit_txt_001 = oNode_favorit.selectSingleNode("txt_1").Text
        favorit_txt_002 = oNode_favorit.selectSingleNode("txt_2").Text
        favorit_txt_003 = oNode_favorit.selectSingleNode("txt_3").Text
        favorit_txt_003 = oNode_favorit.selectSingleNode("txt_3").Text
        favorit_txt_004 = oNode_favorit.selectSingleNode("txt_4").Text
        favorit_txt_005 = oNode_favorit.selectSingleNode("txt_5").Text

        favorit_txt_006 = oNode_favorit.selectSingleNode("txt_6").Text
        favorit_txt_007 = oNode_favorit.selectSingleNode("txt_7").Text
        favorit_txt_008 = oNode_favorit.selectSingleNode("txt_8").Text
        favorit_txt_009 = oNode_favorit.selectSingleNode("txt_9").Text
        favorit_txt_010 = oNode_favorit.selectSingleNode("txt_10").Text
        favorit_txt_011 = oNode_favorit.selectSingleNode("txt_11").Text
    
        favorit_txt_012 = oNode_favorit.selectSingleNode("txt_12").Text
        favorit_txt_013 = oNode_favorit.selectSingleNode("txt_13").Text
        favorit_txt_014 = oNode_favorit.selectSingleNode("txt_14").Text
        favorit_txt_015 = oNode_favorit.selectSingleNode("txt_15").Text
        favorit_txt_016 = oNode_favorit.selectSingleNode("txt_16").Text
        favorit_txt_017 = oNode_favorit.selectSingleNode("txt_17").Text
    
        favorit_txt_018 = oNode_favorit.selectSingleNode("txt_18").Text
        favorit_txt_019 = oNode_favorit.selectSingleNode("txt_19").Text
        favorit_txt_020 = oNode_favorit.selectSingleNode("txt_20").Text
        favorit_txt_021 = oNode_favorit.selectSingleNode("txt_21").Text

        favorit_txt_022 = oNode_favorit.selectSingleNode("txt_22").Text
        favorit_txt_023 = oNode_favorit.selectSingleNode("txt_23").Text
        favorit_txt_024 = oNode_favorit.selectSingleNode("txt_24").Text
        favorit_txt_025 = oNode_favorit.selectSingleNode("txt_25").Text
        favorit_txt_026 = oNode_favorit.selectSingleNode("txt_26").Text
        favorit_txt_027 = oNode_favorit.selectSingleNode("txt_27").Text
        favorit_txt_028 = oNode_favorit.selectSingleNode("txt_28").Text

        favorit_txt_029 = oNode_favorit.selectSingleNode("txt_29").Text
        favorit_txt_030 = oNode_favorit.selectSingleNode("txt_30").Text
        favorit_txt_031 = oNode_favorit.selectSingleNode("txt_31").Text
        favorit_txt_032 = oNode_favorit.selectSingleNode("txt_32").Text
  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>