
<% 
Dim objXMLHTTP_matgrp, objXMLDOM_matgrp, i_matgrp, strHTML_matgrp

Set objXMLDom_matgrp = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_matgrp = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_matgrp.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/mat_grp_sprog.xml", False
'objXmlHttp_matgrp.open "GET", "http://localhost/inc/xml/mat_grp_sprog.xml", False
'objXmlHttp_matgrp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/mat_grp_sprog.xml", False
'objXmlHttp_matgrp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/mat_grp_sprog.xml", False
'objXmlHttp_matgrp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/mat_grp_sprog.xml", False
'objXmlHttp_matgrp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/mat_grp_sprog.xml", False
objXmlHttp_matgrp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/mat_grp_sprog.xml", False

objXmlHttp_matgrp.send


Set objXmlDom_matgrp = objXmlHttp_matgrp.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_matgrp = Nothing



Dim Address_matgrp, Latitude_matgrp, Longitude_matgrp
Dim oNode_matgrp, oNodes_matgrp
Dim sXPathQuery_matgrp

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
sXPathQuery_matgrp = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_matgrp = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_matgrp = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_matgrp = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_matgrp = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_matgrp = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_matgrp = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_matgrp = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_matgrp = objXmlDom_matgrp.documentElement.selectSingleNode(sXPathQuery_matgrp)
Address_matgrp = oNode_matgrp.Text

Set oNodes_matgrp = objXmlDom_matgrp.documentElement.selectNodes(sXPathQuery_matgrp)

    For Each oNode_matgrp in oNodes_matgrp

        matgrp_txt_001 = oNode_matgrp.selectSingleNode("txt_1").Text
        matgrp_txt_002 = oNode_matgrp.selectSingleNode("txt_2").Text
        matgrp_txt_003 = oNode_matgrp.selectSingleNode("txt_3").Text
        matgrp_txt_004 = oNode_matgrp.selectSingleNode("txt_4").Text
        matgrp_txt_005 = oNode_matgrp.selectSingleNode("txt_5").Text
        matgrp_txt_006 = oNode_matgrp.selectSingleNode("txt_6").Text
        matgrp_txt_007 = oNode_matgrp.selectSingleNode("txt_7").Text
        matgrp_txt_008 = oNode_matgrp.selectSingleNode("txt_8").Text
        matgrp_txt_009 = oNode_matgrp.selectSingleNode("txt_9").Text
        matgrp_txt_010 = oNode_matgrp.selectSingleNode("txt_10").Text

        matgrp_txt_011 = oNode_matgrp.selectSingleNode("txt_11").Text
        matgrp_txt_012 = oNode_matgrp.selectSingleNode("txt_12").Text
        matgrp_txt_013 = oNode_matgrp.selectSingleNode("txt_13").Text
        matgrp_txt_014 = oNode_matgrp.selectSingleNode("txt_14").Text
        matgrp_txt_015 = oNode_matgrp.selectSingleNode("txt_15").Text
        matgrp_txt_016 = oNode_matgrp.selectSingleNode("txt_16").Text
        matgrp_txt_017 = oNode_matgrp.selectSingleNode("txt_17").Text
        matgrp_txt_018 = oNode_matgrp.selectSingleNode("txt_18").Text
        matgrp_txt_019 = oNode_matgrp.selectSingleNode("txt_19").Text
        matgrp_txt_020 = oNode_matgrp.selectSingleNode("txt_20").Text

        matgrp_txt_021 = oNode_matgrp.selectSingleNode("txt_21").Text
        matgrp_txt_022 = oNode_matgrp.selectSingleNode("txt_22").Text
        matgrp_txt_023 = oNode_matgrp.selectSingleNode("txt_23").Text
        matgrp_txt_024 = oNode_matgrp.selectSingleNode("txt_24").Text
        matgrp_txt_025 = oNode_matgrp.selectSingleNode("txt_25").Text
        matgrp_txt_026 = oNode_matgrp.selectSingleNode("txt_26").Text
        matgrp_txt_027 = oNode_matgrp.selectSingleNode("txt_27").Text
        matgrp_txt_028 = oNode_matgrp.selectSingleNode("txt_28").Text
        matgrp_txt_029 = oNode_matgrp.selectSingleNode("txt_29").Text

        matgrp_txt_030 = oNode_matgrp.selectSingleNode("txt_30").Text
        matgrp_txt_031 = oNode_matgrp.selectSingleNode("txt_31").Text
        matgrp_txt_032 = oNode_matgrp.selectSingleNode("txt_32").Text
        matgrp_txt_033 = oNode_matgrp.selectSingleNode("txt_33").Text
        matgrp_txt_034 = oNode_matgrp.selectSingleNode("txt_34").Text
        matgrp_txt_035 = oNode_matgrp.selectSingleNode("txt_35").Text
        matgrp_txt_036 = oNode_matgrp.selectSingleNode("txt_36").Text
        matgrp_txt_037 = oNode_matgrp.selectSingleNode("txt_37").Text
        matgrp_txt_038 = oNode_matgrp.selectSingleNode("txt_38").Text
        matgrp_txt_039 = oNode_matgrp.selectSingleNode("txt_39").Text
        matgrp_txt_040 = oNode_matgrp.selectSingleNode("txt_40").Text
        matgrp_txt_041 = oNode_matgrp.selectSingleNode("txt_41").Text
  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>