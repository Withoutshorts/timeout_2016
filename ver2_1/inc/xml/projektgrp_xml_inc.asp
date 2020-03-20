
<% 
Dim objXMLHTTP_progrp, objXMLDOM_progrp, i_progrp, strHTML_progrp

Set objXMLDom_progrp = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_progrp = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_progrp.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/projektgrp_sprog.xml", False
'objXmlHttp_progrp.open "GET", "http://localhost/inc/xml/progrp_sprog.xml", False
'objXmlHttp_progrp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/projektgrp_sprog.xml", False
'objXmlHttp_progrp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/projektgrp_sprog.xml", False
'objXmlHttp_progrp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/projektgrp_sprog.xml", False
'objXmlHttp_progrp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/projektgrp_sprog.xml", False
objXmlHttp_progrp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/projektgrp_sprog.xml", False

objXmlHttp_progrp.send


Set objXmlDom_progrp = objXmlHttp_progrp.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_progrp = Nothing



Dim Address_progrp, Latitude_progrp, Longitude_progrp
Dim oNode_progrp, oNodes_progrp
Dim sXPathQuery_progrp

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
sXPathQuery_progrp = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_progrp = "//sprog/uk"
if lto = "a27" then
    sXPathQuery_progrp = "//sprog/trainerlog_uk"
end if
'Session.LCID = 2057
case 3
sXPathQuery_progrp = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_progrp = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_progrp = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_progrp = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_progrp = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_progrp = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_progrp = objXmlDom_progrp.documentElement.selectSingleNode(sXPathQuery_progrp)
Address_progrp = oNode_progrp.Text

Set oNodes_progrp = objXmlDom_progrp.documentElement.selectNodes(sXPathQuery_progrp)

    For Each oNode_progrp in oNodes_progrp

        progrp_txt_001 = oNode_progrp.selectSingleNode("txt_1").Text
        progrp_txt_002 = oNode_progrp.selectSingleNode("txt_2").Text
        progrp_txt_003 = oNode_progrp.selectSingleNode("txt_3").Text
        progrp_txt_004 = oNode_progrp.selectSingleNode("txt_4").Text
        progrp_txt_005 = oNode_progrp.selectSingleNode("txt_5").Text
        progrp_txt_006 = oNode_progrp.selectSingleNode("txt_6").Text
        progrp_txt_007 = oNode_progrp.selectSingleNode("txt_7").Text
        progrp_txt_008 = oNode_progrp.selectSingleNode("txt_8").Text
        progrp_txt_009 = oNode_progrp.selectSingleNode("txt_9").Text
        progrp_txt_010 = oNode_progrp.selectSingleNode("txt_10").Text

        progrp_txt_011 = oNode_progrp.selectSingleNode("txt_11").Text
        progrp_txt_012 = oNode_progrp.selectSingleNode("txt_12").Text
        progrp_txt_013 = oNode_progrp.selectSingleNode("txt_13").Text
        progrp_txt_014 = oNode_progrp.selectSingleNode("txt_14").Text
        progrp_txt_015 = oNode_progrp.selectSingleNode("txt_15").Text
        progrp_txt_016 = oNode_progrp.selectSingleNode("txt_16").Text
        progrp_txt_017 = oNode_progrp.selectSingleNode("txt_17").Text
        progrp_txt_018 = oNode_progrp.selectSingleNode("txt_18").Text
        progrp_txt_019 = oNode_progrp.selectSingleNode("txt_19").Text
        progrp_txt_020 = oNode_progrp.selectSingleNode("txt_20").Text

        progrp_txt_021 = oNode_progrp.selectSingleNode("txt_21").Text
        progrp_txt_022 = oNode_progrp.selectSingleNode("txt_22").Text
        progrp_txt_023 = oNode_progrp.selectSingleNode("txt_23").Text
        progrp_txt_024 = oNode_progrp.selectSingleNode("txt_24").Text
        progrp_txt_025 = oNode_progrp.selectSingleNode("txt_25").Text
        progrp_txt_026 = oNode_progrp.selectSingleNode("txt_26").Text
        progrp_txt_027 = oNode_progrp.selectSingleNode("txt_27").Text
        progrp_txt_028 = oNode_progrp.selectSingleNode("txt_28").Text
        progrp_txt_029 = oNode_progrp.selectSingleNode("txt_29").Text
        progrp_txt_030 = oNode_progrp.selectSingleNode("txt_30").Text

        progrp_txt_031 = oNode_progrp.selectSingleNode("txt_31").Text
        progrp_txt_032 = oNode_progrp.selectSingleNode("txt_32").Text
        progrp_txt_033 = oNode_progrp.selectSingleNode("txt_33").Text
        progrp_txt_034 = oNode_progrp.selectSingleNode("txt_34").Text
        progrp_txt_035 = oNode_progrp.selectSingleNode("txt_35").Text
        progrp_txt_036 = oNode_progrp.selectSingleNode("txt_36").Text
        progrp_txt_037 = oNode_progrp.selectSingleNode("txt_37").Text
        progrp_txt_038 = oNode_progrp.selectSingleNode("txt_38").Text
        progrp_txt_039 = oNode_progrp.selectSingleNode("txt_39").Text
        progrp_txt_040 = oNode_progrp.selectSingleNode("txt_40").Text

        progrp_txt_041 = oNode_progrp.selectSingleNode("txt_41").Text
        progrp_txt_042 = oNode_progrp.selectSingleNode("txt_42").Text
        progrp_txt_043 = oNode_progrp.selectSingleNode("txt_43").Text
        progrp_txt_044 = oNode_progrp.selectSingleNode("txt_44").Text
        progrp_txt_045 = oNode_progrp.selectSingleNode("txt_45").Text
        progrp_txt_046 = oNode_progrp.selectSingleNode("txt_46").Text
        progrp_txt_047 = oNode_progrp.selectSingleNode("txt_47").Text
        progrp_txt_048 = oNode_progrp.selectSingleNode("txt_48").Text
        progrp_txt_049 = oNode_progrp.selectSingleNode("txt_49").Text
        progrp_txt_050 = oNode_progrp.selectSingleNode("txt_50").Text
        progrp_txt_051 = oNode_progrp.selectSingleNode("txt_51").Text


  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>