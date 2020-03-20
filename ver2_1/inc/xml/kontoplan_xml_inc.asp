
<% 
Dim objXMLHTTP_kontoplan, objXMLDOM_kontoplan, i_kontoplan, strHTML_kontoplan

Set objXMLDom_kontoplan = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_kontoplan = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_kontoplan.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/kontoplan_sprog.xml", False
'objXmlHttp_kontoplan.open "GET", "http://localhost/inc/xml/favorit_sprog.xml", False
'objXmlHttp_kontoplan.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/kontoplan_sprog.xml", False
'objXmlHttp_kontoplan.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/kontoplan_sprog.xml", False
'objXmlHttp_kontoplan.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/kontoplan_sprog.xml", False
'objXmlHttp_kontoplan.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/kontoplan_sprog.xml", False
objXmlHttp_kontoplan.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/kontoplan_sprog.xml", False

objXmlHttp_kontoplan.send


Set objXmlDom_kontoplan = objXmlHttp_kontoplan.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_kontoplan = Nothing



Dim Address_kontoplan, Latitude_kontoplan, Longitude_kontoplan
Dim oNode_kontoplan, oNodes_kontoplan
Dim sXPathQuery_kontoplan

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
sXPathQuery_kontoplan = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_kontoplan = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_kontoplan = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_kontoplan = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_kontoplan = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_kontoplan = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_kontoplan = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_kontoplan = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_kontoplan = objXmlDom_kontoplan.documentElement.selectSingleNode(sXPathQuery_kontoplan)
Address_kontoplan = oNode_kontoplan.Text

Set oNodes_kontoplan = objXmlDom_kontoplan.documentElement.selectNodes(sXPathQuery_kontoplan)

    For Each oNode_kontoplan in oNodes_kontoplan

        kontoplan_txt_001 = oNode_kontoplan.selectSingleNode("txt_1").Text
        kontoplan_txt_002 = oNode_kontoplan.selectSingleNode("txt_2").Text
        kontoplan_txt_003 = oNode_kontoplan.selectSingleNode("txt_3").Text
        kontoplan_txt_004 = oNode_kontoplan.selectSingleNode("txt_4").Text
        kontoplan_txt_005 = oNode_kontoplan.selectSingleNode("txt_5").Text
        kontoplan_txt_006 = oNode_kontoplan.selectSingleNode("txt_6").Text
        kontoplan_txt_007 = oNode_kontoplan.selectSingleNode("txt_7").Text
        kontoplan_txt_008 = oNode_kontoplan.selectSingleNode("txt_8").Text
        kontoplan_txt_009 = oNode_kontoplan.selectSingleNode("txt_9").Text
        kontoplan_txt_010 = oNode_kontoplan.selectSingleNode("txt_10").Text
        
        kontoplan_txt_011 = oNode_kontoplan.selectSingleNode("txt_11").Text
        kontoplan_txt_012 = oNode_kontoplan.selectSingleNode("txt_12").Text
        kontoplan_txt_013 = oNode_kontoplan.selectSingleNode("txt_13").Text
        kontoplan_txt_014 = oNode_kontoplan.selectSingleNode("txt_14").Text
        kontoplan_txt_015 = oNode_kontoplan.selectSingleNode("txt_15").Text
        kontoplan_txt_016 = oNode_kontoplan.selectSingleNode("txt_16").Text
        kontoplan_txt_017 = oNode_kontoplan.selectSingleNode("txt_17").Text
        kontoplan_txt_018 = oNode_kontoplan.selectSingleNode("txt_18").Text
        kontoplan_txt_019 = oNode_kontoplan.selectSingleNode("txt_19").Text
        kontoplan_txt_020 = oNode_kontoplan.selectSingleNode("txt_20").Text

        kontoplan_txt_021 = oNode_kontoplan.selectSingleNode("txt_21").Text
        kontoplan_txt_022 = oNode_kontoplan.selectSingleNode("txt_22").Text
        kontoplan_txt_023 = oNode_kontoplan.selectSingleNode("txt_23").Text
        kontoplan_txt_024 = oNode_kontoplan.selectSingleNode("txt_24").Text
        kontoplan_txt_025 = oNode_kontoplan.selectSingleNode("txt_25").Text
        kontoplan_txt_026 = oNode_kontoplan.selectSingleNode("txt_26").Text
        kontoplan_txt_027 = oNode_kontoplan.selectSingleNode("txt_27").Text
        kontoplan_txt_028 = oNode_kontoplan.selectSingleNode("txt_28").Text
        kontoplan_txt_029 = oNode_kontoplan.selectSingleNode("txt_29").Text
        kontoplan_txt_030 = oNode_kontoplan.selectSingleNode("txt_30").Text

        kontoplan_txt_031 = oNode_kontoplan.selectSingleNode("txt_31").Text
        kontoplan_txt_032 = oNode_kontoplan.selectSingleNode("txt_32").Text
        kontoplan_txt_033 = oNode_kontoplan.selectSingleNode("txt_33").Text
        kontoplan_txt_034 = oNode_kontoplan.selectSingleNode("txt_34").Text
        kontoplan_txt_035 = oNode_kontoplan.selectSingleNode("txt_35").Text
        kontoplan_txt_036 = oNode_kontoplan.selectSingleNode("txt_36").Text
        kontoplan_txt_037 = oNode_kontoplan.selectSingleNode("txt_37").Text
        kontoplan_txt_038 = oNode_kontoplan.selectSingleNode("txt_38").Text
        kontoplan_txt_039 = oNode_kontoplan.selectSingleNode("txt_39").Text
        kontoplan_txt_040 = oNode_kontoplan.selectSingleNode("txt_40").Text

        kontoplan_txt_041 = oNode_kontoplan.selectSingleNode("txt_41").Text
        kontoplan_txt_042 = oNode_kontoplan.selectSingleNode("txt_42").Text
        kontoplan_txt_043 = oNode_kontoplan.selectSingleNode("txt_43").Text
        kontoplan_txt_044 = oNode_kontoplan.selectSingleNode("txt_44").Text
        kontoplan_txt_045 = oNode_kontoplan.selectSingleNode("txt_45").Text
        kontoplan_txt_046 = oNode_kontoplan.selectSingleNode("txt_46").Text
        kontoplan_txt_047 = oNode_kontoplan.selectSingleNode("txt_47").Text
        kontoplan_txt_048 = oNode_kontoplan.selectSingleNode("txt_48").Text
        kontoplan_txt_049 = oNode_kontoplan.selectSingleNode("txt_49").Text
        kontoplan_txt_050 = oNode_kontoplan.selectSingleNode("txt_50").Text
        kontoplan_txt_051 = oNode_kontoplan.selectSingleNode("txt_51").Text


  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>