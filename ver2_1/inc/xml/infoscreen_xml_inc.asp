
<% 
Dim objXMLHTTP_infoscreen, objXMLDOM_infoscreen, i_infoscreen, strHTML_infoscreen

Set objXMLDom_infoscreen = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_infoscreen = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_infoscreen.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/infoscreen_sprog.xml", False
'objXmlHttp_infoscreen.open "GET", "http://localhost/inc/xml/infoscreen_sprog.xml", False
'objXmlHttp_infoscreen.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/infoscreen_sprog.xml", False
'objXmlHttp_infoscreen.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/infoscreen_sprog.xml", False
'objXmlHttp_infoscreen.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/infoscreen_sprog.xml", False
'objXmlHttp_infoscreen.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/infoscreen_sprog.xml", False
objXmlHttp_infoscreen.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/infoscreen_sprog.xml", False

objXmlHttp_infoscreen.send


Set objXmlDom_infoscreen = objXmlHttp_infoscreen.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_infoscreen = Nothing



Dim Address_infoscreen, Latitude_infoscreen, Longitude_infoscreen
Dim oNode_infoscreen, oNodes_infoscreen
Dim sXPathQuery_infoscreen

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
sXPathQuery_infoscreen = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_infoscreen = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_infoscreen = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_infoscreen = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_infoscreen = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_infoscreen = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_infoscreen = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_infoscreen = "//sprog/dk"
'Session.LCID = 1030
end select

if lto = "cool" then
    sXPathQuery_infoscreen = "//sprog/uk"
end if
'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_infoscreen = objXmlDom_infoscreen.documentElement.selectSingleNode(sXPathQuery_infoscreen)
Address_infoscreen = oNode_infoscreen.Text

Set oNodes_infoscreen = objXmlDom_infoscreen.documentElement.selectNodes(sXPathQuery_infoscreen)

    For Each oNode_infoscreen in oNodes_infoscreen

        infoscreen_txt_001 = oNode_infoscreen.selectSingleNode("txt_1").Text
        infoscreen_txt_002 = oNode_infoscreen.selectSingleNode("txt_2").Text
        infoscreen_txt_003 = oNode_infoscreen.selectSingleNode("txt_3").Text
        infoscreen_txt_003 = oNode_infoscreen.selectSingleNode("txt_3").Text
        infoscreen_txt_004 = oNode_infoscreen.selectSingleNode("txt_4").Text
        infoscreen_txt_005 = oNode_infoscreen.selectSingleNode("txt_5").Text

        infoscreen_txt_006 = oNode_infoscreen.selectSingleNode("txt_6").Text
        infoscreen_txt_007 = oNode_infoscreen.selectSingleNode("txt_7").Text
        infoscreen_txt_008 = oNode_infoscreen.selectSingleNode("txt_8").Text
        infoscreen_txt_009 = oNode_infoscreen.selectSingleNode("txt_9").Text
        infoscreen_txt_010 = oNode_infoscreen.selectSingleNode("txt_10").Text
        infoscreen_txt_011 = oNode_infoscreen.selectSingleNode("txt_11").Text
    
        infoscreen_txt_012 = oNode_infoscreen.selectSingleNode("txt_12").Text
        infoscreen_txt_013 = oNode_infoscreen.selectSingleNode("txt_13").Text
        infoscreen_txt_014 = oNode_infoscreen.selectSingleNode("txt_14").Text
        infoscreen_txt_015 = oNode_infoscreen.selectSingleNode("txt_15").Text
        infoscreen_txt_016 = oNode_infoscreen.selectSingleNode("txt_16").Text
        infoscreen_txt_017 = oNode_infoscreen.selectSingleNode("txt_17").Text
    
        infoscreen_txt_018 = oNode_infoscreen.selectSingleNode("txt_18").Text
        infoscreen_txt_019 = oNode_infoscreen.selectSingleNode("txt_19").Text
        infoscreen_txt_020 = oNode_infoscreen.selectSingleNode("txt_20").Text
        infoscreen_txt_021 = oNode_infoscreen.selectSingleNode("txt_21").Text

        infoscreen_txt_022 = oNode_infoscreen.selectSingleNode("txt_22").Text
        infoscreen_txt_023 = oNode_infoscreen.selectSingleNode("txt_23").Text
        infoscreen_txt_024 = oNode_infoscreen.selectSingleNode("txt_24").Text
        infoscreen_txt_025 = oNode_infoscreen.selectSingleNode("txt_25").Text
        infoscreen_txt_026 = oNode_infoscreen.selectSingleNode("txt_26").Text
        infoscreen_txt_027 = oNode_infoscreen.selectSingleNode("txt_27").Text
        infoscreen_txt_028 = oNode_infoscreen.selectSingleNode("txt_28").Text

        infoscreen_txt_029 = oNode_infoscreen.selectSingleNode("txt_29").Text
        infoscreen_txt_030 = oNode_infoscreen.selectSingleNode("txt_30").Text
        infoscreen_txt_031 = oNode_infoscreen.selectSingleNode("txt_31").Text
        infoscreen_txt_032 = oNode_infoscreen.selectSingleNode("txt_32").Text
        infoscreen_txt_033 = oNode_infoscreen.selectSingleNode("txt_33").Text
        infoscreen_txt_034 = oNode_infoscreen.selectSingleNode("txt_34").Text
        infoscreen_txt_035 = oNode_infoscreen.selectSingleNode("txt_35").Text
        infoscreen_txt_036 = oNode_infoscreen.selectSingleNode("txt_36").Text
        infoscreen_txt_037 = oNode_infoscreen.selectSingleNode("txt_37").Text
        infoscreen_txt_038 = oNode_infoscreen.selectSingleNode("txt_38").Text
        infoscreen_txt_039 = oNode_infoscreen.selectSingleNode("txt_39").Text
        infoscreen_txt_040 = oNode_infoscreen.selectSingleNode("txt_40").Text
        infoscreen_txt_041 = oNode_infoscreen.selectSingleNode("txt_41").Text
        infoscreen_txt_042 = oNode_infoscreen.selectSingleNode("txt_42").Text
        infoscreen_txt_043 = oNode_infoscreen.selectSingleNode("txt_43").Text
        infoscreen_txt_044 = oNode_infoscreen.selectSingleNode("txt_44").Text
        infoscreen_txt_045 = oNode_infoscreen.selectSingleNode("txt_45").Text
        infoscreen_txt_046 = oNode_infoscreen.selectSingleNode("txt_46").Text
        infoscreen_txt_047 = oNode_infoscreen.selectSingleNode("txt_47").Text
        infoscreen_txt_048 = oNode_infoscreen.selectSingleNode("txt_48").Text
        infoscreen_txt_049 = oNode_infoscreen.selectSingleNode("txt_49").Text
        infoscreen_txt_050 = oNode_infoscreen.selectSingleNode("txt_50").Text
        infoscreen_txt_051 = oNode_infoscreen.selectSingleNode("txt_51").Text
        infoscreen_txt_052 = oNode_infoscreen.selectSingleNode("txt_52").Text
        infoscreen_txt_053 = oNode_infoscreen.selectSingleNode("txt_53").Text
        infoscreen_txt_054 = oNode_infoscreen.selectSingleNode("txt_54").Text
        infoscreen_txt_055 = oNode_infoscreen.selectSingleNode("txt_55").Text
  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>