
<% 
Dim objXMLHTTP_feriekalender, objXMLDOM_feriekalender, i_feriekalender, strHTML_feriekalender

Set objXMLDom_feriekalender = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_feriekalender = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_feriekalender.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/feriekalender_sprog.xml", False
'objXmlHttp_feriekalender.open "GET", "http://localhost/inc/xml/feriekalender_sprog.xml", False
'objXmlHttp_feriekalender.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/feriekalender_sprog.xml", False
'objXmlHttp_feriekalender.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/feriekalender_sprog.xml", False
'objXmlHttp_feriekalender.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/feriekalender_sprog.xml", False
'objXmlHttp_feriekalender.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/feriekalender_sprog.xml", False
objXmlHttp_feriekalender.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/feriekalender_sprog.xml", False

objXmlHttp_feriekalender.send


Set objXmlDom_feriekalender = objXmlHttp_feriekalender.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_feriekalender = Nothing



Dim Address_feriekalender, Latitude_feriekalender, Longitude_feriekalender
Dim oNode_feriekalender, oNodes_feriekalender
Dim sXPathQuery_feriekalender

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
sXPathQuery_feriekalender = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_feriekalender = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_feriekalender = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_feriekalender = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_feriekalender = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_feriekalender = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_feriekalender = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_feriekalender = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_feriekalender = objXmlDom_feriekalender.documentElement.selectSingleNode(sXPathQuery_feriekalender)
Address_feriekalender = oNode_feriekalender.Text

Set oNodes_feriekalender = objXmlDom_feriekalender.documentElement.selectNodes(sXPathQuery_feriekalender)

    For Each oNode_feriekalender in oNodes_feriekalender

        feriekalender_txt_001 = oNode_feriekalender.selectSingleNode("txt_1").Text
        feriekalender_txt_002 = oNode_feriekalender.selectSingleNode("txt_2").Text
        feriekalender_txt_003 = oNode_feriekalender.selectSingleNode("txt_3").Text
        feriekalender_txt_003 = oNode_feriekalender.selectSingleNode("txt_3").Text
        feriekalender_txt_004 = oNode_feriekalender.selectSingleNode("txt_4").Text
        feriekalender_txt_005 = oNode_feriekalender.selectSingleNode("txt_5").Text

        feriekalender_txt_006 = oNode_feriekalender.selectSingleNode("txt_6").Text
        feriekalender_txt_007 = oNode_feriekalender.selectSingleNode("txt_7").Text
        feriekalender_txt_008 = oNode_feriekalender.selectSingleNode("txt_8").Text
        feriekalender_txt_009 = oNode_feriekalender.selectSingleNode("txt_9").Text
        feriekalender_txt_010 = oNode_feriekalender.selectSingleNode("txt_10").Text
        feriekalender_txt_011 = oNode_feriekalender.selectSingleNode("txt_11").Text
    
        feriekalender_txt_012 = oNode_feriekalender.selectSingleNode("txt_12").Text
        feriekalender_txt_013 = oNode_feriekalender.selectSingleNode("txt_13").Text
        feriekalender_txt_014 = oNode_feriekalender.selectSingleNode("txt_14").Text
        feriekalender_txt_015 = oNode_feriekalender.selectSingleNode("txt_15").Text
        feriekalender_txt_016 = oNode_feriekalender.selectSingleNode("txt_16").Text
        feriekalender_txt_017 = oNode_feriekalender.selectSingleNode("txt_17").Text
    
        feriekalender_txt_018 = oNode_feriekalender.selectSingleNode("txt_18").Text
        feriekalender_txt_019 = oNode_feriekalender.selectSingleNode("txt_19").Text
        feriekalender_txt_020 = oNode_feriekalender.selectSingleNode("txt_20").Text
        feriekalender_txt_021 = oNode_feriekalender.selectSingleNode("txt_21").Text

        feriekalender_txt_022 = oNode_feriekalender.selectSingleNode("txt_22").Text
        feriekalender_txt_023 = oNode_feriekalender.selectSingleNode("txt_23").Text
        feriekalender_txt_024 = oNode_feriekalender.selectSingleNode("txt_24").Text
        feriekalender_txt_025 = oNode_feriekalender.selectSingleNode("txt_25").Text
        feriekalender_txt_026 = oNode_feriekalender.selectSingleNode("txt_26").Text
        feriekalender_txt_027 = oNode_feriekalender.selectSingleNode("txt_27").Text
        feriekalender_txt_028 = oNode_feriekalender.selectSingleNode("txt_28").Text

        feriekalender_txt_029 = oNode_feriekalender.selectSingleNode("txt_29").Text
        feriekalender_txt_030 = oNode_feriekalender.selectSingleNode("txt_30").Text
        feriekalender_txt_031 = oNode_feriekalender.selectSingleNode("txt_31").Text
        feriekalender_txt_032 = oNode_feriekalender.selectSingleNode("txt_32").Text
        feriekalender_txt_033 = oNode_feriekalender.selectSingleNode("txt_33").Text
        feriekalender_txt_034 = oNode_feriekalender.selectSingleNode("txt_34").Text
        feriekalender_txt_035 = oNode_feriekalender.selectSingleNode("txt_35").Text
        feriekalender_txt_036 = oNode_feriekalender.selectSingleNode("txt_36").Text
        feriekalender_txt_037 = oNode_feriekalender.selectSingleNode("txt_37").Text
        feriekalender_txt_038 = oNode_feriekalender.selectSingleNode("txt_38").Text
        feriekalender_txt_039 = oNode_feriekalender.selectSingleNode("txt_39").Text
        feriekalender_txt_040 = oNode_feriekalender.selectSingleNode("txt_40").Text
        feriekalender_txt_041 = oNode_feriekalender.selectSingleNode("txt_41").Text
        feriekalender_txt_042 = oNode_feriekalender.selectSingleNode("txt_42").Text
        feriekalender_txt_043 = oNode_feriekalender.selectSingleNode("txt_43").Text
        feriekalender_txt_044 = oNode_feriekalender.selectSingleNode("txt_44").Text
        feriekalender_txt_045 = oNode_feriekalender.selectSingleNode("txt_45").Text
        feriekalender_txt_046 = oNode_feriekalender.selectSingleNode("txt_46").Text
        feriekalender_txt_047 = oNode_feriekalender.selectSingleNode("txt_47").Text
        feriekalender_txt_048 = oNode_feriekalender.selectSingleNode("txt_48").Text
        feriekalender_txt_049 = oNode_feriekalender.selectSingleNode("txt_49").Text
        feriekalender_txt_050 = oNode_feriekalender.selectSingleNode("txt_50").Text
  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>