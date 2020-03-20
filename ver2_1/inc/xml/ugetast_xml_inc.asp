
<% 
Dim objXMLHTTP_ugetast, objXMLDOM_ugetast, i_ugetast, strHTML_ugetast

Set objXMLDom_ugetast = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_ugetast = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_ugetast.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/ugetast_sprog.xml", False
'objXmlHttp_ugetast.open "GET", "http://localhost/inc/xml/ugetast_sprog.xml", False
'objXmlHttp_ugetast.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/ugetast_sprog.xml", False
'objXmlHttp_ugetast.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/ugetast_sprog.xml", False
'objXmlHttp_ugetast.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/ugetast_sprog.xml", False
'objXmlHttp_ugetast.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/ugetast_sprog.xml", False
objXmlHttp_ugetast.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/ugetast_sprog.xml", False

objXmlHttp_ugetast.send


Set objXmlDom_ugetast = objXmlHttp_ugetast.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_ugetast = Nothing



Dim Address_ugetast, Latitude_ugetast, Longitude_ugetast
Dim oNode_ugetast, oNodes_ugetast
Dim sXPathQuery_ugetast

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
sXPathQuery_ugetast = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_ugetast = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_ugetast = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_ugetast = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_ugetast = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_ugetast = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_ugetast = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_ugetast = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_ugetast = objXmlDom_ugetast.documentElement.selectSingleNode(sXPathQuery_ugetast)
Address_ugetast = oNode_ugetast.Text

Set oNodes_ugetast = objXmlDom_ugetast.documentElement.selectNodes(sXPathQuery_ugetast)

    For Each oNode_ugetast in oNodes_ugetast

        ugetast_txt_001 = oNode_ugetast.selectSingleNode("txt_1").Text
        ugetast_txt_002 = oNode_ugetast.selectSingleNode("txt_2").Text
        ugetast_txt_003 = oNode_ugetast.selectSingleNode("txt_3").Text
        ugetast_txt_003 = oNode_ugetast.selectSingleNode("txt_3").Text
        ugetast_txt_004 = oNode_ugetast.selectSingleNode("txt_4").Text
        ugetast_txt_005 = oNode_ugetast.selectSingleNode("txt_5").Text

        ugetast_txt_006 = oNode_ugetast.selectSingleNode("txt_6").Text
        ugetast_txt_007 = oNode_ugetast.selectSingleNode("txt_7").Text
        ugetast_txt_008 = oNode_ugetast.selectSingleNode("txt_8").Text
        ugetast_txt_009 = oNode_ugetast.selectSingleNode("txt_9").Text
        ugetast_txt_010 = oNode_ugetast.selectSingleNode("txt_10").Text
        ugetast_txt_011 = oNode_ugetast.selectSingleNode("txt_11").Text
    
        ugetast_txt_012 = oNode_ugetast.selectSingleNode("txt_12").Text
        ugetast_txt_013 = oNode_ugetast.selectSingleNode("txt_13").Text
        ugetast_txt_014 = oNode_ugetast.selectSingleNode("txt_14").Text
        ugetast_txt_015 = oNode_ugetast.selectSingleNode("txt_15").Text
        ugetast_txt_016 = oNode_ugetast.selectSingleNode("txt_16").Text
        ugetast_txt_017 = oNode_ugetast.selectSingleNode("txt_17").Text
    
        ugetast_txt_018 = oNode_ugetast.selectSingleNode("txt_18").Text
        ugetast_txt_019 = oNode_ugetast.selectSingleNode("txt_19").Text
        ugetast_txt_020 = oNode_ugetast.selectSingleNode("txt_20").Text
        ugetast_txt_021 = oNode_ugetast.selectSingleNode("txt_21").Text

        ugetast_txt_022 = oNode_ugetast.selectSingleNode("txt_22").Text
        ugetast_txt_023 = oNode_ugetast.selectSingleNode("txt_23").Text
        ugetast_txt_024 = oNode_ugetast.selectSingleNode("txt_24").Text
        ugetast_txt_025 = oNode_ugetast.selectSingleNode("txt_25").Text
        ugetast_txt_026 = oNode_ugetast.selectSingleNode("txt_26").Text
        ugetast_txt_027 = oNode_ugetast.selectSingleNode("txt_27").Text
        ugetast_txt_028 = oNode_ugetast.selectSingleNode("txt_28").Text

        ugetast_txt_029 = oNode_ugetast.selectSingleNode("txt_29").Text
        ugetast_txt_030 = oNode_ugetast.selectSingleNode("txt_30").Text
        ugetast_txt_031 = oNode_ugetast.selectSingleNode("txt_31").Text
        ugetast_txt_032 = oNode_ugetast.selectSingleNode("txt_32").Text
        ugetast_txt_033 = oNode_ugetast.selectSingleNode("txt_33").Text
        ugetast_txt_034 = oNode_ugetast.selectSingleNode("txt_34").Text
        ugetast_txt_035 = oNode_ugetast.selectSingleNode("txt_35").Text
        ugetast_txt_036 = oNode_ugetast.selectSingleNode("txt_36").Text
        ugetast_txt_037 = oNode_ugetast.selectSingleNode("txt_37").Text
        ugetast_txt_038 = oNode_ugetast.selectSingleNode("txt_38").Text

        ugetast_txt_039 = oNode_ugetast.selectSingleNode("txt_39").Text
        ugetast_txt_040 = oNode_ugetast.selectSingleNode("txt_40").Text
        ugetast_txt_041 = oNode_ugetast.selectSingleNode("txt_41").Text
        ugetast_txt_042 = oNode_ugetast.selectSingleNode("txt_42").Text
        ugetast_txt_043 = oNode_ugetast.selectSingleNode("txt_43").Text

  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>