
<% 
Dim objXMLHTTP_fomr, objXMLDOM_fomr, i_fomr, strHTML_fomr

Set objXMLDom_fomr = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_fomr = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_fomr.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/fomr_sprog.xml", False
'objXmlHttp_fomr.open "GET", "http://localhost/inc/xml/fomr_sprog.xml", False
'objXmlHttp_fomr.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/fomr_sprog.xml", False
'objXmlHttp_fomr.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/fomr_sprog.xml", False
'objXmlHttp_fomr.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/fomr_sprog.xml", False
'objXmlHttp_fomr.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/fomr_sprog.xml", False
objXmlHttp_fomr.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/fomr_sprog.xml", False

objXmlHttp_fomr.send 


Set objXmlDom_fomr = objXmlHttp_fomr.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")awd


Set objXmlHttp_fomr = Nothing



Dim Address_fomr, Latitude_fomr, Longitude_fomr
Dim oNode_fomr, oNodes_fomr
Dim sXPathQuery_fomr

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
sXPathQuery_fomr = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_fomr = "//sprog/uk"
if lto = "a27" then
    sXPathQuery_fomr = "//sprog/trainerlog_uk"
end if
'Session.LCID = 2057
case 3
sXPathQuery_fomr = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_fomr = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_fomr = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_fomr = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_fomr = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_fomr = "//sprog/dk"
'Session.LCID = 1030
end select


'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_fomr = objXmlDom_fomr.documentElement.selectSingleNode(sXPathQuery_fomr)
Address_fomr = oNode_fomr.Text

Set oNodes_fomr = objXmlDom_fomr.documentElement.selectNodes(sXPathQuery_fomr)

    For Each oNode_fomr in oNodes_fomr

        fomr_txt_001 = oNode_fomr.selectSingleNode("txt_1").Text
        fomr_txt_002 = oNode_fomr.selectSingleNode("txt_2").Text
        fomr_txt_003 = oNode_fomr.selectSingleNode("txt_3").Text
        fomr_txt_003 = oNode_fomr.selectSingleNode("txt_3").Text
        fomr_txt_004 = oNode_fomr.selectSingleNode("txt_4").Text
        fomr_txt_005 = oNode_fomr.selectSingleNode("txt_5").Text

        fomr_txt_006 = oNode_fomr.selectSingleNode("txt_6").Text
        fomr_txt_007 = oNode_fomr.selectSingleNode("txt_7").Text
        fomr_txt_008 = oNode_fomr.selectSingleNode("txt_8").Text
        fomr_txt_009 = oNode_fomr.selectSingleNode("txt_9").Text
        fomr_txt_010 = oNode_fomr.selectSingleNode("txt_10").Text
        fomr_txt_011 = oNode_fomr.selectSingleNode("txt_11").Text
    
        fomr_txt_012 = oNode_fomr.selectSingleNode("txt_12").Text
        fomr_txt_013 = oNode_fomr.selectSingleNode("txt_13").Text
        fomr_txt_014 = oNode_fomr.selectSingleNode("txt_14").Text
        fomr_txt_015 = oNode_fomr.selectSingleNode("txt_15").Text
        fomr_txt_016 = oNode_fomr.selectSingleNode("txt_16").Text
        fomr_txt_017 = oNode_fomr.selectSingleNode("txt_17").Text
    
        fomr_txt_018 = oNode_fomr.selectSingleNode("txt_18").Text
        fomr_txt_019 = oNode_fomr.selectSingleNode("txt_19").Text
        fomr_txt_020 = oNode_fomr.selectSingleNode("txt_20").Text
        fomr_txt_021 = oNode_fomr.selectSingleNode("txt_21").Text
        fomr_txt_022 = oNode_fomr.selectSingleNode("txt_22").Text
        fomr_txt_023 = oNode_fomr.selectSingleNode("txt_23").Text
        fomr_txt_024 = oNode_fomr.selectSingleNode("txt_24").Text
    
        fomr_txt_025 = oNode_fomr.selectSingleNode("txt_25").Text
        fomr_txt_026 = oNode_fomr.selectSingleNode("txt_26").Text
        fomr_txt_027 = oNode_fomr.selectSingleNode("txt_27").Text
        fomr_txt_028 = oNode_fomr.selectSingleNode("txt_28").Text
        fomr_txt_029 = oNode_fomr.selectSingleNode("txt_29").Text
        fomr_txt_030 = oNode_fomr.selectSingleNode("txt_30").Text
        fomr_txt_031 = oNode_fomr.selectSingleNode("txt_31").Text
        fomr_txt_032 = oNode_fomr.selectSingleNode("txt_32").Text
        fomr_txt_033 = oNode_fomr.selectSingleNode("txt_33").Text
        fomr_txt_034 = oNode_fomr.selectSingleNode("txt_34").Text
    
        fomr_txt_035 = oNode_fomr.selectSingleNode("txt_35").Text
        fomr_txt_036 = oNode_fomr.selectSingleNode("txt_36").Text
        fomr_txt_037 = oNode_fomr.selectSingleNode("txt_37").Text
       
        fomr_txt_038 = oNode_fomr.selectSingleNode("txt_38").Text
        fomr_txt_039 = oNode_fomr.selectSingleNode("txt_39").Text
        fomr_txt_040 = oNode_fomr.selectSingleNode("txt_40").Text
        fomr_txt_041 = oNode_fomr.selectSingleNode("txt_41").Text
        fomr_txt_042 = oNode_fomr.selectSingleNode("txt_42").Text
        fomr_txt_043 = oNode_fomr.selectSingleNode("txt_43").Text
        fomr_txt_044 = oNode_fomr.selectSingleNode("txt_44").Text
        fomr_txt_045 = oNode_fomr.selectSingleNode("txt_45").Text
        fomr_txt_046 = oNode_fomr.selectSingleNode("txt_46").Text
        fomr_txt_047 = oNode_fomr.selectSingleNode("txt_47").Text
        fomr_txt_048 = oNode_fomr.selectSingleNode("txt_48").Text
         
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>