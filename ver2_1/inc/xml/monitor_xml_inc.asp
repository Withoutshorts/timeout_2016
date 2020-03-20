
<% 
Dim objXMLHTTP_monitor, objXMLDOM_monitor, i_monitor, strHTML_monitor

Set objXMLDom_monitor = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_monitor = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_monitor.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/monitor_sprog.xml", False
'objXmlHttp_monitor.open "GET", "http://localhost/inc/xml/monitor_sprog.xml", False
'objXmlHttp_monitor.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/monitor_sprog.xml", False
'objXmlHttp_monitor.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/monitor_sprog.xml", False
objXmlHttp_monitor.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/monitor_sprog.xml", False
'objXmlHttp_monitor.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/monitor_sprog.xml", False
'objXmlHttp_monitor.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/monitor_sprog.xml", False

objXmlHttp_monitor.send


Set objXmlDom_monitor = objXmlHttp_monitor.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_monitor = Nothing



Dim Address_monitor, Latitude_monitor, Longitude_monitor
Dim oNode_monitor, oNodes_monitor
Dim sXPathQuery_monitor

sprog = 1 'DK
if len(trim(session("mid"))) <> 0 then
strSQL = "SELECT sprog FROM medarbejdere WHERE mid = " & session("mid")

oRec.open strSQL, oConn, 3
if not oRec.EOF then
sprog = oRec("sprog")
end if
oRec.close
end if

if lto = "cool" then
    sprog = 2
end if

select case sprog
case 1
sXPathQuery_monitor = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_monitor = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_monitor = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_monitor = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_monitor = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_monitor = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_monitor = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_monitor = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_monitor = objXmlDom_monitor.documentElement.selectSingleNode(sXPathQuery_monitor)
Address_monitor = oNode_monitor.Text

Set oNodes_monitor = objXmlDom_monitor.documentElement.selectNodes(sXPathQuery_monitor)

    For Each oNode_monitor in oNodes_monitor

        monitor_txt_001 = oNode_monitor.selectSingleNode("txt_1").Text
        monitor_txt_002 = oNode_monitor.selectSingleNode("txt_2").Text
        monitor_txt_003 = oNode_monitor.selectSingleNode("txt_3").Text
        monitor_txt_003 = oNode_monitor.selectSingleNode("txt_3").Text
        monitor_txt_004 = oNode_monitor.selectSingleNode("txt_4").Text
        monitor_txt_005 = oNode_monitor.selectSingleNode("txt_5").Text

        monitor_txt_006 = oNode_monitor.selectSingleNode("txt_6").Text
        monitor_txt_007 = oNode_monitor.selectSingleNode("txt_7").Text
        monitor_txt_008 = oNode_monitor.selectSingleNode("txt_8").Text
        monitor_txt_009 = oNode_monitor.selectSingleNode("txt_9").Text
        monitor_txt_010 = oNode_monitor.selectSingleNode("txt_10").Text
        monitor_txt_011 = oNode_monitor.selectSingleNode("txt_11").Text
    
        monitor_txt_012 = oNode_monitor.selectSingleNode("txt_12").Text
        monitor_txt_013 = oNode_monitor.selectSingleNode("txt_13").Text
        monitor_txt_014 = oNode_monitor.selectSingleNode("txt_14").Text
        monitor_txt_015 = oNode_monitor.selectSingleNode("txt_15").Text
        monitor_txt_016 = oNode_monitor.selectSingleNode("txt_16").Text
        monitor_txt_017 = oNode_monitor.selectSingleNode("txt_17").Text
    
        monitor_txt_018 = oNode_monitor.selectSingleNode("txt_18").Text
        monitor_txt_019 = oNode_monitor.selectSingleNode("txt_19").Text
        monitor_txt_020 = oNode_monitor.selectSingleNode("txt_20").Text
        monitor_txt_021 = oNode_monitor.selectSingleNode("txt_21").Text

        monitor_txt_022 = oNode_monitor.selectSingleNode("txt_22").Text
        monitor_txt_023 = oNode_monitor.selectSingleNode("txt_23").Text
        monitor_txt_024 = oNode_monitor.selectSingleNode("txt_24").Text
        monitor_txt_025 = oNode_monitor.selectSingleNode("txt_25").Text
        monitor_txt_026 = oNode_monitor.selectSingleNode("txt_26").Text
        monitor_txt_027 = oNode_monitor.selectSingleNode("txt_27").Text
        monitor_txt_028 = oNode_monitor.selectSingleNode("txt_28").Text

        monitor_txt_029 = oNode_monitor.selectSingleNode("txt_29").Text
        monitor_txt_030 = oNode_monitor.selectSingleNode("txt_30").Text
        monitor_txt_031 = oNode_monitor.selectSingleNode("txt_31").Text
        monitor_txt_032 = oNode_monitor.selectSingleNode("txt_32").Text
        monitor_txt_033 = oNode_monitor.selectSingleNode("txt_33").Text
        monitor_txt_034 = oNode_monitor.selectSingleNode("txt_34").Text
        monitor_txt_035 = oNode_monitor.selectSingleNode("txt_35").Text
        monitor_txt_036 = oNode_monitor.selectSingleNode("txt_36").Text
        monitor_txt_037 = oNode_monitor.selectSingleNode("txt_37").Text
        monitor_txt_038 = oNode_monitor.selectSingleNode("txt_38").Text
        monitor_txt_039 = oNode_monitor.selectSingleNode("txt_39").Text
        monitor_txt_040 = oNode_monitor.selectSingleNode("txt_40").Text
        monitor_txt_041 = oNode_monitor.selectSingleNode("txt_41").Text
        monitor_txt_042 = oNode_monitor.selectSingleNode("txt_42").Text
        monitor_txt_043 = oNode_monitor.selectSingleNode("txt_43").Text
        monitor_txt_044 = oNode_monitor.selectSingleNode("txt_44").Text
        monitor_txt_045 = oNode_monitor.selectSingleNode("txt_45").Text
        monitor_txt_046 = oNode_monitor.selectSingleNode("txt_46").Text
        monitor_txt_047 = oNode_monitor.selectSingleNode("txt_47").Text
        monitor_txt_048 = oNode_monitor.selectSingleNode("txt_48").Text
        monitor_txt_049 = oNode_monitor.selectSingleNode("txt_49").Text
        monitor_txt_050 = oNode_monitor.selectSingleNode("txt_50").Text
        monitor_txt_051 = oNode_monitor.selectSingleNode("txt_51").Text
        monitor_txt_052 = oNode_monitor.selectSingleNode("txt_52").Text
        monitor_txt_053 = oNode_monitor.selectSingleNode("txt_53").Text
        monitor_txt_054 = oNode_monitor.selectSingleNode("txt_54").Text
        monitor_txt_055 = oNode_monitor.selectSingleNode("txt_55").Text
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>