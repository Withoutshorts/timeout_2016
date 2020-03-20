
<% 
Dim objXMLHTTP_forecastkap, objXMLDOM_forecastkap, i_forecastkap, strHTML_forecastkap

Set objXMLDom_forecastkap = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_forecastkap = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_forecastkap.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/forecast_kapacitet.xml", False
'objXmlHttp_forecastkap.open "GET", "http://localhost/inc/xml/forecast_kapacitet.xml", False
'objXmlHttp_forecastkap.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/forecast_kapacitet.xml", False
'objXmlHttp_forecastkap.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/forecast_kapacitet.xml", False
'objXmlHttp_forecastkap.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/forecast_kapacitet.xml", False
'objXmlHttp_forecastkap.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/forecast_kapacitet.xml", False
objXmlHttp_forecastkap.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/forecast_kapacitet.xml", False

objXmlHttp_forecastkap.send 


Set objXmlDom_forecastkap = objXmlHttp_forecastkap.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_forecastkap = Nothing



Dim Address_forecastkap, Latitude_forecastkap, Longitude_forecastkap
Dim oNode_forecastkap, oNodes_forecastkap
Dim sXPathQuery_forecastkap

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
sXPathQuery_forecastkap = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_forecastkap = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_forecastkap = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_forecastkap = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_forecastkap = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_forecastkap = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_forecastkap = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_forecastkap = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_forecastkap = objXmlDom_forecastkap.documentElement.selectSingleNode(sXPathQuery_forecastkap)
Address_forecastkap = oNode_forecastkap.Text

Set oNodes_forecastkap = objXmlDom_forecastkap.documentElement.selectNodes(sXPathQuery_forecastkap)

    For Each oNode_forecastkap in oNodes_forecastkap

        forecastkap_txt_001 = oNode_forecastkap.selectSingleNode("txt_1").Text
        forecastkap_txt_002 = oNode_forecastkap.selectSingleNode("txt_2").Text
        forecastkap_txt_003 = oNode_forecastkap.selectSingleNode("txt_3").Text
        forecastkap_txt_004 = oNode_forecastkap.selectSingleNode("txt_4").Text
        forecastkap_txt_005 = oNode_forecastkap.selectSingleNode("txt_5").Text
        forecastkap_txt_006 = oNode_forecastkap.selectSingleNode("txt_6").Text
        forecastkap_txt_007 = oNode_forecastkap.selectSingleNode("txt_7").Text
        forecastkap_txt_008 = oNode_forecastkap.selectSingleNode("txt_8").Text
        forecastkap_txt_009 = oNode_forecastkap.selectSingleNode("txt_9").Text
        forecastkap_txt_010 = oNode_forecastkap.selectSingleNode("txt_10").Text

        forecastkap_txt_011 = oNode_forecastkap.selectSingleNode("txt_11").Text
  
    
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>