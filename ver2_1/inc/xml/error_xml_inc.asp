
<%




Dim objXMLHTTP_error, objXMLDOM_error, i_error, strHTML_error
Dim Address_error, Latitude_error, Longitude_error
Dim oNode_error, oNodes_error
Dim sXPathQuery_error



Set objXMLDom_error = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_error = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_error.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/error_sprog.xml", False
'objXmlHttp_error.open "GET", "http://localhost/inc/xml/error_sprog.xml", False
objXmlHttp_error.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/error_sprog.xml", False
'objXmlHttp_error.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/error_sprog.xml", False
'objXmlHttp_error.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/error_sprog.xml", False
objXmlHttp_error.send


Set objXmlDom_error = objXmlHttp_error.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_error = Nothing





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
sXPathQuery_error = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_error = "//sprog/uk"
'Session.LCID = 1033
case 3
sXPathQuery_error = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_error = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_error = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_error = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_error = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_error = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle bel�b og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030


Set oNode_error = objXmlDom_error.documentElement.selectSingleNode(sXPathQuery_error)
Address_error = oNode_error.Text

Set oNodes_error = objXmlDom_error.documentElement.selectNodes(sXPathQuery_error)

    For Each oNode_error in oNodes_error
          
          err_txt_001 = oNode_error.selectSingleNode("txt_001").Text
          err_txt_004 = oNode_error.selectSingleNode("txt_004").Text
          err_txt_028 = oNode_error.selectSingleNode("txt_028").Text
          
          err_txt_074 = oNode_error.selectSingleNode("txt_074").Text
          
          err_txt_090 = oNode_error.selectSingleNode("txt_090").Text
          err_txt_106 = oNode_error.selectSingleNode("txt_106").Text
          err_txt_107 = oNode_error.selectSingleNode("txt_107").Text
          
          err_txt_108A = oNode_error.selectSingleNode("txt_108A").Text
          err_txt_108B = oNode_error.selectSingleNode("txt_108B").Text
          err_txt_108C = oNode_error.selectSingleNode("txt_108C").Text
          
          err_txt_120 = oNode_error.selectSingleNode("txt_120").Text
        

    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>