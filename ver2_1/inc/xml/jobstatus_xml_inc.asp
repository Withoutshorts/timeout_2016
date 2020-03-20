
<%




Dim objXMLHTTP_jobstatus, objXMLDOM_jobstatus, i_jobstatus, strHTML_jobstatus
Dim Address_jobstatus, Latitude_jobstatus, Longitude_jobstatus
Dim oNode_jobstatus, oNodes_jobstatus
Dim sXPathQuery_jobstatus



Set objXMLDom_jobstatus = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_jobstatus = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_jobstatus.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/jobstatus_sprog.xml", False
'objXmlHttp_jobstatus.open "GET", "http://localhost/inc/xml/jobstatus_sprog.xml", False
objXmlHttp_jobstatus.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/jobstatus_sprog.xml", False
'objXmlHttp_jobstatus.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/jobstatus_sprog.xml", False
'objXmlHttp_jobstatus.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/jobstatus_sprog.xml", False
'objXmlHttp_jobstatus.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/jobstatus_sprog.xml", False
objXmlHttp_jobstatus.send


Set objXmlDom_jobstatus = objXmlHttp_jobstatus.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_jobstatus = Nothing





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
sXPathQuery_jobstatus = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_jobstatus = "//sprog/uk"
'Session.LCID = 1033
case 3
sXPathQuery_jobstatus = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_jobstatus = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_jobstatus = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_jobstatus = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_jobstatus = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_jobstatus = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030


Set oNode_jobstatus = objXmlDom_jobstatus.documentElement.selectSingleNode(sXPathQuery_jobstatus)
Address_jobstatus = oNode_jobstatus.Text

Set oNodes_jobstatus = objXmlDom_jobstatus.documentElement.selectNodes(sXPathQuery_jobstatus)

    For Each oNode_jobstatus in oNodes_jobstatus
          
        jobstatus_txt_001 = oNode_jobstatus.selectSingleNode("txt_1").Text
        jobstatus_txt_002 = oNode_jobstatus.selectSingleNode("txt_2").Text

        if lto <> "dencker" then
            jobstatus_txt_003 = oNode_jobstatus.selectSingleNode("txt_3").Text
        else
            jobstatus_txt_003 = oNode_jobstatus.selectSingleNode("txt_3_dencker").Text
        end if

        jobstatus_txt_004 = oNode_jobstatus.selectSingleNode("txt_4").Text
        jobstatus_txt_005 = oNode_jobstatus.selectSingleNode("txt_5").Text
        jobstatus_txt_006 = oNode_jobstatus.selectSingleNode("txt_6").Text
        jobstatus_txt_007 = oNode_jobstatus.selectSingleNode("txt_7").Text
        jobstatus_txt_008 = oNode_jobstatus.selectSingleNode("txt_8").Text
        jobstatus_txt_009 = oNode_jobstatus.selectSingleNode("txt_9").Text
        jobstatus_txt_010 = oNode_jobstatus.selectSingleNode("txt_10").Text
        jobstatus_txt_011 = oNode_jobstatus.selectSingleNode("txt_11").Text
        jobstatus_txt_012 = oNode_jobstatus.selectSingleNode("txt_12").Text

    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>









