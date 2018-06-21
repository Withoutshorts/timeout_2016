
<% 
Dim objXMLHTTP_smiley, objXMLDOM_smiley, i_smiley, strHTML_smiley

Set objXMLDom_smiley = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_smiley = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_smiley.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/smileystatus.xml", False
'objXmlHttp_smiley.open "GET", "http://localhost/inc/xml/smileystatus.xml", False
'objXmlHttp_smiley.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/smileystatus.xml", False
'objXmlHttp_smiley.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/smileystatus.xml", False
'objXmlHttp_smiley.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/smileystatus.xml", False
objXmlHttp_smiley.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/smileystatus.xml", False

objXmlHttp_smiley.send 


Set objXmlDom_smiley = objXmlHttp_smiley.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_smiley = Nothing



Dim Address_smiley, Latitude_smiley, Longitude_smiley
Dim oNode_smiley, oNodes_smiley
Dim sXPathQuery_smiley

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
sXPathQuery_smiley = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_smiley = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_smiley = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_smiley = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_smiley = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_smiley = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_smiley = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_smiley = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_smiley = objXmlDom_smiley.documentElement.selectSingleNode(sXPathQuery_smiley)
Address_smiley = oNode_smiley.Text

Set oNodes_smiley = objXmlDom_smiley.documentElement.selectNodes(sXPathQuery_smiley)

    For Each oNode_smiley in oNodes_smiley

        smiley_txt_001 = oNode_smiley.selectSingleNode("txt_1").Text
        smiley_txt_002 = oNode_smiley.selectSingleNode("txt_2").Text
        smiley_txt_003 = oNode_smiley.selectSingleNode("txt_3").Text
        smiley_txt_004 = oNode_smiley.selectSingleNode("txt_4").Text
        smiley_txt_005 = oNode_smiley.selectSingleNode("txt_5").Text
        smiley_txt_006 = oNode_smiley.selectSingleNode("txt_6").Text
        smiley_txt_007 = oNode_smiley.selectSingleNode("txt_7").Text
        smiley_txt_008 = oNode_smiley.selectSingleNode("txt_8").Text
        smiley_txt_009 = oNode_smiley.selectSingleNode("txt_9").Text
        smiley_txt_010 = oNode_smiley.selectSingleNode("txt_10").Text

        



  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>