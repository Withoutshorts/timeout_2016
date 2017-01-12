
<% 
Dim objXMLHTTP_ttw, objXMLDOM_ttw, i_ttw, strHTML_ttw

Set objXMLDom_ttw = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_ttw = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_ttw.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/timetag_web.xml", False
'objXmlHttp_ttw.open "GET", "http://localhost/inc/xml/timetag_web.xml", False
'objXmlHttp_ttw.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/timetag_web.xml", False
'objXmlHttp_ttw.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/timetag_web.xml", False
'objXmlHttp_ttw.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/timetag_web.xml", False
objXmlHttp_ttw.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/timetag_web.xml", False

objXmlHttp_ttw.send


Set objXmlDom_ttw = objXmlHttp_ttw.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_ttw = Nothing



Dim Address_ttw, Latitude_ttw, Longitude_ttw
Dim oNode_ttw, oNodes_ttw
Dim sXPathQuery_ttw

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
sXPathQuery_ttw = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_ttw = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_ttw = "//sprog/se"
Session.LCID = 1053
case 4
sXPathQuery_ttw = "//sprog/no"
Session.LCID = 2068
case 5
sXPathQuery_ttw = "//sprog/es"
Session.LCID = 1034
case 6
sXPathQuery_ttw = "//sprog/de"
Session.LCID = 1031
case 7
sXPathQuery_ttw = "//sprog/fr"
Session.LCID = 1036
case else
sXPathQuery_ttw = "//sprog/dk"
Session.LCID = 1030
end select





Set oNode_ttw = objXmlDom_ttw.documentElement.selectSingleNode(sXPathQuery_ttw)
Address_ttw = oNode_ttw.Text

Set oNodes_ttw = objXmlDom_ttw.documentElement.selectNodes(sXPathQuery_ttw)

    For Each oNode_ttw in oNodes_ttw


                  
          ttw_txt_001 = oNode_ttw.selectSingleNode("txt_1").Text
          ttw_txt_002 = oNode_ttw.selectSingleNode("txt_2").Text
          ttw_txt_003 = oNode_ttw.selectSingleNode("txt_3").Text
          ttw_txt_004 = oNode_ttw.selectSingleNode("txt_4").Text
          ttw_txt_005 = oNode_ttw.selectSingleNode("txt_5").Text
          ttw_txt_006 = oNode_ttw.selectSingleNode("txt_6").Text
          ttw_txt_007 = oNode_ttw.selectSingleNode("txt_7").Text
          ttw_txt_008 = oNode_ttw.selectSingleNode("txt_8").Text
          ttw_txt_009 = oNode_ttw.selectSingleNode("txt_9").Text
          ttw_txt_010 = oNode_ttw.selectSingleNode("txt_10").Text
          ttw_txt_011 = oNode_ttw.selectSingleNode("txt_11").Text
          ttw_txt_012 = oNode_ttw.selectSingleNode("txt_12").Text
          ttw_txt_013 = oNode_ttw.selectSingleNode("txt_13").Text
          ttw_txt_014 = oNode_ttw.selectSingleNode("txt_14").Text
          ttw_txt_015 = oNode_ttw.selectSingleNode("txt_15").Text
          ttw_txt_016 = oNode_ttw.selectSingleNode("txt_16").Text
          ttw_txt_017 = oNode_ttw.selectSingleNode("txt_17").Text
          ttw_txt_018 = oNode_ttw.selectSingleNode("txt_18").Text
          ttw_txt_019 = oNode_ttw.selectSingleNode("txt_19").Text
          ttw_txt_020 = oNode_ttw.selectSingleNode("txt_20").Text
          ttw_txt_021 = oNode_ttw.selectSingleNode("txt_21").Text
          ttw_txt_022 = oNode_ttw.selectSingleNode("txt_22").Text
          ttw_txt_023 = oNode_ttw.selectSingleNode("txt_23").Text
          ttw_txt_024 = oNode_ttw.selectSingleNode("txt_24").Text
          ttw_txt_025 = oNode_ttw.selectSingleNode("txt_25").Text

  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>