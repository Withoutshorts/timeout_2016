
<% 

Dim objXMLHTTPTraveldp, objXMLDOMTraveldp, iTraveldp, strHTMLTraveldp

Set objXMLDomTraveldp = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttpTraveldp = Server.CreateObject("Msxml2.ServerXMLHTTP")
objXmlHttpTraveldp.open "GET", "http://localhost/inc/xml/tsa_sprog_traveldp.xml", False
'objXmlHttpTraveldp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_1/inc/xml/tsa_sprog.xml", False
'objXmlHttpTraveldp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver3_99/inc/xml/tsa_sprog.xml", False
'objXmlHttpTraveldp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/tsa_sprog.xml", False
'objXmlHttpTraveldp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver4_22/inc/xml/tsa_sprog.xml", False

objXmlHttpTraveldp.send


Set objXmlDomTraveldp = objXmlHttpTraveldp.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttpTraveldp = Nothing



Dim AddressTraveldp
Dim oNodeTraveldp, oNodesTraveldp
Dim sXPathQueryTraveldp


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
sXPathQueryTraveldp = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQueryTraveldp = "//sprog/uk"
Session.LCID = 1033
case 3
sXPathQueryTraveldp = "//sprog/se"
Session.LCID = 1053
case 4
sXPathQueryTraveldp = "//sprog/no"
Session.LCID = 1044
case 5
sXPathQueryTraveldp = "//sprog/es"
Session.LCID = 1036
case 6
sXPathQueryTraveldp = "//sprog/de"
Session.LCID = 1053
case 7
sXPathQueryTraveldp = "//sprog/fr"
Session.LCID = 1033
case else
sXPathQueryTraveldp = "//sprog/dk"
Session.LCID = 1030
end select

'Response.Write "Session.LCID" &  Session.LCID

Set oNodeTraveldp = objXmlDomTraveldp.documentElement.selectSingleNode(sXPathQueryTraveldp)
AddressTraveldp = oNodeTraveldp.Text

Set oNodesTraveldp = objXmlDomTraveldp.documentElement.selectNodes(sXPathQueryTraveldp)

For Each oNodeTraveldp in oNodesTraveldp

    
    tsa_txt_traveldp_001 = oNodeTraveldp.selectSingleNode("txt_1").Text
    tsa_txt_traveldp_002 = oNodeTraveldp.selectSingleNode("txt_2").Text
    tsa_txt_traveldp_003 = oNodeTraveldp.selectSingleNode("txt_3").Text
    tsa_txt_traveldp_004 = oNodeTraveldp.selectSingleNode("txt_4").Text
    tsa_txt_traveldp_005 = oNodeTraveldp.selectSingleNode("txt_5").Text
    tsa_txt_traveldp_006 = oNodeTraveldp.selectSingleNode("txt_6").Text
    tsa_txt_traveldp_007 = oNodeTraveldp.selectSingleNode("txt_7").Text
    tsa_txt_traveldp_008 = oNodeTraveldp.selectSingleNode("txt_8").Text
    tsa_txt_traveldp_009 = oNodeTraveldp.selectSingleNode("txt_9").Text
    tsa_txt_traveldp_010 = oNodeTraveldp.selectSingleNode("txt_10").Text
    tsa_txt_traveldp_011 = oNodeTraveldp.selectSingleNode("txt_11").Text
    tsa_txt_traveldp_012 = oNodeTraveldp.selectSingleNode("txt_12").Text
    tsa_txt_traveldp_013 = oNodeTraveldp.selectSingleNode("txt_13").Text
    tsa_txt_traveldp_014 = oNodeTraveldp.selectSingleNode("txt_14").Text
    tsa_txt_traveldp_015 = oNodeTraveldp.selectSingleNode("txt_15").Text
    tsa_txt_traveldp_016 = oNodeTraveldp.selectSingleNode("txt_16").Text
    tsa_txt_traveldp_017 = oNodeTraveldp.selectSingleNode("txt_17").Text
    tsa_txt_traveldp_018 = oNodeTraveldp.selectSingleNode("txt_18").Text
    tsa_txt_traveldp_019 = oNodeTraveldp.selectSingleNode("txt_19").Text
    tsa_txt_traveldp_020 = oNodeTraveldp.selectSingleNode("txt_20").Text
    tsa_txt_traveldp_021 = oNodeTraveldp.selectSingleNode("txt_21").Text
    tsa_txt_traveldp_022 = oNodeTraveldp.selectSingleNode("txt_22").Text
    tsa_txt_traveldp_023 = oNodeTraveldp.selectSingleNode("txt_23").Text
    tsa_txt_traveldp_024 = oNodeTraveldp.selectSingleNode("txt_24").Text
    tsa_txt_traveldp_025 = oNodeTraveldp.selectSingleNode("txt_25").Text
    tsa_txt_traveldp_026 = oNodeTraveldp.selectSingleNode("txt_26").Text
    tsa_txt_traveldp_027 = oNodeTraveldp.selectSingleNode("txt_27").Text
    tsa_txt_traveldp_028 = oNodeTraveldp.selectSingleNode("txt_28").Text
    tsa_txt_traveldp_029 = oNodeTraveldp.selectSingleNode("txt_29").Text
    tsa_txt_traveldp_030 = oNodeTraveldp.selectSingleNode("txt_30").Text
    tsa_txt_traveldp_031 = oNodeTraveldp.selectSingleNode("txt_31").Text
    tsa_txt_traveldp_032 = oNodeTraveldp.selectSingleNode("txt_32").Text
    tsa_txt_traveldp_033 = oNodeTraveldp.selectSingleNode("txt_33").Text
    tsa_txt_traveldp_034 = oNodeTraveldp.selectSingleNode("txt_34").Text
    tsa_txt_traveldp_035 = oNodeTraveldp.selectSingleNode("txt_35").Text
    tsa_txt_traveldp_036 = oNodeTraveldp.selectSingleNode("txt_36").Text
    tsa_txt_traveldp_037 = oNodeTraveldp.selectSingleNode("txt_37").Text
    tsa_txt_traveldp_038 = oNodeTraveldp.selectSingleNode("txt_38").Text
    tsa_txt_traveldp_039 = oNodeTraveldp.selectSingleNode("txt_39").Text
    tsa_txt_traveldp_040 = oNodeTraveldp.selectSingleNode("txt_40").Text
    
    
next




%>