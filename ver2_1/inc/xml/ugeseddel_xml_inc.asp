
<% 

Dim objXMLHTTP_uge, objXMLDOM_uge, i_uge, strHTML_uge

Set objXMLDom = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
objXmlHttp.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/ugeseddel_sprog.xml", False
'objXmlHttp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_1/inc/xml/ugeseddel_sprog.xml", False
'objXmlHttp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver3_99/inc/xml/ugeseddel_sprog.xml", False
'objXmlHttp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/ugeseddel_sprog.xml", False
'objXmlHttp.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver4_22/inc/xml/ugeseddel_sprog.xml", False

objXmlHttp.send


Set objXmlDom = objXmlHttp.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp = Nothing



Dim Address_uge, Latitude_uge, Longitude_uge
Dim oNode_uge, oNodes_uge
Dim sXPathQuery_uge


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
sXPathQuery_uge = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_uge = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_uge = "//sprog/se"
Session.LCID = 1053
case else
sXPathQuery_uge = "//sprog/dk"
Session.LCID = 1030
end select

'Response.Write "Session.LCID" &  Session.LCID

Set oNode_uge = objXmlDom.documentElement.selectSingleNode(sXPathQuery_uge)
Address_uge = oNode_uge.Text

Set oNodes_uge = objXmlDom.documentElement.selectNodes(sXPathQuery_uge)

For Each oNode_uge in oNodes_uge

    uge_txt_001 = oNode_uge.selectSingleNode("txt_1").Text
    uge_txt_002 = oNode_uge.selectSingleNode("txt_2").Text
    uge_txt_003 = oNode_uge.selectSingleNode("txt_3").Text
    uge_txt_003 = oNode_uge.selectSingleNode("txt_3").Text
    uge_txt_004 = oNode_uge.selectSingleNode("txt_4").Text
    uge_txt_005 = oNode_uge.selectSingleNode("txt_5").Text
    uge_txt_006 = oNode_uge.selectSingleNode("txt_6").Text
    uge_txt_007 = oNode_uge.selectSingleNode("txt_7").Text
    uge_txt_008 = oNode_uge.selectSingleNode("txt_8").Text
    uge_txt_009 = oNode_uge.selectSingleNode("txt_9").Text
    uge_txt_010 = oNode_uge.selectSingleNode("txt_10").Text
    uge_txt_011 = oNode_uge.selectSingleNode("txt_11").Text
    uge_txt_012 = oNode_uge.selectSingleNode("txt_12").Text
    uge_txt_013 = oNode_uge.selectSingleNode("txt_13").Text
    uge_txt_014 = oNode_uge.selectSingleNode("txt_14").Text
    uge_txt_015 = oNode_uge.selectSingleNode("txt_15").Text
    uge_txt_016 = oNode_uge.selectSingleNode("txt_16").Text
    uge_txt_017 = oNode_uge.selectSingleNode("txt_17").Text
    uge_txt_018 = oNode_uge.selectSingleNode("txt_18").Text
    uge_txt_019 = oNode_uge.selectSingleNode("txt_19").Text
    uge_txt_020 = oNode_uge.selectSingleNode("txt_20").Text
    uge_txt_021 = oNode_uge.selectSingleNode("txt_21").Text
    uge_txt_022 = oNode_uge.selectSingleNode("txt_22").Text
    uge_txt_023 = oNode_uge.selectSingleNode("txt_23").Text
    uge_txt_024 = oNode_uge.selectSingleNode("txt_24").Text 
    uge_txt_025 = oNode_uge.selectSingleNode("txt_25").Text
    uge_txt_026 = oNode_uge.selectSingleNode("txt_26").Text
    uge_txt_027 = oNode_uge.selectSingleNode("txt_27").Text
    uge_txt_028 = oNode_uge.selectSingleNode("txt_28").Text
    uge_txt_029 = oNode_uge.selectSingleNode("txt_29").Text
    uge_txt_030 = oNode_uge.selectSingleNode("txt_30").Text
    uge_txt_031 = oNode_uge.selectSingleNode("txt_31").Text
    uge_txt_032 = oNode_uge.selectSingleNode("txt_32").Text
    uge_txt_033 = oNode_uge.selectSingleNode("txt_33").Text
    uge_txt_034 = oNode_uge.selectSingleNode("txt_34").Text   
    uge_txt_035 = oNode_uge.selectSingleNode("txt_35").Text
    uge_txt_036 = oNode_uge.selectSingleNode("txt_36").Text
    uge_txt_037 = oNode_uge.selectSingleNode("txt_37").Text    
    uge_txt_038 = oNode_uge.selectSingleNode("txt_38").Text
    uge_txt_039 = oNode_uge.selectSingleNode("txt_39").Text
    uge_txt_040 = oNode_uge.selectSingleNode("txt_40").Text
    uge_txt_041 = oNode_uge.selectSingleNode("txt_41").Text
    uge_txt_042 = oNode_uge.selectSingleNode("txt_42").Text
    uge_txt_043 = oNode_uge.selectSingleNode("txt_43").Text
    uge_txt_044 = oNode_uge.selectSingleNode("txt_44").Text    
    uge_txt_045 = oNode_uge.selectSingleNode("txt_45").Text
    uge_txt_046 = oNode_uge.selectSingleNode("txt_46").Text
    uge_txt_047 = oNode_uge.selectSingleNode("txt_47").Text
    uge_txt_048 = oNode_uge.selectSingleNode("txt_48").Text
    uge_txt_049 = oNode_uge.selectSingleNode("txt_49").Text
    uge_txt_050 = oNode_uge.selectSingleNode("txt_50").Text
    uge_txt_051 = oNode_uge.selectSingleNode("txt_51").Text
    uge_txt_052 = oNode_uge.selectSingleNode("txt_52").Text
    uge_txt_053 = oNode_uge.selectSingleNode("txt_53").Text
    uge_txt_054 = oNode_uge.selectSingleNode("txt_54").Text
    uge_txt_055 = oNode_uge.selectSingleNode("txt_55").Text
    uge_txt_056 = oNode_uge.selectSingleNode("txt_56").Text
     

    
next




%>