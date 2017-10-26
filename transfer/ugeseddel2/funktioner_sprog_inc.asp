
<% 
Dim objXMLHTTP_funk, objXMLDOM_funk, i_funk, strHTML_funk

Set objXMLDom_funk = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_funk = Server.CreateObject("Msxml2.ServerXMLHTTP")
objXmlHttp_funk.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/funktioner_sprog.xml", False
'objXmlHttp_funk.open "GET", "http://localhost/inc/xml/funktioner_sprog.xml", False
'objXmlHttp_funk.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/funktioner_sprog.xml", False
'objXmlHttp_funk.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/funktioner_sprog.xml", False
'objXmlHttp_funk.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/funktioner_sprog.xml", False
'objXmlHttp_funk.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/funktioner_sprog.xml", False

objXmlHttp_funk.send


Set objXmlDom_funk = objXmlHttp_funk.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_funk = Nothing



Dim Address_funk, Latitude_funk, Longitude_funk
Dim oNode_funk, oNodes_funk
Dim sXPathQuery_funk

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
sXPathQuery_funk = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_funk = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_funk = "//sprog/se"
Session.LCID = 1053
case 4
sXPathQuery_funk = "//sprog/no"
Session.LCID = 2068
case 5
sXPathQuery_funk = "//sprog/es"
Session.LCID = 1034
case 6
sXPathQuery_funk = "//sprog/de"
Session.LCID = 1031
case 7
sXPathQuery_funk = "//sprog/fr"
Session.LCID = 1036
case else
sXPathQuery_funk = "//sprog/dk"
Session.LCID = 1030
end select





Set oNode_funk = objXmlDom_funk.documentElement.selectSingleNode(sXPathQuery_funk)
Address_funk = oNode_funk.Text

Set oNodes_funk = objXmlDom_funk.documentElement.selectNodes(sXPathQuery_funk)

    For Each oNode_funk in oNodes_funk

          
        funk_txt_001 = oNode_funk.selectSingleNode("txt_1").Text
     funk_txt_002 = oNode_funk.selectSingleNode("txt_2").Text
     funk_txt_003 = oNode_funk.selectSingleNode("txt_3").Text
     funk_txt_003 = oNode_funk.selectSingleNode("txt_3").Text
    funk_txt_004 = oNode_funk.selectSingleNode("txt_4").Text
    funk_txt_005 = oNode_funk.selectSingleNode("txt_5").Text

    funk_txt_006 = oNode_funk.selectSingleNode("txt_6").Text
    funk_txt_007 = oNode_funk.selectSingleNode("txt_7").Text
    funk_txt_008 = oNode_funk.selectSingleNode("txt_8").Text
    funk_txt_009 = oNode_funk.selectSingleNode("txt_9").Text
    funk_txt_010 = oNode_funk.selectSingleNode("txt_10").Text
    funk_txt_011 = oNode_funk.selectSingleNode("txt_11").Text
    
    funk_txt_012 = oNode_funk.selectSingleNode("txt_12").Text
    funk_txt_013 = oNode_funk.selectSingleNode("txt_13").Text
    funk_txt_014 = oNode_funk.selectSingleNode("txt_14").Text
    funk_txt_015 = oNode_funk.selectSingleNode("txt_15").Text
    funk_txt_016 = oNode_funk.selectSingleNode("txt_16").Text
    funk_txt_017 = oNode_funk.selectSingleNode("txt_17").Text
    
    funk_txt_018 = oNode_funk.selectSingleNode("txt_18").Text
    funk_txt_019 = oNode_funk.selectSingleNode("txt_19").Text
    funk_txt_020 = oNode_funk.selectSingleNode("txt_20").Text
    funk_txt_021 = oNode_funk.selectSingleNode("txt_21").Text
    funk_txt_022 = oNode_funk.selectSingleNode("txt_22").Text
    funk_txt_023 = oNode_funk.selectSingleNode("txt_23").Text
    funk_txt_024 = oNode_funk.selectSingleNode("txt_24").Text
    
    funk_txt_025 = oNode_funk.selectSingleNode("txt_25").Text
    funk_txt_026 = oNode_funk.selectSingleNode("txt_26").Text
    funk_txt_027 = oNode_funk.selectSingleNode("txt_27").Text
    funk_txt_028 = oNode_funk.selectSingleNode("txt_28").Text
    funk_txt_029 = oNode_funk.selectSingleNode("txt_29").Text
    funk_txt_030 = oNode_funk.selectSingleNode("txt_30").Text
    funk_txt_031 = oNode_funk.selectSingleNode("txt_31").Text
    funk_txt_032 = oNode_funk.selectSingleNode("txt_32").Text
    funk_txt_033 = oNode_funk.selectSingleNode("txt_33").Text
    funk_txt_034 = oNode_funk.selectSingleNode("txt_34").Text
    
     funk_txt_035 = oNode_funk.selectSingleNode("txt_35").Text
      funk_txt_036 = oNode_funk.selectSingleNode("txt_36").Text
       funk_txt_037 = oNode_funk.selectSingleNode("txt_37").Text
       
        funk_txt_038 = oNode_funk.selectSingleNode("txt_38").Text
         funk_txt_039 = oNode_funk.selectSingleNode("txt_39").Text
         funk_txt_040 = oNode_funk.selectSingleNode("txt_40").Text
         funk_txt_041 = oNode_funk.selectSingleNode("txt_41").Text
         funk_txt_042 = oNode_funk.selectSingleNode("txt_42").Text
          funk_txt_043 = oNode_funk.selectSingleNode("txt_43").Text
          funk_txt_044 = oNode_funk.selectSingleNode("txt_44").Text
          
          funk_txt_045 = oNode_funk.selectSingleNode("txt_45").Text
          funk_txt_046 = oNode_funk.selectSingleNode("txt_46").Text
          funk_txt_047 = oNode_funk.selectSingleNode("txt_47").Text
          funk_txt_048 = oNode_funk.selectSingleNode("txt_48").Text
          funk_txt_049 = oNode_funk.selectSingleNode("txt_49").Text
          funk_txt_050 = oNode_funk.selectSingleNode("txt_50").Text
          
          funk_txt_051 = oNode_funk.selectSingleNode("txt_51").Text
          funk_txt_052 = oNode_funk.selectSingleNode("txt_52").Text
          funk_txt_053 = oNode_funk.selectSingleNode("txt_53").Text
          funk_txt_054 = oNode_funk.selectSingleNode("txt_54").Text
          funk_txt_055 = oNode_funk.selectSingleNode("txt_55").Text
          funk_txt_056 = oNode_funk.selectSingleNode("txt_56").Text
          funk_txt_057 = oNode_funk.selectSingleNode("txt_57").Text
          funk_txt_058 = oNode_funk.selectSingleNode("txt_58").Text
          funk_txt_059 = oNode_funk.selectSingleNode("txt_59").Text
          funk_txt_060 = oNode_funk.selectSingleNode("txt_60").Text
          funk_txt_061 = oNode_funk.selectSingleNode("txt_61").Text
          funk_txt_062 = oNode_funk.selectSingleNode("txt_62").Text
          funk_txt_063 = oNode_funk.selectSingleNode("txt_63").Text
          funk_txt_064 = oNode_funk.selectSingleNode("txt_64").Text
          funk_txt_065 = oNode_funk.selectSingleNode("txt_65").Text
          funk_txt_066 = oNode_funk.selectSingleNode("txt_66").Text
          
          funk_txt_067 = oNode_funk.selectSingleNode("txt_67").Text
          funk_txt_068 = oNode_funk.selectSingleNode("txt_68").Text
          funk_txt_069 = oNode_funk.selectSingleNode("txt_69").Text
          funk_txt_070 = oNode_funk.selectSingleNode("txt_70").Text

            funk_txt_071 = oNode_funk.selectSingleNode("txt_71").Text
            funk_txt_072 = oNode_funk.selectSingleNode("txt_72").Text
            funk_txt_073 = oNode_funk.selectSingleNode("txt_73").Text
            funk_txt_074 = oNode_funk.selectSingleNode("txt_74").Text
            funk_txt_075 = oNode_funk.selectSingleNode("txt_75").Text
            funk_txt_076 = oNode_funk.selectSingleNode("txt_76").Text
            funk_txt_077 = oNode_funk.selectSingleNode("txt_77").Text
            funk_txt_078 = oNode_funk.selectSingleNode("txt_78").Text
            funk_txt_079 = oNode_funk.selectSingleNode("txt_79").Text
          
    next




'Response.Write "week_txt_001: " & week_txt_001 & "<br>"
'Response.Write "week_txt_002: " & week_txt_002 & "<br>"


%>