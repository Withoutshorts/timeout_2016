
<% 
Dim objXMLHTTP_weekreal, objXMLDOM_weekreal, i_weekreal, strHTML_weekreal

Set objXMLDom_weekreal = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_weekreal = Server.CreateObject("Msxml2.ServerXMLHTTP")
objXmlHttp_weekreal.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/week_real_sprog.xml", False
'objXmlHttp_weekreal.open "GET", "http://localhost/inc/xml/week_real_sprog.xml", False
'objXmlHttp_weekreal.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/week_real_sprog.xml", False
'objXmlHttp_weekreal.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/week_real_sprog.xml", False
'objXmlHttp_weekreal.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/week_real_sprog.xml", False
'objXmlHttp_weekreal.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/week_real_sprog.xml", False

objXmlHttp_weekreal.send


Set objXmlDom_weekreal = objXmlHttp_weekreal.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_weekreal = Nothing



Dim Address_weekreal, Latitude_weekreal, Longitude_weekreal
Dim oNode_weekreal, oNodes_weekreal
Dim sXPathQuery_weekreal

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
sXPathQuery_weekreal = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_weekreal = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_weekreal = "//sprog/se"
Session.LCID = 1053
case 4
sXPathQuery_weekreal = "//sprog/no"
Session.LCID = 2068
case 5
sXPathQuery_weekreal = "//sprog/es"
Session.LCID = 1034
case 6
sXPathQuery_weekreal = "//sprog/de"
Session.LCID = 1031
case 7
sXPathQuery_weekreal = "//sprog/fr"
Session.LCID = 1036
case else
sXPathQuery_weekreal = "//sprog/dk"
Session.LCID = 1030
end select





Set oNode_weekreal = objXmlDom_weekreal.documentElement.selectSingleNode(sXPathQuery_weekreal)
Address_weekreal = oNode_weekreal.Text

Set oNodes_weekreal = objXmlDom_weekreal.documentElement.selectNodes(sXPathQuery_weekreal)

    For Each oNode_weekreal in oNodes_weekreal

          
    weekreal_txt_001 = oNode_weekreal.selectSingleNode("txt_1").Text
    weekreal_txt_002 = oNode_weekreal.selectSingleNode("txt_2").Text
    weekreal_txt_003 = oNode_weekreal.selectSingleNode("txt_3").Text
    weekreal_txt_003 = oNode_weekreal.selectSingleNode("txt_3").Text
    weekreal_txt_004 = oNode_weekreal.selectSingleNode("txt_4").Text
    weekreal_txt_005 = oNode_weekreal.selectSingleNode("txt_5").Text
   weekreal_txt_006 = oNode_weekreal.selectSingleNode("txt_6").Text
   weekreal_txt_007 = oNode_weekreal.selectSingleNode("txt_7").Text
   weekreal_txt_008 = oNode_weekreal.selectSingleNode("txt_8").Text
   weekreal_txt_009 = oNode_weekreal.selectSingleNode("txt_9").Text
   weekreal_txt_010 = oNode_weekreal.selectSingleNode("txt_10").Text
   weekreal_txt_011 = oNode_weekreal.selectSingleNode("txt_11").Text
    
   weekreal_txt_012 = oNode_weekreal.selectSingleNode("txt_12").Text
   weekreal_txt_013 = oNode_weekreal.selectSingleNode("txt_13").Text
   weekreal_txt_014 = oNode_weekreal.selectSingleNode("txt_14").Text
   weekreal_txt_015 = oNode_weekreal.selectSingleNode("txt_15").Text
   weekreal_txt_016 = oNode_weekreal.selectSingleNode("txt_16").Text
   weekreal_txt_017 = oNode_weekreal.selectSingleNode("txt_17").Text
    
   weekreal_txt_018 = oNode_weekreal.selectSingleNode("txt_18").Text
   weekreal_txt_019 = oNode_weekreal.selectSingleNode("txt_19").Text
   weekreal_txt_020 = oNode_weekreal.selectSingleNode("txt_20").Text
   weekreal_txt_021 = oNode_weekreal.selectSingleNode("txt_21").Text
   weekreal_txt_022 = oNode_weekreal.selectSingleNode("txt_22").Text
   weekreal_txt_023 = oNode_weekreal.selectSingleNode("txt_23").Text
   weekreal_txt_024 = oNode_weekreal.selectSingleNode("txt_24").Text
    
   weekreal_txt_025 = oNode_weekreal.selectSingleNode("txt_25").Text
   weekreal_txt_026 = oNode_weekreal.selectSingleNode("txt_26").Text
   weekreal_txt_027 = oNode_weekreal.selectSingleNode("txt_27").Text
   weekreal_txt_028 = oNode_weekreal.selectSingleNode("txt_28").Text
   weekreal_txt_029 = oNode_weekreal.selectSingleNode("txt_29").Text
   weekreal_txt_030 = oNode_weekreal.selectSingleNode("txt_30").Text
   weekreal_txt_031 = oNode_weekreal.selectSingleNode("txt_31").Text
   weekreal_txt_032 = oNode_weekreal.selectSingleNode("txt_32").Text
   weekreal_txt_033 = oNode_weekreal.selectSingleNode("txt_33").Text
   weekreal_txt_034 = oNode_weekreal.selectSingleNode("txt_34").Text
    
    weekreal_txt_035 = oNode_weekreal.selectSingleNode("txt_35").Text
     weekreal_txt_036 = oNode_weekreal.selectSingleNode("txt_36").Text
      weekreal_txt_037 = oNode_weekreal.selectSingleNode("txt_37").Text
       
       weekreal_txt_038 = oNode_weekreal.selectSingleNode("txt_38").Text
        weekreal_txt_039 = oNode_weekreal.selectSingleNode("txt_39").Text
        weekreal_txt_040 = oNode_weekreal.selectSingleNode("txt_40").Text
        weekreal_txt_041 = oNode_weekreal.selectSingleNode("txt_41").Text
        weekreal_txt_042 = oNode_weekreal.selectSingleNode("txt_42").Text
         weekreal_txt_043 = oNode_weekreal.selectSingleNode("txt_43").Text
         weekreal_txt_044 = oNode_weekreal.selectSingleNode("txt_44").Text
          
         weekreal_txt_045 = oNode_weekreal.selectSingleNode("txt_45").Text
         weekreal_txt_046 = oNode_weekreal.selectSingleNode("txt_46").Text
         weekreal_txt_047 = oNode_weekreal.selectSingleNode("txt_47").Text
         weekreal_txt_048 = oNode_weekreal.selectSingleNode("txt_48").Text
         weekreal_txt_049 = oNode_weekreal.selectSingleNode("txt_49").Text
         weekreal_txt_050 = oNode_weekreal.selectSingleNode("txt_50").Text
         weekreal_txt_051 = oNode_weekreal.selectSingleNode("txt_51").Text
         weekreal_txt_052 = oNode_weekreal.selectSingleNode("txt_52").Text
         weekreal_txt_053 = oNode_weekreal.selectSingleNode("txt_53").Text
         weekreal_txt_054 = oNode_weekreal.selectSingleNode("txt_54").Text
         weekreal_txt_055 = oNode_weekreal.selectSingleNode("txt_55").Text
         weekreal_txt_056 = oNode_weekreal.selectSingleNode("txt_56").Text
         weekreal_txt_057 = oNode_weekreal.selectSingleNode("txt_57").Text
        
        weekreal_txt_058 = oNode_weekreal.selectSingleNode("txt_58").Text
        weekreal_txt_059 = oNode_weekreal.selectSingleNode("txt_59").Text
    
   weekreal_txt_060 = oNode_weekreal.selectSingleNode("txt_60").Text
   weekreal_txt_061 = oNode_weekreal.selectSingleNode("txt_61").Text
   weekreal_txt_062 = oNode_weekreal.selectSingleNode("txt_62").Text
   weekreal_txt_063 = oNode_weekreal.selectSingleNode("txt_63").Text
   weekreal_txt_064 = oNode_weekreal.selectSingleNode("txt_64").Text
   weekreal_txt_065 = oNode_weekreal.selectSingleNode("txt_65").Text
   weekreal_txt_066 = oNode_weekreal.selectSingleNode("txt_66").Text
   weekreal_txt_067 = oNode_weekreal.selectSingleNode("txt_67").Text
   weekreal_txt_068 = oNode_weekreal.selectSingleNode("txt_68").Text
   weekreal_txt_069 = oNode_weekreal.selectSingleNode("txt_69").Text
   weekreal_txt_070 = oNode_weekreal.selectSingleNode("txt_70").Text
   weekreal_txt_071 = oNode_weekreal.selectSingleNode("txt_71").Text
   weekreal_txt_072 = oNode_weekreal.selectSingleNode("txt_72").Text
   weekreal_txt_070 = oNode_weekreal.selectSingleNode("txt_73").Text
   weekreal_txt_074 = oNode_weekreal.selectSingleNode("txt_74").Text
   weekreal_txt_075 = oNode_weekreal.selectSingleNode("txt_75").Text
   weekreal_txt_076 = oNode_weekreal.selectSingleNode("txt_76").Text
   weekreal_txt_077 = oNode_weekreal.selectSingleNode("txt_77").Text
   weekreal_txt_078 = oNode_weekreal.selectSingleNode("txt_78").Text
   weekreal_txt_079 = oNode_weekreal.selectSingleNode("txt_79").Text
    
   weekreal_txt_080 = oNode_weekreal.selectSingleNode("txt_80").Text
   weekreal_txt_081 = oNode_weekreal.selectSingleNode("txt_81").Text
   weekreal_txt_082 = oNode_weekreal.selectSingleNode("txt_82").Text
   weekreal_txt_083 = oNode_weekreal.selectSingleNode("txt_83").Text
   weekreal_txt_084 = oNode_weekreal.selectSingleNode("txt_84").Text
   weekreal_txt_085 = oNode_weekreal.selectSingleNode("txt_85").Text
   weekreal_txt_086 = oNode_weekreal.selectSingleNode("txt_86").Text
   weekreal_txt_087 = oNode_weekreal.selectSingleNode("txt_87").Text
   weekreal_txt_088 = oNode_weekreal.selectSingleNode("txt_88").Text
   weekreal_txt_089 = oNode_weekreal.selectSingleNode("txt_89").Text
   weekreal_txt_090 = oNode_weekreal.selectSingleNode("txt_90").Text
   weekreal_txt_091 = oNode_weekreal.selectSingleNode("txt_91").Text
   weekreal_txt_092 = oNode_weekreal.selectSingleNode("txt_92").Text
   weekreal_txt_093 = oNode_weekreal.selectSingleNode("txt_93").Text
   weekreal_txt_094 = oNode_weekreal.selectSingleNode("txt_94").Text
   weekreal_txt_095 = oNode_weekreal.selectSingleNode("txt_95").Text
   weekreal_txt_096 = oNode_weekreal.selectSingleNode("txt_96").Text
   weekreal_txt_097 = oNode_weekreal.selectSingleNode("txt_97").Text
   weekreal_txt_098 = oNode_weekreal.selectSingleNode("txt_98").Text
   weekreal_txt_099 = oNode_weekreal.selectSingleNode("txt_99").Text
   weekreal_txt_100 = oNode_weekreal.selectSingleNode("txt_100").Text
   weekreal_txt_101 = oNode_weekreal.selectSingleNode("txt_101").Text
   weekreal_txt_102 = oNode_weekreal.selectSingleNode("txt_102").Text
   weekreal_txt_103 = oNode_weekreal.selectSingleNode("txt_103").Text
   weekreal_txt_104 = oNode_weekreal.selectSingleNode("txt_104").Text
   weekreal_txt_105 = oNode_weekreal.selectSingleNode("txt_105").Text
   weekreal_txt_106 = oNode_weekreal.selectSingleNode("txt_106").Text
   weekreal_txt_107 = oNode_weekreal.selectSingleNode("txt_107").Text
   weekreal_txt_108 = oNode_weekreal.selectSingleNode("txt_108").Text
   weekreal_txt_109 = oNode_weekreal.selectSingleNode("txt_109").Text
   weekreal_txt_110 = oNode_weekreal.selectSingleNode("txt_110").Text
   weekreal_txt_111 = oNode_weekreal.selectSingleNode("txt_111").Text
   weekreal_txt_112 = oNode_weekreal.selectSingleNode("txt_112").Text
   weekreal_txt_113 = oNode_weekreal.selectSingleNode("txt_113").Text
   weekreal_txt_114 = oNode_weekreal.selectSingleNode("txt_114").Text
   weekreal_txt_115 = oNode_weekreal.selectSingleNode("txt_115").Text
   weekreal_txt_116 = oNode_weekreal.selectSingleNode("txt_116").Text
   weekreal_txt_117 = oNode_weekreal.selectSingleNode("txt_117").Text
   weekreal_txt_118 = oNode_weekreal.selectSingleNode("txt_118").Text
   weekreal_txt_119 = oNode_weekreal.selectSingleNode("txt_119").Text
    
   weekreal_txt_120 = oNode_weekreal.selectSingleNode("txt_120").Text

   weekreal_txt_122 = oNode_weekreal.selectSingleNode("txt_122").Text
   weekreal_txt_123 = oNode_weekreal.selectSingleNode("txt_123").Text
   weekreal_txt_124 = oNode_weekreal.selectSingleNode("txt_124").Text
          
    next




'Response.Write "week_txt_001: " & week_txt_001 & "<br>"
'Response.Write "week_txt_002: " & week_txt_002 & "<br>"


%>