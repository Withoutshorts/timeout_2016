
<% 
Dim objXMLHTTP_extra, objXMLDOM_extra, i_extra, strHTML_extra

Set objXMLDom_extra = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_extra = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_extra.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/extra_sprog.xml", False
'objXmlHttp_extra.open "GET", "http://localhost/inc/xml/extra_sprog.xml", False
'objXmlHttp_extra.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/extra_sprog.xml", False
objXmlHttp_extra.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/extra_sprog.xml", False
'objXmlHttp_extra.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/extra_sprog.xml", False
'objXmlHttp_extra.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/extra_sprog.xml", False

objXmlHttp_extra.send


Set objXmlDom_extra = objXmlHttp_extra.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_extra = Nothing



Dim Address_extra, Latitude_extra, Longitude_extra
Dim oNode_extra, oNodes_extra
Dim sXPathQuery_extra

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
sXPathQuery_extra = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_extra = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_extra = "//sprog/se"
Session.LCID = 1053
case 4
sXPathQuery_extra = "//sprog/no"
Session.LCID = 2068
case 5
sXPathQuery_extra = "//sprog/es"
Session.LCID = 1034
case 6
sXPathQuery_extra = "//sprog/de"
Session.LCID = 1031
case 7
sXPathQuery_extra = "//sprog/fr"
Session.LCID = 1036
case else
sXPathQuery_extra = "//sprog/dk"
Session.LCID = 1030
end select





Set oNode_extra = objXmlDom_extra.documentElement.selectSingleNode(sXPathQuery_extra)
Address_extra = oNode_extra.Text

Set oNodes_extra = objXmlDom_extra.documentElement.selectNodes(sXPathQuery_extra)

    For Each oNode_extra in oNodes_extra

    extra_txt_001 = oNode_extra.selectSingleNode("txt_1").Text
    extra_txt_002 = oNode_extra.selectSingleNode("txt_2").Text
    extra_txt_003 = oNode_extra.selectSingleNode("txt_3").Text
    extra_txt_003 = oNode_extra.selectSingleNode("txt_3").Text
    extra_txt_004 = oNode_extra.selectSingleNode("txt_4").Text
    extra_txt_005 = oNode_extra.selectSingleNode("txt_5").Text

   extra_txt_006 = oNode_extra.selectSingleNode("txt_6").Text
   extra_txt_007 = oNode_extra.selectSingleNode("txt_7").Text
   extra_txt_008 = oNode_extra.selectSingleNode("txt_8").Text
   extra_txt_009 = oNode_extra.selectSingleNode("txt_9").Text
   extra_txt_010 = oNode_extra.selectSingleNode("txt_10").Text
   extra_txt_011 = oNode_extra.selectSingleNode("txt_11").Text
    
   extra_txt_012 = oNode_extra.selectSingleNode("txt_12").Text
   extra_txt_013 = oNode_extra.selectSingleNode("txt_13").Text
   extra_txt_014 = oNode_extra.selectSingleNode("txt_14").Text
   extra_txt_015 = oNode_extra.selectSingleNode("txt_15").Text
   extra_txt_016 = oNode_extra.selectSingleNode("txt_16").Text
   extra_txt_017 = oNode_extra.selectSingleNode("txt_17").Text
    
   extra_txt_018 = oNode_extra.selectSingleNode("txt_18").Text
   extra_txt_019 = oNode_extra.selectSingleNode("txt_19").Text
   extra_txt_020 = oNode_extra.selectSingleNode("txt_20").Text
   extra_txt_021 = oNode_extra.selectSingleNode("txt_21").Text
   extra_txt_022 = oNode_extra.selectSingleNode("txt_22").Text
   extra_txt_023 = oNode_extra.selectSingleNode("txt_23").Text
   extra_txt_024 = oNode_extra.selectSingleNode("txt_24").Text
    
   extra_txt_025 = oNode_extra.selectSingleNode("txt_25").Text
   extra_txt_026 = oNode_extra.selectSingleNode("txt_26").Text
   extra_txt_027 = oNode_extra.selectSingleNode("txt_27").Text
   extra_txt_028 = oNode_extra.selectSingleNode("txt_28").Text
   extra_txt_029 = oNode_extra.selectSingleNode("txt_29").Text
   extra_txt_030 = oNode_extra.selectSingleNode("txt_30").Text
   extra_txt_031 = oNode_extra.selectSingleNode("txt_31").Text
   extra_txt_032 = oNode_extra.selectSingleNode("txt_32").Text
   extra_txt_033 = oNode_extra.selectSingleNode("txt_33").Text
   extra_txt_034 = oNode_extra.selectSingleNode("txt_34").Text
    
   extra_txt_035 = oNode_extra.selectSingleNode("txt_35").Text
   extra_txt_036 = oNode_extra.selectSingleNode("txt_36").Text
   extra_txt_037 = oNode_extra.selectSingleNode("txt_37").Text
       
   extra_txt_038 = oNode_extra.selectSingleNode("txt_38").Text
   extra_txt_039 = oNode_extra.selectSingleNode("txt_39").Text
   extra_txt_040 = oNode_extra.selectSingleNode("txt_40").Text
   extra_txt_041 = oNode_extra.selectSingleNode("txt_41").Text
   extra_txt_042 = oNode_extra.selectSingleNode("txt_42").Text
   extra_txt_043 = oNode_extra.selectSingleNode("txt_43").Text
   extra_txt_044 = oNode_extra.selectSingleNode("txt_44").Text
          
   extra_txt_045 = oNode_extra.selectSingleNode("txt_45").Text
   extra_txt_046 = oNode_extra.selectSingleNode("txt_46").Text
   extra_txt_047 = oNode_extra.selectSingleNode("txt_47").Text
   extra_txt_048 = oNode_extra.selectSingleNode("txt_48").Text
   extra_txt_049 = oNode_extra.selectSingleNode("txt_49").Text
   extra_txt_050 = oNode_extra.selectSingleNode("txt_50").Text
          
   extra_txt_051 = oNode_extra.selectSingleNode("txt_51").Text
   extra_txt_052 = oNode_extra.selectSingleNode("txt_52").Text
   extra_txt_053 = oNode_extra.selectSingleNode("txt_53").Text
   extra_txt_054 = oNode_extra.selectSingleNode("txt_54").Text
   extra_txt_055 = oNode_extra.selectSingleNode("txt_55").Text
   extra_txt_056 = oNode_extra.selectSingleNode("txt_56").Text

   extra_txt_057 = oNode_extra.selectSingleNode("txt_57").Text
   extra_txt_058 = oNode_extra.selectSingleNode("txt_58").Text
   extra_txt_059 = oNode_extra.selectSingleNode("txt_59").Text
   extra_txt_060 = oNode_extra.selectSingleNode("txt_60").Text
   extra_txt_061 = oNode_extra.selectSingleNode("txt_61").Text
   extra_txt_062 = oNode_extra.selectSingleNode("txt_62").Text
   extra_txt_063 = oNode_extra.selectSingleNode("txt_63").Text
   extra_txt_064 = oNode_extra.selectSingleNode("txt_64").Text
   extra_txt_065 = oNode_extra.selectSingleNode("txt_65").Text
   extra_txt_066 = oNode_extra.selectSingleNode("txt_66").Text
   extra_txt_067 = oNode_extra.selectSingleNode("txt_67").Text
   extra_txt_068 = oNode_extra.selectSingleNode("txt_68").Text
   extra_txt_069 = oNode_extra.selectSingleNode("txt_69").Text

   extra_txt_070 = oNode_extra.selectSingleNode("txt_70").Text
   extra_txt_071 = oNode_extra.selectSingleNode("txt_71").Text
   extra_txt_072 = oNode_extra.selectSingleNode("txt_72").Text
   extra_txt_073 = oNode_extra.selectSingleNode("txt_73").Text
   extra_txt_074 = oNode_extra.selectSingleNode("txt_74").Text
   extra_txt_075 = oNode_extra.selectSingleNode("txt_75").Text
   extra_txt_076 = oNode_extra.selectSingleNode("txt_76").Text
   extra_txt_077 = oNode_extra.selectSingleNode("txt_77").Text
   extra_txt_078 = oNode_extra.selectSingleNode("txt_78").Text

   extra_txt_079 = oNode_extra.selectSingleNode("txt_79").Text
   extra_txt_080 = oNode_extra.selectSingleNode("txt_80").Text
   extra_txt_081 = oNode_extra.selectSingleNode("txt_81").Text
   extra_txt_082 = oNode_extra.selectSingleNode("txt_82").Text
   extra_txt_083 = oNode_extra.selectSingleNode("txt_83").Text
   extra_txt_084 = oNode_extra.selectSingleNode("txt_84").Text
   extra_txt_085 = oNode_extra.selectSingleNode("txt_85").Text
   extra_txt_086 = oNode_extra.selectSingleNode("txt_86").Text
   extra_txt_087 = oNode_extra.selectSingleNode("txt_87").Text
   extra_txt_088 = oNode_extra.selectSingleNode("txt_88").Text
   extra_txt_089 = oNode_extra.selectSingleNode("txt_89").Text
   extra_txt_090 = oNode_extra.selectSingleNode("txt_90").Text
   extra_txt_091 = oNode_extra.selectSingleNode("txt_91").Text
   extra_txt_092 = oNode_extra.selectSingleNode("txt_92").Text
   extra_txt_093 = oNode_extra.selectSingleNode("txt_93").Text
   extra_txt_094 = oNode_extra.selectSingleNode("txt_94").Text

    extra_txt_095 = oNode_extra.selectSingleNode("txt_95").Text
    extra_txt_096 = oNode_extra.selectSingleNode("txt_96").Text
    extra_txt_097 = oNode_extra.selectSingleNode("txt_97").Text
    extra_txt_098 = oNode_extra.selectSingleNode("txt_98").Text
    extra_txt_099 = oNode_extra.selectSingleNode("txt_99").Text
          
    extra_txt_100 = oNode_extra.selectSingleNode("txt_100").Text
    extra_txt_101 = oNode_extra.selectSingleNode("txt_101").Text
    extra_txt_102 = oNode_extra.selectSingleNode("txt_102").Text
    extra_txt_103 = oNode_extra.selectSingleNode("txt_103").Text
          
    extra_txt_104 = oNode_extra.selectSingleNode("txt_104").Text
    extra_txt_105 = oNode_extra.selectSingleNode("txt_105").Text
    extra_txt_106 = oNode_extra.selectSingleNode("txt_106").Text
         
          
    extra_txt_116 = oNode_extra.selectSingleNode("txt_116").Text
    extra_txt_117 = oNode_extra.selectSingleNode("txt_117").Text
    extra_txt_118 = oNode_extra.selectSingleNode("txt_118").Text
    extra_txt_119 = oNode_extra.selectSingleNode("txt_119").Text
    extra_txt_120 = oNode_extra.selectSingleNode("txt_120").Text
    extra_txt_121 = oNode_extra.selectSingleNode("txt_121").Text
    extra_txt_122 = oNode_extra.selectSingleNode("txt_122").Text
    extra_txt_123 = oNode_extra.selectSingleNode("txt_123").Text
    extra_txt_124 = oNode_extra.selectSingleNode("txt_124").Text
    extra_txt_125 = oNode_extra.selectSingleNode("txt_125").Text
    extra_txt_126 = oNode_extra.selectSingleNode("txt_126").Text
    extra_txt_127 = oNode_extra.selectSingleNode("txt_127").Text
    extra_txt_128 = oNode_extra.selectSingleNode("txt_128").Text
    extra_txt_129 = oNode_extra.selectSingleNode("txt_129").Text
    extra_txt_130 = oNode_extra.selectSingleNode("txt_130").Text
   

  
          
    next




'Response.Write "extra_txt_001: " & extra_txt_001 & "<br>"
'Response.Write "extra_txt_002: " & extra_txt_002 & "<br>"


%>