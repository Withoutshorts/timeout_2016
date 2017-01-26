
<% 
Dim objXMLHTTP_timereg, objXMLDOM_timereg, i_timereg, strHTML_timereg

Set objXMLDom_timereg = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_timereg = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_timereg.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/timereg_akt_sprog.xml", False
'objXmlHttp_timereg.open "GET", "http://localhost/inc/xml/timereg_akt_sprog.xml", False
'objXmlHttp_timereg.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/timereg_akt_sprog.xml", False
'objXmlHttp_timereg.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/timereg_akt_sprog.xml", False
'objXmlHttp_timereg.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/timereg_akt_sprog.xml", False
objXmlHttp_timereg.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/timereg_akt_sprog.xml", False

objXmlHttp_timereg.send


Set objXmlDom_timereg = objXmlHttp_timereg.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_timereg = Nothing



Dim Address_timereg, Latitude_timereg, Longitude_timereg
Dim oNode_timereg, oNodes_timereg
Dim sXPathQuery_timereg

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
sXPathQuery_timereg = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_timereg = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_timereg = "//sprog/se"
Session.LCID = 1053
case 4
sXPathQuery_timereg = "//sprog/no"
Session.LCID = 2068
case 5
sXPathQuery_timereg = "//sprog/es"
Session.LCID = 1034
case 6
sXPathQuery_timereg = "//sprog/de"
Session.LCID = 1031
case 7
sXPathQuery_timereg = "//sprog/fr"
Session.LCID = 1036
case else
sXPathQuery_timereg = "//sprog/dk"
Session.LCID = 1030
end select





Set oNode_timereg = objXmlDom_timereg.documentElement.selectSingleNode(sXPathQuery_timereg)
Address_timereg = oNode_timereg.Text

Set oNodes_timereg = objXmlDom_timereg.documentElement.selectNodes(sXPathQuery_timereg)

    For Each oNode_timereg in oNodes_timereg

          
        timereg_txt_001 = oNode_timereg.selectSingleNode("txt_1").Text
     timereg_txt_002 = oNode_timereg.selectSingleNode("txt_2").Text
     timereg_txt_003 = oNode_timereg.selectSingleNode("txt_3").Text
     timereg_txt_003 = oNode_timereg.selectSingleNode("txt_3").Text
    timereg_txt_004 = oNode_timereg.selectSingleNode("txt_4").Text
    timereg_txt_005 = oNode_timereg.selectSingleNode("txt_5").Text

    timereg_txt_006 = oNode_timereg.selectSingleNode("txt_6").Text
    timereg_txt_007 = oNode_timereg.selectSingleNode("txt_7").Text
    timereg_txt_008 = oNode_timereg.selectSingleNode("txt_8").Text
    timereg_txt_009 = oNode_timereg.selectSingleNode("txt_9").Text
    timereg_txt_010 = oNode_timereg.selectSingleNode("txt_10").Text
    timereg_txt_011 = oNode_timereg.selectSingleNode("txt_11").Text
    
    timereg_txt_012 = oNode_timereg.selectSingleNode("txt_12").Text
    timereg_txt_013 = oNode_timereg.selectSingleNode("txt_13").Text
    timereg_txt_014 = oNode_timereg.selectSingleNode("txt_14").Text
    timereg_txt_015 = oNode_timereg.selectSingleNode("txt_15").Text
    timereg_txt_016 = oNode_timereg.selectSingleNode("txt_16").Text
    timereg_txt_017 = oNode_timereg.selectSingleNode("txt_17").Text
    
    timereg_txt_018 = oNode_timereg.selectSingleNode("txt_18").Text
    timereg_txt_019 = oNode_timereg.selectSingleNode("txt_19").Text
    timereg_txt_020 = oNode_timereg.selectSingleNode("txt_20").Text
    timereg_txt_021 = oNode_timereg.selectSingleNode("txt_21").Text
    timereg_txt_022 = oNode_timereg.selectSingleNode("txt_22").Text
    timereg_txt_023 = oNode_timereg.selectSingleNode("txt_23").Text
    timereg_txt_024 = oNode_timereg.selectSingleNode("txt_24").Text
    
    timereg_txt_025 = oNode_timereg.selectSingleNode("txt_25").Text
    timereg_txt_026 = oNode_timereg.selectSingleNode("txt_26").Text
    timereg_txt_027 = oNode_timereg.selectSingleNode("txt_27").Text
    timereg_txt_028 = oNode_timereg.selectSingleNode("txt_28").Text
    timereg_txt_029 = oNode_timereg.selectSingleNode("txt_29").Text
    timereg_txt_030 = oNode_timereg.selectSingleNode("txt_30").Text
    timereg_txt_031 = oNode_timereg.selectSingleNode("txt_31").Text
    timereg_txt_032 = oNode_timereg.selectSingleNode("txt_32").Text
    timereg_txt_033 = oNode_timereg.selectSingleNode("txt_33").Text
    timereg_txt_034 = oNode_timereg.selectSingleNode("txt_34").Text
    
     timereg_txt_035 = oNode_timereg.selectSingleNode("txt_35").Text
      timereg_txt_036 = oNode_timereg.selectSingleNode("txt_36").Text
       timereg_txt_037 = oNode_timereg.selectSingleNode("txt_37").Text
       
        timereg_txt_038 = oNode_timereg.selectSingleNode("txt_38").Text
         timereg_txt_039 = oNode_timereg.selectSingleNode("txt_39").Text
         timereg_txt_040 = oNode_timereg.selectSingleNode("txt_40").Text
         timereg_txt_041 = oNode_timereg.selectSingleNode("txt_41").Text
         timereg_txt_042 = oNode_timereg.selectSingleNode("txt_42").Text
          timereg_txt_043 = oNode_timereg.selectSingleNode("txt_43").Text
          timereg_txt_044 = oNode_timereg.selectSingleNode("txt_44").Text
          
          timereg_txt_045 = oNode_timereg.selectSingleNode("txt_45").Text
          timereg_txt_046 = oNode_timereg.selectSingleNode("txt_46").Text
          timereg_txt_047 = oNode_timereg.selectSingleNode("txt_47").Text
          timereg_txt_048 = oNode_timereg.selectSingleNode("txt_48").Text
          timereg_txt_049 = oNode_timereg.selectSingleNode("txt_49").Text
          timereg_txt_050 = oNode_timereg.selectSingleNode("txt_50").Text
          
          timereg_txt_051 = oNode_timereg.selectSingleNode("txt_51").Text
          timereg_txt_052 = oNode_timereg.selectSingleNode("txt_52").Text
          timereg_txt_053 = oNode_timereg.selectSingleNode("txt_53").Text
          timereg_txt_054 = oNode_timereg.selectSingleNode("txt_54").Text
          timereg_txt_055 = oNode_timereg.selectSingleNode("txt_55").Text
          timereg_txt_056 = oNode_timereg.selectSingleNode("txt_56").Text
          timereg_txt_057 = oNode_timereg.selectSingleNode("txt_57").Text
          timereg_txt_058 = oNode_timereg.selectSingleNode("txt_58").Text
          timereg_txt_059 = oNode_timereg.selectSingleNode("txt_59").Text
          timereg_txt_060 = oNode_timereg.selectSingleNode("txt_60").Text
          timereg_txt_061 = oNode_timereg.selectSingleNode("txt_61").Text
          timereg_txt_062 = oNode_timereg.selectSingleNode("txt_62").Text
          timereg_txt_063 = oNode_timereg.selectSingleNode("txt_63").Text
          timereg_txt_064 = oNode_timereg.selectSingleNode("txt_64").Text
          timereg_txt_065 = oNode_timereg.selectSingleNode("txt_65").Text
          timereg_txt_066 = oNode_timereg.selectSingleNode("txt_66").Text
          
          timereg_txt_067 = oNode_timereg.selectSingleNode("txt_67").Text
          timereg_txt_068 = oNode_timereg.selectSingleNode("txt_68").Text
          timereg_txt_069 = oNode_timereg.selectSingleNode("txt_69").Text
          timereg_txt_070 = oNode_timereg.selectSingleNode("txt_70").Text
          timereg_txt_071 = oNode_timereg.selectSingleNode("txt_71").Text
          timereg_txt_072 = oNode_timereg.selectSingleNode("txt_72").Text
          timereg_txt_073 = oNode_timereg.selectSingleNode("txt_73").Text
          timereg_txt_074 = oNode_timereg.selectSingleNode("txt_74").Text
          timereg_txt_075 = oNode_timereg.selectSingleNode("txt_75").Text
          timereg_txt_076 = oNode_timereg.selectSingleNode("txt_76").Text
          timereg_txt_077 = oNode_timereg.selectSingleNode("txt_77").Text
          timereg_txt_078 = oNode_timereg.selectSingleNode("txt_78").Text
          timereg_txt_079 = oNode_timereg.selectSingleNode("txt_79").Text
        
          
          timereg_txt_080 = oNode_timereg.selectSingleNode("txt_80").Text
          timereg_txt_081 = oNode_timereg.selectSingleNode("txt_81").Text
          timereg_txt_082 = oNode_timereg.selectSingleNode("txt_82").Text
          timereg_txt_083 = oNode_timereg.selectSingleNode("txt_83").Text
          
          timereg_txt_084 = oNode_timereg.selectSingleNode("txt_84").Text
          timereg_txt_085 = oNode_timereg.selectSingleNode("txt_85").Text
          timereg_txt_086 = oNode_timereg.selectSingleNode("txt_86").Text
          
          timereg_txt_087 = oNode_timereg.selectSingleNode("txt_87").Text
          
          timereg_txt_088 = oNode_timereg.selectSingleNode("txt_88").Text
          timereg_txt_089 = oNode_timereg.selectSingleNode("txt_89").Text
          timereg_txt_090 = oNode_timereg.selectSingleNode("txt_90").Text
          
          timereg_txt_091 = oNode_timereg.selectSingleNode("txt_91").Text
          timereg_txt_092 = oNode_timereg.selectSingleNode("txt_92").Text
          timereg_txt_093 = oNode_timereg.selectSingleNode("txt_93").Text
          timereg_txt_094 = oNode_timereg.selectSingleNode("txt_94").Text
          timereg_txt_095 = oNode_timereg.selectSingleNode("txt_95").Text
          timereg_txt_096 = oNode_timereg.selectSingleNode("txt_96").Text
          timereg_txt_097 = oNode_timereg.selectSingleNode("txt_97").Text
          timereg_txt_098 = oNode_timereg.selectSingleNode("txt_98").Text
          timereg_txt_099 = oNode_timereg.selectSingleNode("txt_99").Text
          
          timereg_txt_100 = oNode_timereg.selectSingleNode("txt_100").Text
          timereg_txt_101 = oNode_timereg.selectSingleNode("txt_101").Text
          timereg_txt_102 = oNode_timereg.selectSingleNode("txt_102").Text
          timereg_txt_103 = oNode_timereg.selectSingleNode("txt_103").Text
          
          timereg_txt_104 = oNode_timereg.selectSingleNode("txt_104").Text
          timereg_txt_105 = oNode_timereg.selectSingleNode("txt_105").Text
          timereg_txt_106 = oNode_timereg.selectSingleNode("txt_106").Text
          timereg_txt_107 = oNode_timereg.selectSingleNode("txt_107").Text
          timereg_txt_108 = oNode_timereg.selectSingleNode("txt_108").Text
          timereg_txt_109 = oNode_timereg.selectSingleNode("txt_109").Text
          timereg_txt_110 = oNode_timereg.selectSingleNode("txt_110").Text
          timereg_txt_111 = oNode_timereg.selectSingleNode("txt_111").Text
          timereg_txt_112 = oNode_timereg.selectSingleNode("txt_112").Text
          timereg_txt_113 = oNode_timereg.selectSingleNode("txt_113").Text
          timereg_txt_114 = oNode_timereg.selectSingleNode("txt_114").Text
        timereg_txt_115 = oNode_timereg.selectSingleNode("txt_115").Text
        timereg_txt_116 = oNode_timereg.selectSingleNode("txt_116").Text
        timereg_txt_117 = oNode_timereg.selectSingleNode("txt_117").Text
        timereg_txt_118 = oNode_timereg.selectSingleNode("txt_118").Text
        timereg_txt_119 = oNode_timereg.selectSingleNode("txt_119").Text
        timereg_txt_120 = oNode_timereg.selectSingleNode("txt_120").Text
        timereg_txt_121 = oNode_timereg.selectSingleNode("txt_121").Text
        timereg_txt_122 = oNode_timereg.selectSingleNode("txt_122").Text
        timereg_txt_123 = oNode_timereg.selectSingleNode("txt_123").Text

            timereg_txt_124 = oNode_timereg.selectSingleNode("txt_124").Text
           timereg_txt_125 = oNode_timereg.selectSingleNode("txt_125").Text
            timereg_txt_126 = oNode_timereg.selectSingleNode("txt_126").Text
            timereg_txt_127 = oNode_timereg.selectSingleNode("txt_127").Text
            timereg_txt_128 = oNode_timereg.selectSingleNode("txt_128").Text
            
            timereg_txt_129 = oNode_timereg.selectSingleNode("txt_129").Text
            timereg_txt_130 = oNode_timereg.selectSingleNode("txt_130").Text
            timereg_txt_131 = oNode_timereg.selectSingleNode("txt_131").Text
            timereg_txt_132 = oNode_timereg.selectSingleNode("txt_132").Text
            timereg_txt_133 = oNode_timereg.selectSingleNode("txt_133").Text
            
            timereg_txt_134 = oNode_timereg.selectSingleNode("txt_134").Text
            timereg_txt_135 = oNode_timereg.selectSingleNode("txt_135").Text
            timereg_txt_136 = oNode_timereg.selectSingleNode("txt_136").Text
            timereg_txt_137 = oNode_timereg.selectSingleNode("txt_137").Text
            timereg_txt_138 = oNode_timereg.selectSingleNode("txt_138").Text
            timereg_txt_139 = oNode_timereg.selectSingleNode("txt_139").Text
            
            timereg_txt_140 = oNode_timereg.selectSingleNode("txt_140").Text
            timereg_txt_141 = oNode_timereg.selectSingleNode("txt_141").Text
            
            timereg_txt_142 = oNode_timereg.selectSingleNode("txt_142").Text
            timereg_txt_143 = oNode_timereg.selectSingleNode("txt_143").Text
            timereg_txt_144 = oNode_timereg.selectSingleNode("txt_144").Text
            timereg_txt_145 = oNode_timereg.selectSingleNode("txt_145").Text
            timereg_txt_146 = oNode_timereg.selectSingleNode("txt_146").Text
            
            timereg_txt_147 = oNode_timereg.selectSingleNode("txt_147").Text
            timereg_txt_148 = oNode_timereg.selectSingleNode("txt_148").Text
            timereg_txt_149 = oNode_timereg.selectSingleNode("txt_149").Text
            
            timereg_txt_150 = oNode_timereg.selectSingleNode("txt_150").Text
            timereg_txt_151 = oNode_timereg.selectSingleNode("txt_151").Text
            timereg_txt_152 = oNode_timereg.selectSingleNode("txt_152").Text
            timereg_txt_153 = oNode_timereg.selectSingleNode("txt_153").Text
            timereg_txt_154 = oNode_timereg.selectSingleNode("txt_154").Text

          
    next




'Response.Write "week_txt_001: " & week_txt_001 & "<br>"
'Response.Write "week_txt_002: " & week_txt_002 & "<br>"


%>