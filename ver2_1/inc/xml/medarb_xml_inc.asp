
<% 
Dim objXMLHTTP_medarb, objXMLDOM_medarb, i_medarb, strHTML_medarb

Set objXMLDom_medarb = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_medarb = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_medarb.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/medarb_sprog.xml", False
'objXmlHttp_medarb.open "GET", "http://localhost/inc/xml/medarb_sprog.xml", False
'objXmlHttp_medarb.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/medarb_sprog.xml", False
'objXmlHttp_medarb.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/medarb_sprog.xml", False
'objXmlHttp_medarb.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/medarb_sprog.xml", False
'objXmlHttp_medarb.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/medarb_sprog.xml", False
objXmlHttp_medarb.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/medarb_sprog.xml", False

objXmlHttp_medarb.send


Set objXmlDom_medarb = objXmlHttp_medarb.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_medarb = Nothing



Dim Address_medarb, Latitude_medarb, Longitude_medarb
Dim oNode_medarb, oNodes_medarb
Dim sXPathQuery_medarb

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
sXPathQuery_medarb = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_medarb = "//sprog/uk"
if lto = "a27" then
    sXPathQuery_medarb = "//sprog/trainerlog_uk"
end if
'Session.LCID = 2057
case 3
sXPathQuery_medarb = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_medarb = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_medarb = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_medarb = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_medarb = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_medarb = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle bel�b og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_medarb = objXmlDom_medarb.documentElement.selectSingleNode(sXPathQuery_medarb)
Address_medarb = oNode_medarb.Text

Set oNodes_medarb = objXmlDom_medarb.documentElement.selectNodes(sXPathQuery_medarb)

    For Each oNode_medarb in oNodes_medarb

          
        medarb_txt_001 = oNode_medarb.selectSingleNode("txt_1").Text
        medarb_txt_002 = oNode_medarb.selectSingleNode("txt_2").Text
        medarb_txt_003 = oNode_medarb.selectSingleNode("txt_3").Text
        medarb_txt_003 = oNode_medarb.selectSingleNode("txt_3").Text
        medarb_txt_004 = oNode_medarb.selectSingleNode("txt_4").Text
        medarb_txt_005 = oNode_medarb.selectSingleNode("txt_5").Text

        medarb_txt_006 = oNode_medarb.selectSingleNode("txt_6").Text
        medarb_txt_007 = oNode_medarb.selectSingleNode("txt_7").Text
        medarb_txt_008 = oNode_medarb.selectSingleNode("txt_8").Text
        medarb_txt_009 = oNode_medarb.selectSingleNode("txt_9").Text
        medarb_txt_010 = oNode_medarb.selectSingleNode("txt_10").Text
        medarb_txt_011 = oNode_medarb.selectSingleNode("txt_11").Text
    
        medarb_txt_012 = oNode_medarb.selectSingleNode("txt_12").Text
        medarb_txt_013 = oNode_medarb.selectSingleNode("txt_13").Text
        medarb_txt_014 = oNode_medarb.selectSingleNode("txt_14").Text
        medarb_txt_015 = oNode_medarb.selectSingleNode("txt_15").Text
        medarb_txt_016 = oNode_medarb.selectSingleNode("txt_16").Text
        medarb_txt_017 = oNode_medarb.selectSingleNode("txt_17").Text
    
        medarb_txt_018 = oNode_medarb.selectSingleNode("txt_18").Text
        medarb_txt_019 = oNode_medarb.selectSingleNode("txt_19").Text
        medarb_txt_020 = oNode_medarb.selectSingleNode("txt_20").Text
        medarb_txt_021 = oNode_medarb.selectSingleNode("txt_21").Text
        medarb_txt_022 = oNode_medarb.selectSingleNode("txt_22").Text
        medarb_txt_023 = oNode_medarb.selectSingleNode("txt_23").Text
        medarb_txt_024 = oNode_medarb.selectSingleNode("txt_24").Text
    
        medarb_txt_025 = oNode_medarb.selectSingleNode("txt_25").Text
        medarb_txt_026 = oNode_medarb.selectSingleNode("txt_26").Text
        medarb_txt_027 = oNode_medarb.selectSingleNode("txt_27").Text
        medarb_txt_028 = oNode_medarb.selectSingleNode("txt_28").Text
        medarb_txt_029 = oNode_medarb.selectSingleNode("txt_29").Text
        medarb_txt_030 = oNode_medarb.selectSingleNode("txt_30").Text
        medarb_txt_031 = oNode_medarb.selectSingleNode("txt_31").Text
        medarb_txt_032 = oNode_medarb.selectSingleNode("txt_32").Text
        medarb_txt_033 = oNode_medarb.selectSingleNode("txt_33").Text
        medarb_txt_034 = oNode_medarb.selectSingleNode("txt_34").Text
    
        medarb_txt_035 = oNode_medarb.selectSingleNode("txt_35").Text
        medarb_txt_036 = oNode_medarb.selectSingleNode("txt_36").Text
        medarb_txt_037 = oNode_medarb.selectSingleNode("txt_37").Text
       
        medarb_txt_038 = oNode_medarb.selectSingleNode("txt_38").Text
        medarb_txt_039 = oNode_medarb.selectSingleNode("txt_39").Text
        medarb_txt_040 = oNode_medarb.selectSingleNode("txt_40").Text
        medarb_txt_041 = oNode_medarb.selectSingleNode("txt_41").Text
        medarb_txt_042 = oNode_medarb.selectSingleNode("txt_42").Text
        medarb_txt_043 = oNode_medarb.selectSingleNode("txt_43").Text
        medarb_txt_044 = oNode_medarb.selectSingleNode("txt_44").Text
          
        medarb_txt_045 = oNode_medarb.selectSingleNode("txt_45").Text
        medarb_txt_046 = oNode_medarb.selectSingleNode("txt_46").Text
        medarb_txt_047 = oNode_medarb.selectSingleNode("txt_47").Text
        medarb_txt_048 = oNode_medarb.selectSingleNode("txt_48").Text
        medarb_txt_049 = oNode_medarb.selectSingleNode("txt_49").Text
        medarb_txt_050 = oNode_medarb.selectSingleNode("txt_50").Text
          
        medarb_txt_051 = oNode_medarb.selectSingleNode("txt_51").Text
        medarb_txt_052 = oNode_medarb.selectSingleNode("txt_52").Text
        medarb_txt_053 = oNode_medarb.selectSingleNode("txt_53").Text
        medarb_txt_054 = oNode_medarb.selectSingleNode("txt_54").Text
        medarb_txt_055 = oNode_medarb.selectSingleNode("txt_55").Text
        medarb_txt_056 = oNode_medarb.selectSingleNode("txt_56").Text
        medarb_txt_057 = oNode_medarb.selectSingleNode("txt_57").Text
        medarb_txt_058 = oNode_medarb.selectSingleNode("txt_58").Text
        medarb_txt_059 = oNode_medarb.selectSingleNode("txt_59").Text
        medarb_txt_060 = oNode_medarb.selectSingleNode("txt_60").Text
        medarb_txt_061 = oNode_medarb.selectSingleNode("txt_61").Text
        medarb_txt_062 = oNode_medarb.selectSingleNode("txt_62").Text
        medarb_txt_063 = oNode_medarb.selectSingleNode("txt_63").Text
        medarb_txt_064 = oNode_medarb.selectSingleNode("txt_64").Text
        medarb_txt_065 = oNode_medarb.selectSingleNode("txt_65").Text
        medarb_txt_066 = oNode_medarb.selectSingleNode("txt_66").Text
          
        medarb_txt_067 = oNode_medarb.selectSingleNode("txt_67").Text
        medarb_txt_068 = oNode_medarb.selectSingleNode("txt_68").Text
        medarb_txt_069 = oNode_medarb.selectSingleNode("txt_69").Text
        medarb_txt_070 = oNode_medarb.selectSingleNode("txt_70").Text

        medarb_txt_071 = oNode_medarb.selectSingleNode("txt_71").Text
        medarb_txt_072 = oNode_medarb.selectSingleNode("txt_72").Text
        medarb_txt_073 = oNode_medarb.selectSingleNode("txt_73").Text
        medarb_txt_074 = oNode_medarb.selectSingleNode("txt_74").Text
        medarb_txt_075 = oNode_medarb.selectSingleNode("txt_75").Text
        medarb_txt_076 = oNode_medarb.selectSingleNode("txt_76").Text
        medarb_txt_077 = oNode_medarb.selectSingleNode("txt_77").Text
        medarb_txt_078 = oNode_medarb.selectSingleNode("txt_78").Text
        medarb_txt_079 = oNode_medarb.selectSingleNode("txt_79").Text

        medarb_txt_080 = oNode_medarb.selectSingleNode("txt_80").Text
        medarb_txt_081 = oNode_medarb.selectSingleNode("txt_81").Text
        medarb_txt_082 = oNode_medarb.selectSingleNode("txt_82").Text
        medarb_txt_083 = oNode_medarb.selectSingleNode("txt_83").Text
          
        medarb_txt_084 = oNode_medarb.selectSingleNode("txt_84").Text
        medarb_txt_085 = oNode_medarb.selectSingleNode("txt_85").Text
        medarb_txt_086 = oNode_medarb.selectSingleNode("txt_86").Text
          
        medarb_txt_087 = oNode_medarb.selectSingleNode("txt_87").Text
          
        medarb_txt_088 = oNode_medarb.selectSingleNode("txt_88").Text
        medarb_txt_089 = oNode_medarb.selectSingleNode("txt_89").Text
        medarb_txt_090 = oNode_medarb.selectSingleNode("txt_90").Text
          
        medarb_txt_091 = oNode_medarb.selectSingleNode("txt_91").Text
        medarb_txt_092 = oNode_medarb.selectSingleNode("txt_92").Text
        medarb_txt_093 = oNode_medarb.selectSingleNode("txt_93").Text
        medarb_txt_094 = oNode_medarb.selectSingleNode("txt_94").Text
        medarb_txt_095 = oNode_medarb.selectSingleNode("txt_95").Text
        medarb_txt_096 = oNode_medarb.selectSingleNode("txt_96").Text
        medarb_txt_097 = oNode_medarb.selectSingleNode("txt_97").Text
        medarb_txt_098 = oNode_medarb.selectSingleNode("txt_98").Text
        medarb_txt_099 = oNode_medarb.selectSingleNode("txt_99").Text
          
        medarb_txt_100 = oNode_medarb.selectSingleNode("txt_100").Text
        medarb_txt_101 = oNode_medarb.selectSingleNode("txt_101").Text
        medarb_txt_102 = oNode_medarb.selectSingleNode("txt_102").Text
        medarb_txt_103 = oNode_medarb.selectSingleNode("txt_103").Text
          
        medarb_txt_104 = oNode_medarb.selectSingleNode("txt_104").Text
        medarb_txt_105 = oNode_medarb.selectSingleNode("txt_105").Text
        medarb_txt_106 = oNode_medarb.selectSingleNode("txt_106").Text
        medarb_txt_107 = oNode_medarb.selectSingleNode("txt_107").Text
        medarb_txt_108 = oNode_medarb.selectSingleNode("txt_108").Text
        medarb_txt_109 = oNode_medarb.selectSingleNode("txt_109").Text
        medarb_txt_110 = oNode_medarb.selectSingleNode("txt_110").Text
        medarb_txt_111 = oNode_medarb.selectSingleNode("txt_111").Text
        medarb_txt_112 = oNode_medarb.selectSingleNode("txt_112").Text
        medarb_txt_113 = oNode_medarb.selectSingleNode("txt_113").Text
        medarb_txt_114 = oNode_medarb.selectSingleNode("txt_114").Text
        medarb_txt_115 = oNode_medarb.selectSingleNode("txt_115").Text
        medarb_txt_116 = oNode_medarb.selectSingleNode("txt_116").Text
        medarb_txt_117 = oNode_medarb.selectSingleNode("txt_117").Text
        medarb_txt_118 = oNode_medarb.selectSingleNode("txt_118").Text
        medarb_txt_119 = oNode_medarb.selectSingleNode("txt_119").Text
        medarb_txt_120 = oNode_medarb.selectSingleNode("txt_120").Text
        medarb_txt_121 = oNode_medarb.selectSingleNode("txt_121").Text
        medarb_txt_122 = oNode_medarb.selectSingleNode("txt_122").Text
        medarb_txt_123 = oNode_medarb.selectSingleNode("txt_123").Text
        medarb_txt_124 = oNode_medarb.selectSingleNode("txt_124").Text
        medarb_txt_125 = oNode_medarb.selectSingleNode("txt_125").Text
        medarb_txt_126 = oNode_medarb.selectSingleNode("txt_126").Text
        medarb_txt_127 = oNode_medarb.selectSingleNode("txt_127").Text
        medarb_txt_128 = oNode_medarb.selectSingleNode("txt_128").Text
        medarb_txt_129 = oNode_medarb.selectSingleNode("txt_129").Text
        medarb_txt_130 = oNode_medarb.selectSingleNode("txt_130").Text
        medarb_txt_131 = oNode_medarb.selectSingleNode("txt_131").Text
        medarb_txt_132 = oNode_medarb.selectSingleNode("txt_132").Text
        medarb_txt_133 = oNode_medarb.selectSingleNode("txt_133").Text
        medarb_txt_134 = oNode_medarb.selectSingleNode("txt_134").Text
        medarb_txt_135 = oNode_medarb.selectSingleNode("txt_135").Text
        medarb_txt_136 = oNode_medarb.selectSingleNode("txt_136").Text
        medarb_txt_137 = oNode_medarb.selectSingleNode("txt_137").Text
        medarb_txt_138 = oNode_medarb.selectSingleNode("txt_138").Text
        medarb_txt_139 = oNode_medarb.selectSingleNode("txt_139").Text
        medarb_txt_140 = oNode_medarb.selectSingleNode("txt_140").Text
        medarb_txt_141 = oNode_medarb.selectSingleNode("txt_141").Text
        medarb_txt_142 = oNode_medarb.selectSingleNode("txt_142").Text
        medarb_txt_143 = oNode_medarb.selectSingleNode("txt_143").Text
        medarb_txt_144 = oNode_medarb.selectSingleNode("txt_144").Text

        medarb_txt_145 = oNode_medarb.selectSingleNode("txt_145").Text
        medarb_txt_146 = oNode_medarb.selectSingleNode("txt_146").Text
        medarb_txt_147 = oNode_medarb.selectSingleNode("txt_147").Text
        medarb_txt_148 = oNode_medarb.selectSingleNode("txt_148").Text
        medarb_txt_149 = oNode_medarb.selectSingleNode("txt_149").Text
        medarb_txt_150 = oNode_medarb.selectSingleNode("txt_150").Text
        medarb_txt_151 = oNode_medarb.selectSingleNode("txt_151").Text
        medarb_txt_152 = oNode_medarb.selectSingleNode("txt_152").Text
        medarb_txt_153 = oNode_medarb.selectSingleNode("txt_153").Text
        medarb_txt_154 = oNode_medarb.selectSingleNode("txt_154").Text
        medarb_txt_155 = oNode_medarb.selectSingleNode("txt_155").Text
        medarb_txt_156 = oNode_medarb.selectSingleNode("txt_156").Text
        medarb_txt_157 = oNode_medarb.selectSingleNode("txt_157").Text

    next




'Response.Write "week_txt_001: " & week_txt_001 & "<br>"
'Response.Write "week_txt_002: " & week_txt_002 & "<br>"


%>