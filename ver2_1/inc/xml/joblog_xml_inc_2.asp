
<% 
Dim objXMLHTTP_joblog2, objXMLDOM_joblog2, i_joblog2, strHTML_joblog2

Set objXMLDom_joblog2 = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_joblog2 = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_joblog2.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/joblog_sprog_2.xml", False
'objXmlHttp_joblog2.open "GET", "http://localhost/inc/xml/joblog_sprog.xml", False
'objXmlHttp_joblog2.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/joblog_sprog_2.xml", False
'objXmlHttp_joblog2.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/joblog_sprog_2.xml", False
'objXmlHttp_joblog2.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/joblog_sprog_2.xml", False
objXmlHttp_joblog2.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/joblog_sprog_2.xml", False

objXmlHttp_joblog2.send


Set objXmlDom_joblog2 = objXmlHttp_joblog2.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_joblog2 = Nothing



Dim Address_joblog2, Latitude_joblog2, Longitude_joblog2
Dim oNode_joblog2, oNodes_joblog2
Dim sXPathQuery_joblog2

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
sXPathQuery_joblog2 = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_joblog2 = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_joblog2 = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_joblog2 = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_joblog2 = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_joblog2 = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_joblog2 = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_joblog2 = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_joblog2 = objXmlDom_joblog2.documentElement.selectSingleNode(sXPathQuery_joblog2)
Address_joblog2 = oNode_joblog2.Text

Set oNodes_joblog2 = objXmlDom_joblog2.documentElement.selectNodes(sXPathQuery_joblog2)

    For Each oNode_joblog2 in oNodes_joblog2

    joblog2_txt_001 = oNode_joblog2.selectSingleNode("txt_1").Text
    joblog2_txt_002 = oNode_joblog2.selectSingleNode("txt_2").Text
    joblog2_txt_003 = oNode_joblog2.selectSingleNode("txt_3").Text
    joblog2_txt_004 = oNode_joblog2.selectSingleNode("txt_4").Text
    joblog2_txt_005 = oNode_joblog2.selectSingleNode("txt_5").Text
    joblog2_txt_006 = oNode_joblog2.selectSingleNode("txt_6").Text
    joblog2_txt_007 = oNode_joblog2.selectSingleNode("txt_7").Text
    joblog2_txt_008 = oNode_joblog2.selectSingleNode("txt_8").Text
    joblog2_txt_009 = oNode_joblog2.selectSingleNode("txt_9").Text
    joblog2_txt_010 = oNode_joblog2.selectSingleNode("txt_10").Text

    joblog2_txt_011 = oNode_joblog2.selectSingleNode("txt_11").Text
    joblog2_txt_012 = oNode_joblog2.selectSingleNode("txt_12").Text
    joblog2_txt_013 = oNode_joblog2.selectSingleNode("txt_13").Text
    joblog2_txt_014 = oNode_joblog2.selectSingleNode("txt_14").Text
    joblog2_txt_015 = oNode_joblog2.selectSingleNode("txt_15").Text
    joblog2_txt_016 = oNode_joblog2.selectSingleNode("txt_16").Text
    joblog2_txt_017 = oNode_joblog2.selectSingleNode("txt_17").Text
    joblog2_txt_018 = oNode_joblog2.selectSingleNode("txt_18").Text
    joblog2_txt_019 = oNode_joblog2.selectSingleNode("txt_19").Text
    joblog2_txt_020 = oNode_joblog2.selectSingleNode("txt_20").Text

    joblog2_txt_021 = oNode_joblog2.selectSingleNode("txt_21").Text
    joblog2_txt_022 = oNode_joblog2.selectSingleNode("txt_22").Text
    joblog2_txt_023 = oNode_joblog2.selectSingleNode("txt_23").Text
    joblog2_txt_024 = oNode_joblog2.selectSingleNode("txt_24").Text
    joblog2_txt_025 = oNode_joblog2.selectSingleNode("txt_25").Text
    joblog2_txt_026 = oNode_joblog2.selectSingleNode("txt_26").Text
    joblog2_txt_027 = oNode_joblog2.selectSingleNode("txt_27").Text
    joblog2_txt_028 = oNode_joblog2.selectSingleNode("txt_28").Text
    joblog2_txt_029 = oNode_joblog2.selectSingleNode("txt_29").Text
    joblog2_txt_030 = oNode_joblog2.selectSingleNode("txt_30").Text

    joblog2_txt_031 = oNode_joblog2.selectSingleNode("txt_31").Text
    joblog2_txt_032 = oNode_joblog2.selectSingleNode("txt_32").Text
    joblog2_txt_033 = oNode_joblog2.selectSingleNode("txt_33").Text
    joblog2_txt_034 = oNode_joblog2.selectSingleNode("txt_34").Text
    joblog2_txt_035 = oNode_joblog2.selectSingleNode("txt_35").Text
    joblog2_txt_036 = oNode_joblog2.selectSingleNode("txt_36").Text
    joblog2_txt_037 = oNode_joblog2.selectSingleNode("txt_37").Text
    joblog2_txt_038 = oNode_joblog2.selectSingleNode("txt_38").Text
    joblog2_txt_039 = oNode_joblog2.selectSingleNode("txt_39").Text
    joblog2_txt_040 = oNode_joblog2.selectSingleNode("txt_40").Text

    joblog2_txt_041 = oNode_joblog2.selectSingleNode("txt_41").Text
    joblog2_txt_042 = oNode_joblog2.selectSingleNode("txt_42").Text
    joblog2_txt_043 = oNode_joblog2.selectSingleNode("txt_43").Text
    joblog2_txt_044 = oNode_joblog2.selectSingleNode("txt_44").Text
    joblog2_txt_045 = oNode_joblog2.selectSingleNode("txt_45").Text
    joblog2_txt_046 = oNode_joblog2.selectSingleNode("txt_46").Text
    joblog2_txt_047 = oNode_joblog2.selectSingleNode("txt_47").Text
    joblog2_txt_048 = oNode_joblog2.selectSingleNode("txt_48").Text
    joblog2_txt_049 = oNode_joblog2.selectSingleNode("txt_49").Text
    joblog2_txt_050 = oNode_joblog2.selectSingleNode("txt_50").Text

    joblog2_txt_051 = oNode_joblog2.selectSingleNode("txt_51").Text
    joblog2_txt_052 = oNode_joblog2.selectSingleNode("txt_52").Text
    joblog2_txt_053 = oNode_joblog2.selectSingleNode("txt_53").Text
    joblog2_txt_054 = oNode_joblog2.selectSingleNode("txt_54").Text
    joblog2_txt_055 = oNode_joblog2.selectSingleNode("txt_55").Text
    joblog2_txt_056 = oNode_joblog2.selectSingleNode("txt_56").Text
    joblog2_txt_057 = oNode_joblog2.selectSingleNode("txt_57").Text
    joblog2_txt_058 = oNode_joblog2.selectSingleNode("txt_58").Text
    joblog2_txt_059 = oNode_joblog2.selectSingleNode("txt_59").Text
    joblog2_txt_060 = oNode_joblog2.selectSingleNode("txt_60").Text

    joblog2_txt_061 = oNode_joblog2.selectSingleNode("txt_61").Text
    joblog2_txt_062 = oNode_joblog2.selectSingleNode("txt_62").Text
    joblog2_txt_063 = oNode_joblog2.selectSingleNode("txt_63").Text
    joblog2_txt_064 = oNode_joblog2.selectSingleNode("txt_64").Text
    joblog2_txt_065 = oNode_joblog2.selectSingleNode("txt_65").Text
    joblog2_txt_066 = oNode_joblog2.selectSingleNode("txt_66").Text
    joblog2_txt_067 = oNode_joblog2.selectSingleNode("txt_67").Text
    joblog2_txt_068 = oNode_joblog2.selectSingleNode("txt_68").Text
    joblog2_txt_069 = oNode_joblog2.selectSingleNode("txt_69").Text
    joblog2_txt_070 = oNode_joblog2.selectSingleNode("txt_70").Text

    joblog2_txt_071 = oNode_joblog2.selectSingleNode("txt_71").Text
    joblog2_txt_072 = oNode_joblog2.selectSingleNode("txt_72").Text
    joblog2_txt_073 = oNode_joblog2.selectSingleNode("txt_73").Text
    joblog2_txt_074 = oNode_joblog2.selectSingleNode("txt_74").Text
    joblog2_txt_075 = oNode_joblog2.selectSingleNode("txt_75").Text
    joblog2_txt_076 = oNode_joblog2.selectSingleNode("txt_76").Text
    joblog2_txt_077 = oNode_joblog2.selectSingleNode("txt_77").Text
    joblog2_txt_078 = oNode_joblog2.selectSingleNode("txt_78").Text
    joblog2_txt_079 = oNode_joblog2.selectSingleNode("txt_79").Text
    joblog2_txt_080 = oNode_joblog2.selectSingleNode("txt_80").Text

    joblog2_txt_081 = oNode_joblog2.selectSingleNode("txt_81").Text
    joblog2_txt_082 = oNode_joblog2.selectSingleNode("txt_82").Text
    joblog2_txt_083 = oNode_joblog2.selectSingleNode("txt_83").Text
    joblog2_txt_084 = oNode_joblog2.selectSingleNode("txt_84").Text
    joblog2_txt_085 = oNode_joblog2.selectSingleNode("txt_85").Text
    joblog2_txt_086 = oNode_joblog2.selectSingleNode("txt_86").Text
    joblog2_txt_087 = oNode_joblog2.selectSingleNode("txt_87").Text
    joblog2_txt_088 = oNode_joblog2.selectSingleNode("txt_88").Text
    joblog2_txt_089 = oNode_joblog2.selectSingleNode("txt_89").Text
    joblog2_txt_090 = oNode_joblog2.selectSingleNode("txt_90").Text

    joblog2_txt_091 = oNode_joblog2.selectSingleNode("txt_91").Text
    joblog2_txt_092 = oNode_joblog2.selectSingleNode("txt_92").Text
    joblog2_txt_093 = oNode_joblog2.selectSingleNode("txt_93").Text
    joblog2_txt_094 = oNode_joblog2.selectSingleNode("txt_94").Text
    joblog2_txt_095 = oNode_joblog2.selectSingleNode("txt_95").Text
    joblog2_txt_096 = oNode_joblog2.selectSingleNode("txt_96").Text
    joblog2_txt_097 = oNode_joblog2.selectSingleNode("txt_97").Text
    joblog2_txt_098 = oNode_joblog2.selectSingleNode("txt_98").Text
    joblog2_txt_099 = oNode_joblog2.selectSingleNode("txt_99").Text
    joblog2_txt_100 = oNode_joblog2.selectSingleNode("txt_100").Text

    joblog2_txt_101 = oNode_joblog2.selectSingleNode("txt_101").Text
    joblog2_txt_102 = oNode_joblog2.selectSingleNode("txt_102").Text
    joblog2_txt_103 = oNode_joblog2.selectSingleNode("txt_103").Text
    joblog2_txt_104 = oNode_joblog2.selectSingleNode("txt_104").Text
    joblog2_txt_105 = oNode_joblog2.selectSingleNode("txt_105").Text
    joblog2_txt_106 = oNode_joblog2.selectSingleNode("txt_106").Text
    joblog2_txt_107 = oNode_joblog2.selectSingleNode("txt_107").Text
    joblog2_txt_108 = oNode_joblog2.selectSingleNode("txt_108").Text
    joblog2_txt_109 = oNode_joblog2.selectSingleNode("txt_109").Text
    joblog2_txt_110 = oNode_joblog2.selectSingleNode("txt_110").Text

    joblog2_txt_111 = oNode_joblog2.selectSingleNode("txt_111").Text
    joblog2_txt_112 = oNode_joblog2.selectSingleNode("txt_112").Text
    joblog2_txt_113 = oNode_joblog2.selectSingleNode("txt_113").Text
    joblog2_txt_114 = oNode_joblog2.selectSingleNode("txt_114").Text
    joblog2_txt_115 = oNode_joblog2.selectSingleNode("txt_115").Text
    joblog2_txt_116 = oNode_joblog2.selectSingleNode("txt_116").Text
    joblog2_txt_117 = oNode_joblog2.selectSingleNode("txt_117").Text

    joblog2_txt_118 = oNode_joblog2.selectSingleNode("txt_118").Text
    joblog2_txt_119 = oNode_joblog2.selectSingleNode("txt_119").Text
    joblog2_txt_120 = oNode_joblog2.selectSingleNode("txt_120").Text
    joblog2_txt_121 = oNode_joblog2.selectSingleNode("txt_121").Text
    joblog2_txt_122 = oNode_joblog2.selectSingleNode("txt_122").Text
    joblog2_txt_123 = oNode_joblog2.selectSingleNode("txt_123").Text
    joblog2_txt_124 = oNode_joblog2.selectSingleNode("txt_124").Text
    joblog2_txt_125 = oNode_joblog2.selectSingleNode("txt_125").Text
    joblog2_txt_126 = oNode_joblog2.selectSingleNode("txt_126").Text
    joblog2_txt_127 = oNode_joblog2.selectSingleNode("txt_127").Text
    joblog2_txt_128 = oNode_joblog2.selectSingleNode("txt_128").Text
    joblog2_txt_129 = oNode_joblog2.selectSingleNode("txt_129").Text

    joblog2_txt_130 = oNode_joblog2.selectSingleNode("txt_130").Text
    joblog2_txt_131 = oNode_joblog2.selectSingleNode("txt_131").Text
    joblog2_txt_132 = oNode_joblog2.selectSingleNode("txt_132").Text
    joblog2_txt_133 = oNode_joblog2.selectSingleNode("txt_133").Text
    joblog2_txt_134 = oNode_joblog2.selectSingleNode("txt_134").Text
    joblog2_txt_135 = oNode_joblog2.selectSingleNode("txt_135").Text
    joblog2_txt_136 = oNode_joblog2.selectSingleNode("txt_136").Text

    joblog2_txt_137 = oNode_joblog2.selectSingleNode("txt_137").Text
    joblog2_txt_138 = oNode_joblog2.selectSingleNode("txt_138").Text
    joblog2_txt_139 = oNode_joblog2.selectSingleNode("txt_139").Text
    joblog2_txt_140 = oNode_joblog2.selectSingleNode("txt_140").Text
    joblog2_txt_141 = oNode_joblog2.selectSingleNode("txt_141").Text
    joblog2_txt_142 = oNode_joblog2.selectSingleNode("txt_142").Text

    joblog2_txt_143 = oNode_joblog2.selectSingleNode("txt_143").Text
    joblog2_txt_144 = oNode_joblog2.selectSingleNode("txt_144").Text
    joblog2_txt_145 = oNode_joblog2.selectSingleNode("txt_145").Text
    joblog2_txt_146 = oNode_joblog2.selectSingleNode("txt_146").Text
    joblog2_txt_147 = oNode_joblog2.selectSingleNode("txt_147").Text
    joblog2_txt_148 = oNode_joblog2.selectSingleNode("txt_148").Text
    joblog2_txt_149 = oNode_joblog2.selectSingleNode("txt_149").Text
    joblog2_txt_150 = oNode_joblog2.selectSingleNode("txt_150").Text

    joblog2_txt_151 = oNode_joblog2.selectSingleNode("txt_151").Text
    joblog2_txt_152 = oNode_joblog2.selectSingleNode("txt_152").Text
    joblog2_txt_153 = oNode_joblog2.selectSingleNode("txt_153").Text
    joblog2_txt_154 = oNode_joblog2.selectSingleNode("txt_154").Text
    joblog2_txt_155 = oNode_joblog2.selectSingleNode("txt_155").Text
    joblog2_txt_156 = oNode_joblog2.selectSingleNode("txt_156").Text
    joblog2_txt_157 = oNode_joblog2.selectSingleNode("txt_157").Text
    joblog2_txt_158 = oNode_joblog2.selectSingleNode("txt_158").Text
    joblog2_txt_159 = oNode_joblog2.selectSingleNode("txt_159").Text
    joblog2_txt_160 = oNode_joblog2.selectSingleNode("txt_160").Text

    joblog2_txt_161 = oNode_joblog2.selectSingleNode("txt_161").Text
    joblog2_txt_162 = oNode_joblog2.selectSingleNode("txt_162").Text
    joblog2_txt_163 = oNode_joblog2.selectSingleNode("txt_163").Text
    joblog2_txt_164 = oNode_joblog2.selectSingleNode("txt_164").Text
    joblog2_txt_165 = oNode_joblog2.selectSingleNode("txt_165").Text
    joblog2_txt_166 = oNode_joblog2.selectSingleNode("txt_166").Text
    joblog2_txt_167 = oNode_joblog2.selectSingleNode("txt_167").Text
    joblog2_txt_168 = oNode_joblog2.selectSingleNode("txt_168").Text
    joblog2_txt_169 = oNode_joblog2.selectSingleNode("txt_169").Text
    joblog2_txt_170 = oNode_joblog2.selectSingleNode("txt_170").Text


      
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>