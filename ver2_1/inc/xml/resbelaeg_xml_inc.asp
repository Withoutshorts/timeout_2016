
<% 
Dim objXMLHTTP_resbelaeg, objXMLDOM_resbelaeg, i_resbelaeg, strHTML_resbelaeg

Set objXMLDom_resbelaeg = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_resbelaeg = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_resbelaeg.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/resbelaeg_sprog.xml", False
'objXmlHttp_resbelaeg.open "GET", "http://localhost/inc/xml/resbelaeg_sprog.xml", False
'objXmlHttp_resbelaeg.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/resbelaeg_sprog.xml", False
'objXmlHttp_resbelaeg.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/resbelaeg_sprog.xml", False
objXmlHttp_resbelaeg.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/resbelaeg_sprog.xml", False
'objXmlHttp_resbelaeg.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/resbelaeg_sprog.xml", False
'objXmlHttp_resbelaeg.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/resbelaeg_sprog.xml", False

objXmlHttp_resbelaeg.send


Set objXmlDom_resbelaeg = objXmlHttp_resbelaeg.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_resbelaeg = Nothing



Dim Address_resbelaeg, Latitude_resbelaeg, Longitude_resbelaeg
Dim oNode_resbelaeg, oNodes_resbelaeg
Dim sXPathQuery_resbelaeg

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
sXPathQuery_resbelaeg = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_resbelaeg = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_resbelaeg = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_resbelaeg = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_resbelaeg = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_resbelaeg = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_resbelaeg = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_resbelaeg = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_resbelaeg = objXmlDom_resbelaeg.documentElement.selectSingleNode(sXPathQuery_resbelaeg)
Address_resbelaeg = oNode_resbelaeg.Text

Set oNodes_resbelaeg = objXmlDom_resbelaeg.documentElement.selectNodes(sXPathQuery_resbelaeg)

    For Each oNode_resbelaeg in oNodes_resbelaeg

    resbelaeg_txt_001 = oNode_resbelaeg.selectSingleNode("txt_1").Text
    resbelaeg_txt_002 = oNode_resbelaeg.selectSingleNode("txt_2").Text
    resbelaeg_txt_003 = oNode_resbelaeg.selectSingleNode("txt_3").Text
    resbelaeg_txt_004 = oNode_resbelaeg.selectSingleNode("txt_4").Text
    resbelaeg_txt_005 = oNode_resbelaeg.selectSingleNode("txt_5").Text
    resbelaeg_txt_006 = oNode_resbelaeg.selectSingleNode("txt_6").Text
    resbelaeg_txt_007 = oNode_resbelaeg.selectSingleNode("txt_7").Text
    resbelaeg_txt_008 = oNode_resbelaeg.selectSingleNode("txt_8").Text
    resbelaeg_txt_009 = oNode_resbelaeg.selectSingleNode("txt_9").Text
    resbelaeg_txt_010 = oNode_resbelaeg.selectSingleNode("txt_10").Text

    resbelaeg_txt_011 = oNode_resbelaeg.selectSingleNode("txt_11").Text
    resbelaeg_txt_012 = oNode_resbelaeg.selectSingleNode("txt_12").Text
    resbelaeg_txt_013 = oNode_resbelaeg.selectSingleNode("txt_13").Text
    resbelaeg_txt_014 = oNode_resbelaeg.selectSingleNode("txt_14").Text
    resbelaeg_txt_015 = oNode_resbelaeg.selectSingleNode("txt_15").Text
    resbelaeg_txt_016 = oNode_resbelaeg.selectSingleNode("txt_16").Text
    resbelaeg_txt_017 = oNode_resbelaeg.selectSingleNode("txt_17").Text
    resbelaeg_txt_018 = oNode_resbelaeg.selectSingleNode("txt_18").Text
    resbelaeg_txt_019 = oNode_resbelaeg.selectSingleNode("txt_19").Text
    resbelaeg_txt_020 = oNode_resbelaeg.selectSingleNode("txt_20").Text

    resbelaeg_txt_021 = oNode_resbelaeg.selectSingleNode("txt_21").Text
    resbelaeg_txt_022 = oNode_resbelaeg.selectSingleNode("txt_22").Text
    resbelaeg_txt_023 = oNode_resbelaeg.selectSingleNode("txt_23").Text
    resbelaeg_txt_024 = oNode_resbelaeg.selectSingleNode("txt_24").Text
    resbelaeg_txt_025 = oNode_resbelaeg.selectSingleNode("txt_25").Text
    resbelaeg_txt_026 = oNode_resbelaeg.selectSingleNode("txt_26").Text
    resbelaeg_txt_027 = oNode_resbelaeg.selectSingleNode("txt_27").Text
    resbelaeg_txt_028 = oNode_resbelaeg.selectSingleNode("txt_28").Text
    resbelaeg_txt_029 = oNode_resbelaeg.selectSingleNode("txt_29").Text
    resbelaeg_txt_030 = oNode_resbelaeg.selectSingleNode("txt_30").Text

    resbelaeg_txt_031 = oNode_resbelaeg.selectSingleNode("txt_31").Text
    resbelaeg_txt_032 = oNode_resbelaeg.selectSingleNode("txt_32").Text
    resbelaeg_txt_033 = oNode_resbelaeg.selectSingleNode("txt_33").Text
    resbelaeg_txt_034 = oNode_resbelaeg.selectSingleNode("txt_34").Text
    resbelaeg_txt_035 = oNode_resbelaeg.selectSingleNode("txt_35").Text
    resbelaeg_txt_036 = oNode_resbelaeg.selectSingleNode("txt_36").Text
    resbelaeg_txt_037 = oNode_resbelaeg.selectSingleNode("txt_37").Text
    resbelaeg_txt_038 = oNode_resbelaeg.selectSingleNode("txt_38").Text
    resbelaeg_txt_039 = oNode_resbelaeg.selectSingleNode("txt_39").Text
    resbelaeg_txt_040 = oNode_resbelaeg.selectSingleNode("txt_40").Text

    resbelaeg_txt_041 = oNode_resbelaeg.selectSingleNode("txt_41").Text
    resbelaeg_txt_042 = oNode_resbelaeg.selectSingleNode("txt_42").Text
    resbelaeg_txt_043 = oNode_resbelaeg.selectSingleNode("txt_43").Text
    resbelaeg_txt_044 = oNode_resbelaeg.selectSingleNode("txt_44").Text
    resbelaeg_txt_045 = oNode_resbelaeg.selectSingleNode("txt_45").Text
    resbelaeg_txt_046 = oNode_resbelaeg.selectSingleNode("txt_46").Text
    resbelaeg_txt_047 = oNode_resbelaeg.selectSingleNode("txt_47").Text
    resbelaeg_txt_048 = oNode_resbelaeg.selectSingleNode("txt_48").Text
    resbelaeg_txt_049 = oNode_resbelaeg.selectSingleNode("txt_49").Text
    resbelaeg_txt_050 = oNode_resbelaeg.selectSingleNode("txt_50").Text

    resbelaeg_txt_051 = oNode_resbelaeg.selectSingleNode("txt_51").Text
    resbelaeg_txt_052 = oNode_resbelaeg.selectSingleNode("txt_52").Text
    resbelaeg_txt_053 = oNode_resbelaeg.selectSingleNode("txt_53").Text
    resbelaeg_txt_054 = oNode_resbelaeg.selectSingleNode("txt_54").Text
    resbelaeg_txt_055 = oNode_resbelaeg.selectSingleNode("txt_55").Text
    resbelaeg_txt_056 = oNode_resbelaeg.selectSingleNode("txt_56").Text
    resbelaeg_txt_057 = oNode_resbelaeg.selectSingleNode("txt_57").Text
    resbelaeg_txt_058 = oNode_resbelaeg.selectSingleNode("txt_58").Text
    resbelaeg_txt_059 = oNode_resbelaeg.selectSingleNode("txt_59").Text
    resbelaeg_txt_060 = oNode_resbelaeg.selectSingleNode("txt_60").Text

    resbelaeg_txt_061 = oNode_resbelaeg.selectSingleNode("txt_61").Text
    resbelaeg_txt_062 = oNode_resbelaeg.selectSingleNode("txt_62").Text
    resbelaeg_txt_063 = oNode_resbelaeg.selectSingleNode("txt_63").Text
    resbelaeg_txt_064 = oNode_resbelaeg.selectSingleNode("txt_64").Text
    resbelaeg_txt_065 = oNode_resbelaeg.selectSingleNode("txt_65").Text
    resbelaeg_txt_066 = oNode_resbelaeg.selectSingleNode("txt_66").Text
    resbelaeg_txt_067 = oNode_resbelaeg.selectSingleNode("txt_67").Text
    resbelaeg_txt_068 = oNode_resbelaeg.selectSingleNode("txt_68").Text
    resbelaeg_txt_069 = oNode_resbelaeg.selectSingleNode("txt_69").Text
    resbelaeg_txt_070 = oNode_resbelaeg.selectSingleNode("txt_70").Text

    resbelaeg_txt_071 = oNode_resbelaeg.selectSingleNode("txt_71").Text
    resbelaeg_txt_072 = oNode_resbelaeg.selectSingleNode("txt_72").Text
    resbelaeg_txt_073 = oNode_resbelaeg.selectSingleNode("txt_73").Text
    resbelaeg_txt_074 = oNode_resbelaeg.selectSingleNode("txt_74").Text
    resbelaeg_txt_075 = oNode_resbelaeg.selectSingleNode("txt_75").Text
    resbelaeg_txt_076 = oNode_resbelaeg.selectSingleNode("txt_76").Text
    resbelaeg_txt_077 = oNode_resbelaeg.selectSingleNode("txt_77").Text

    resbelaeg_txt_078 = oNode_resbelaeg.selectSingleNode("txt_78").Text
    resbelaeg_txt_079 = oNode_resbelaeg.selectSingleNode("txt_79").Text

    resbelaeg_txt_080 = oNode_resbelaeg.selectSingleNode("txt_80").Text
    resbelaeg_txt_081 = oNode_resbelaeg.selectSingleNode("txt_81").Text
    resbelaeg_txt_082 = oNode_resbelaeg.selectSingleNode("txt_82").Text
    resbelaeg_txt_083 = oNode_resbelaeg.selectSingleNode("txt_83").Text
    resbelaeg_txt_084 = oNode_resbelaeg.selectSingleNode("txt_84").Text
    resbelaeg_txt_085 = oNode_resbelaeg.selectSingleNode("txt_85").Text
    resbelaeg_txt_086 = oNode_resbelaeg.selectSingleNode("txt_86").Text
    resbelaeg_txt_087 = oNode_resbelaeg.selectSingleNode("txt_87").Text
    resbelaeg_txt_088 = oNode_resbelaeg.selectSingleNode("txt_88").Text
    resbelaeg_txt_089 = oNode_resbelaeg.selectSingleNode("txt_89").Text

    resbelaeg_txt_090 = oNode_resbelaeg.selectSingleNode("txt_90").Text
    resbelaeg_txt_091 = oNode_resbelaeg.selectSingleNode("txt_91").Text
    resbelaeg_txt_092 = oNode_resbelaeg.selectSingleNode("txt_92").Text
    resbelaeg_txt_093 = oNode_resbelaeg.selectSingleNode("txt_93").Text
    resbelaeg_txt_094 = oNode_resbelaeg.selectSingleNode("txt_94").Text
    resbelaeg_txt_095 = oNode_resbelaeg.selectSingleNode("txt_95").Text
    resbelaeg_txt_096 = oNode_resbelaeg.selectSingleNode("txt_96").Text
    resbelaeg_txt_097 = oNode_resbelaeg.selectSingleNode("txt_97").Text
    resbelaeg_txt_098 = oNode_resbelaeg.selectSingleNode("txt_98").Text
    resbelaeg_txt_099 = oNode_resbelaeg.selectSingleNode("txt_99").Text
    resbelaeg_txt_100 = oNode_resbelaeg.selectSingleNode("txt_100").Text

    resbelaeg_txt_101 = oNode_resbelaeg.selectSingleNode("txt_101").Text
    resbelaeg_txt_102 = oNode_resbelaeg.selectSingleNode("txt_102").Text
    resbelaeg_txt_103 = oNode_resbelaeg.selectSingleNode("txt_103").Text
    resbelaeg_txt_104 = oNode_resbelaeg.selectSingleNode("txt_104").Text
    resbelaeg_txt_105 = oNode_resbelaeg.selectSingleNode("txt_105").Text
    resbelaeg_txt_106 = oNode_resbelaeg.selectSingleNode("txt_106").Text
    resbelaeg_txt_107 = oNode_resbelaeg.selectSingleNode("txt_107").Text
    resbelaeg_txt_108 = oNode_resbelaeg.selectSingleNode("txt_108").Text
    resbelaeg_txt_109 = oNode_resbelaeg.selectSingleNode("txt_109").Text
    resbelaeg_txt_110 = oNode_resbelaeg.selectSingleNode("txt_110").Text

    resbelaeg_txt_111 = oNode_resbelaeg.selectSingleNode("txt_111").Text
    resbelaeg_txt_112 = oNode_resbelaeg.selectSingleNode("txt_112").Text
    resbelaeg_txt_113 = oNode_resbelaeg.selectSingleNode("txt_113").Text
    resbelaeg_txt_114 = oNode_resbelaeg.selectSingleNode("txt_114").Text
    resbelaeg_txt_115 = oNode_resbelaeg.selectSingleNode("txt_115").Text
    resbelaeg_txt_116 = oNode_resbelaeg.selectSingleNode("txt_116").Text
    resbelaeg_txt_117 = oNode_resbelaeg.selectSingleNode("txt_117").Text
    resbelaeg_txt_118 = oNode_resbelaeg.selectSingleNode("txt_118").Text
    resbelaeg_txt_119 = oNode_resbelaeg.selectSingleNode("txt_119").Text

    resbelaeg_txt_120 = oNode_resbelaeg.selectSingleNode("txt_119").Text
    resbelaeg_txt_121 = oNode_resbelaeg.selectSingleNode("txt_121").Text
    resbelaeg_txt_122 = oNode_resbelaeg.selectSingleNode("txt_122").Text
    resbelaeg_txt_123 = oNode_resbelaeg.selectSingleNode("txt_123").Text
    resbelaeg_txt_124 = oNode_resbelaeg.selectSingleNode("txt_124").Text
    resbelaeg_txt_125 = oNode_resbelaeg.selectSingleNode("txt_125").Text
    resbelaeg_txt_126 = oNode_resbelaeg.selectSingleNode("txt_126").Text
    resbelaeg_txt_127 = oNode_resbelaeg.selectSingleNode("txt_127").Text
    resbelaeg_txt_128 = oNode_resbelaeg.selectSingleNode("txt_128").Text
    resbelaeg_txt_129 = oNode_resbelaeg.selectSingleNode("txt_129").Text
    resbelaeg_txt_130 = oNode_resbelaeg.selectSingleNode("txt_130").Text
    resbelaeg_txt_131 = oNode_resbelaeg.selectSingleNode("txt_131").Text
    resbelaeg_txt_132 = oNode_resbelaeg.selectSingleNode("txt_132").Text
    resbelaeg_txt_133 = oNode_resbelaeg.selectSingleNode("txt_133").Text
    resbelaeg_txt_134 = oNode_resbelaeg.selectSingleNode("txt_134").Text
    resbelaeg_txt_135 = oNode_resbelaeg.selectSingleNode("txt_135").Text
    resbelaeg_txt_136 = oNode_resbelaeg.selectSingleNode("txt_136").Text
    resbelaeg_txt_137 = oNode_resbelaeg.selectSingleNode("txt_137").Text
    resbelaeg_txt_138 = oNode_resbelaeg.selectSingleNode("txt_138").Text
    resbelaeg_txt_139 = oNode_resbelaeg.selectSingleNode("txt_139").Text
    resbelaeg_txt_140 = oNode_resbelaeg.selectSingleNode("txt_140").Text
    resbelaeg_txt_141 = oNode_resbelaeg.selectSingleNode("txt_141").Text
    resbelaeg_txt_142 = oNode_resbelaeg.selectSingleNode("txt_142").Text
    resbelaeg_txt_143 = oNode_resbelaeg.selectSingleNode("txt_143").Text
    resbelaeg_txt_144 = oNode_resbelaeg.selectSingleNode("txt_144").Text
    resbelaeg_txt_145 = oNode_resbelaeg.selectSingleNode("txt_145").Text
    resbelaeg_txt_146 = oNode_resbelaeg.selectSingleNode("txt_146").Text
    resbelaeg_txt_147 = oNode_resbelaeg.selectSingleNode("txt_147").Text
    resbelaeg_txt_148 = oNode_resbelaeg.selectSingleNode("txt_148").Text

    resbelaeg_txt_149 = oNode_resbelaeg.selectSingleNode("txt_149").Text
    resbelaeg_txt_150 = oNode_resbelaeg.selectSingleNode("txt_150").Text
    resbelaeg_txt_151 = oNode_resbelaeg.selectSingleNode("txt_151").Text
    resbelaeg_txt_152 = oNode_resbelaeg.selectSingleNode("txt_152").Text
    resbelaeg_txt_153 = oNode_resbelaeg.selectSingleNode("txt_153").Text
    resbelaeg_txt_154 = oNode_resbelaeg.selectSingleNode("txt_154").Text
    resbelaeg_txt_155 = oNode_resbelaeg.selectSingleNode("txt_155").Text
    resbelaeg_txt_156 = oNode_resbelaeg.selectSingleNode("txt_156").Text
    resbelaeg_txt_157 = oNode_resbelaeg.selectSingleNode("txt_157").Text
    resbelaeg_txt_158 = oNode_resbelaeg.selectSingleNode("txt_158").Text
    resbelaeg_txt_159 = oNode_resbelaeg.selectSingleNode("txt_159").Text
    resbelaeg_txt_160 = oNode_resbelaeg.selectSingleNode("txt_160").Text

    resbelaeg_txt_161 = oNode_resbelaeg.selectSingleNode("txt_161").Text
    resbelaeg_txt_162 = oNode_resbelaeg.selectSingleNode("txt_162").Text
    resbelaeg_txt_163 = oNode_resbelaeg.selectSingleNode("txt_163").Text
    resbelaeg_txt_164 = oNode_resbelaeg.selectSingleNode("txt_164").Text
    resbelaeg_txt_165 = oNode_resbelaeg.selectSingleNode("txt_165").Text
    resbelaeg_txt_166 = oNode_resbelaeg.selectSingleNode("txt_166").Text 
    resbelaeg_txt_167 = oNode_resbelaeg.selectSingleNode("txt_167").Text 
    resbelaeg_txt_168 = oNode_resbelaeg.selectSingleNode("txt_168").Text 
    resbelaeg_txt_169 = oNode_resbelaeg.selectSingleNode("txt_169").Text 
    resbelaeg_txt_170 = oNode_resbelaeg.selectSingleNode("txt_170").Text 

    resbelaeg_txt_171 = oNode_resbelaeg.selectSingleNode("txt_171").Text
    resbelaeg_txt_172 = oNode_resbelaeg.selectSingleNode("txt_172").Text
    resbelaeg_txt_173 = oNode_resbelaeg.selectSingleNode("txt_173").Text
    resbelaeg_txt_174 = oNode_resbelaeg.selectSingleNode("txt_174").Text
    resbelaeg_txt_175 = oNode_resbelaeg.selectSingleNode("txt_175").Text
    resbelaeg_txt_176 = oNode_resbelaeg.selectSingleNode("txt_176").Text 
    resbelaeg_txt_177 = oNode_resbelaeg.selectSingleNode("txt_177").Text 
    resbelaeg_txt_178 = oNode_resbelaeg.selectSingleNode("txt_178").Text 
    resbelaeg_txt_179 = oNode_resbelaeg.selectSingleNode("txt_179").Text 
    resbelaeg_txt_180 = oNode_resbelaeg.selectSingleNode("txt_180").Text
    
    resbelaeg_txt_181 = oNode_resbelaeg.selectSingleNode("txt_181").Text
    resbelaeg_txt_182 = oNode_resbelaeg.selectSingleNode("txt_182").Text
    resbelaeg_txt_183 = oNode_resbelaeg.selectSingleNode("txt_183").Text
    resbelaeg_txt_184 = oNode_resbelaeg.selectSingleNode("txt_184").Text
    resbelaeg_txt_185 = oNode_resbelaeg.selectSingleNode("txt_185").Text
    resbelaeg_txt_186 = oNode_resbelaeg.selectSingleNode("txt_186").Text 
    resbelaeg_txt_187 = oNode_resbelaeg.selectSingleNode("txt_187").Text 
    resbelaeg_txt_188 = oNode_resbelaeg.selectSingleNode("txt_188").Text 
    resbelaeg_txt_189 = oNode_resbelaeg.selectSingleNode("txt_189").Text 
    resbelaeg_txt_190 = oNode_resbelaeg.selectSingleNode("txt_190").Text
    
    resbelaeg_txt_191 = oNode_resbelaeg.selectSingleNode("txt_191").Text 
    resbelaeg_txt_192 = oNode_resbelaeg.selectSingleNode("txt_192").Text 
    resbelaeg_txt_193 = oNode_resbelaeg.selectSingleNode("txt_193").Text 
    resbelaeg_txt_194 = oNode_resbelaeg.selectSingleNode("txt_194").Text
    resbelaeg_txt_195 = oNode_resbelaeg.selectSingleNode("txt_195").Text
    resbelaeg_txt_196 = oNode_resbelaeg.selectSingleNode("txt_196").Text
    resbelaeg_txt_197 = oNode_resbelaeg.selectSingleNode("txt_197").Text
    resbelaeg_txt_198 = oNode_resbelaeg.selectSingleNode("txt_198").Text
    resbelaeg_txt_199 = oNode_resbelaeg.selectSingleNode("txt_199").Text
    resbelaeg_txt_200 = oNode_resbelaeg.selectSingleNode("txt_200").Text
    resbelaeg_txt_201 = oNode_resbelaeg.selectSingleNode("txt_201").Text
    resbelaeg_txt_202 = oNode_resbelaeg.selectSingleNode("txt_202").Text





  
          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>