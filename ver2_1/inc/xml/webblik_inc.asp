
<% 
Dim objXMLHTTP_webblik, objXMLDOM_webblik, i_webblik, strHTML_webblik

Set objXMLDom_webblik = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_webblik = Server.CreateObject("Msxml2.ServerXMLHTTP")
objXmlHttp_webblik.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/webblik_sprog.xml", False
'objXmlHttp_webblik.open "GET", "http://localhost/inc/xml/webblik_sprog.xml", False
'objXmlHttp_webblik.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/webblik_sprog.xml", False
'objXmlHttp_webblik.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/webblik_sprog.xml", False
'objXmlHttp_webblik.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/webblik_sprog.xml", False
'objXmlHttp_webblik.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/webblik_sprog.xml", False

objXmlHttp_webblik.send


Set objXmlDom_webblik = objXmlHttp_webblik.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_webblik = Nothing



Dim Address_webblik, Latitude_webblik, Longitude_webblik
Dim oNode_webblik, oNodes_webblik
Dim sXPathQuery_webblik

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
sXPathQuery_webblik = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_webblik = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_webblik = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_webblik = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_webblik = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_webblik = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_webblik = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_webblik = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
'Session.LCID = 1030



Set oNode_webblik = objXmlDom_webblik.documentElement.selectSingleNode(sXPathQuery_webblik)
Address_webblik = oNode_webblik.Text

Set oNodes_webblik = objXmlDom_webblik.documentElement.selectNodes(sXPathQuery_webblik)

    For Each oNode_webblik in oNodes_webblik

          
    tsa_txt_001 = oNode_webblik.selectSingleNode("txt_1").Text
     tsa_txt_002 = oNode_webblik.selectSingleNode("txt_2").Text
     tsa_txt_003 = oNode_webblik.selectSingleNode("txt_3").Text
     tsa_txt_003 = oNode_webblik.selectSingleNode("txt_3").Text
    tsa_txt_004 = oNode_webblik.selectSingleNode("txt_4").Text
    tsa_txt_005 = oNode_webblik.selectSingleNode("txt_5").Text

    tsa_txt_006 = oNode_webblik.selectSingleNode("txt_6").Text
    tsa_txt_007 = oNode_webblik.selectSingleNode("txt_7").Text
    tsa_txt_008 = oNode_webblik.selectSingleNode("txt_8").Text
    tsa_txt_009 = oNode_webblik.selectSingleNode("txt_9").Text
    tsa_txt_010 = oNode_webblik.selectSingleNode("txt_10").Text
    tsa_txt_011 = oNode_webblik.selectSingleNode("txt_11").Text
    
    tsa_txt_012 = oNode_webblik.selectSingleNode("txt_12").Text
    tsa_txt_013 = oNode_webblik.selectSingleNode("txt_13").Text
    tsa_txt_014 = oNode_webblik.selectSingleNode("txt_14").Text
    tsa_txt_015 = oNode_webblik.selectSingleNode("txt_15").Text
    tsa_txt_016 = oNode_webblik.selectSingleNode("txt_16").Text
    tsa_txt_017 = oNode_webblik.selectSingleNode("txt_17").Text
    
    tsa_txt_018 = oNode_webblik.selectSingleNode("txt_18").Text
    tsa_txt_019 = oNode_webblik.selectSingleNode("txt_19").Text
    tsa_txt_020 = oNode_webblik.selectSingleNode("txt_20").Text
    tsa_txt_021 = oNode_webblik.selectSingleNode("txt_21").Text
    tsa_txt_022 = oNode_webblik.selectSingleNode("txt_22").Text
    tsa_txt_023 = oNode_webblik.selectSingleNode("txt_23").Text
    tsa_txt_024 = oNode_webblik.selectSingleNode("txt_24").Text
    
    tsa_txt_025 = oNode_webblik.selectSingleNode("txt_25").Text
    tsa_txt_026 = oNode_webblik.selectSingleNode("txt_26").Text
    tsa_txt_027 = oNode_webblik.selectSingleNode("txt_27").Text
    tsa_txt_028 = oNode_webblik.selectSingleNode("txt_28").Text
    tsa_txt_029 = oNode_webblik.selectSingleNode("txt_29").Text
    tsa_txt_030 = oNode_webblik.selectSingleNode("txt_30").Text
    tsa_txt_031 = oNode_webblik.selectSingleNode("txt_31").Text
    tsa_txt_032 = oNode_webblik.selectSingleNode("txt_32").Text
    tsa_txt_033 = oNode_webblik.selectSingleNode("txt_33").Text
    tsa_txt_034 = oNode_webblik.selectSingleNode("txt_34").Text
    
     tsa_txt_035 = oNode_webblik.selectSingleNode("txt_35").Text
      tsa_txt_036 = oNode_webblik.selectSingleNode("txt_36").Text
       tsa_txt_037 = oNode_webblik.selectSingleNode("txt_37").Text
       
        tsa_txt_038 = oNode_webblik.selectSingleNode("txt_38").Text
         tsa_txt_039 = oNode_webblik.selectSingleNode("txt_39").Text
         tsa_txt_040 = oNode_webblik.selectSingleNode("txt_40").Text
         tsa_txt_041 = oNode_webblik.selectSingleNode("txt_41").Text
         tsa_txt_042 = oNode_webblik.selectSingleNode("txt_42").Text
          tsa_txt_043 = oNode_webblik.selectSingleNode("txt_43").Text
          tsa_txt_044 = oNode_webblik.selectSingleNode("txt_44").Text
          
          tsa_txt_045 = oNode_webblik.selectSingleNode("txt_45").Text
          tsa_txt_046 = oNode_webblik.selectSingleNode("txt_46").Text
          tsa_txt_047 = oNode_webblik.selectSingleNode("txt_47").Text
          tsa_txt_048 = oNode_webblik.selectSingleNode("txt_48").Text
          tsa_txt_049 = oNode_webblik.selectSingleNode("txt_49").Text
          tsa_txt_050 = oNode_webblik.selectSingleNode("txt_50").Text
          
          tsa_txt_051 = oNode_webblik.selectSingleNode("txt_51").Text
          tsa_txt_052 = oNode_webblik.selectSingleNode("txt_52").Text
          tsa_txt_053 = oNode_webblik.selectSingleNode("txt_53").Text
          tsa_txt_054 = oNode_webblik.selectSingleNode("txt_54").Text
          tsa_txt_055 = oNode_webblik.selectSingleNode("txt_55").Text
          tsa_txt_056 = oNode_webblik.selectSingleNode("txt_56").Text
          tsa_txt_057 = oNode_webblik.selectSingleNode("txt_57").Text
          tsa_txt_058 = oNode_webblik.selectSingleNode("txt_58").Text
          tsa_txt_059 = oNode_webblik.selectSingleNode("txt_59").Text
          tsa_txt_060 = oNode_webblik.selectSingleNode("txt_60").Text
          tsa_txt_061 = oNode_webblik.selectSingleNode("txt_61").Text
          tsa_txt_062 = oNode_webblik.selectSingleNode("txt_62").Text
          tsa_txt_063 = oNode_webblik.selectSingleNode("txt_63").Text
          tsa_txt_064 = oNode_webblik.selectSingleNode("txt_64").Text
          tsa_txt_065 = oNode_webblik.selectSingleNode("txt_65").Text
          tsa_txt_066 = oNode_webblik.selectSingleNode("txt_66").Text
          
          tsa_txt_067 = oNode_webblik.selectSingleNode("txt_67").Text
          tsa_txt_068 = oNode_webblik.selectSingleNode("txt_68").Text
          tsa_txt_069 = oNode_webblik.selectSingleNode("txt_69").Text
          tsa_txt_070 = oNode_webblik.selectSingleNode("txt_70").Text
          tsa_txt_071 = oNode_webblik.selectSingleNode("txt_71").Text
          tsa_txt_072 = oNode_webblik.selectSingleNode("txt_72").Text
          tsa_txt_073 = oNode_webblik.selectSingleNode("txt_73").Text
          tsa_txt_074 = oNode_webblik.selectSingleNode("txt_74").Text
          tsa_txt_075 = oNode_webblik.selectSingleNode("txt_75").Text
          tsa_txt_076 = oNode_webblik.selectSingleNode("txt_76").Text
          tsa_txt_077 = oNode_webblik.selectSingleNode("txt_77").Text
          tsa_txt_078 = oNode_webblik.selectSingleNode("txt_78").Text
          tsa_txt_079 = oNode_webblik.selectSingleNode("txt_79").Text
        
          
          tsa_txt_080 = oNode_webblik.selectSingleNode("txt_80").Text
          tsa_txt_081 = oNode_webblik.selectSingleNode("txt_81").Text
          tsa_txt_082 = oNode_webblik.selectSingleNode("txt_82").Text
          tsa_txt_083 = oNode_webblik.selectSingleNode("txt_83").Text
          
          tsa_txt_084 = oNode_webblik.selectSingleNode("txt_84").Text
          tsa_txt_085 = oNode_webblik.selectSingleNode("txt_85").Text
          tsa_txt_086 = oNode_webblik.selectSingleNode("txt_86").Text
          
          tsa_txt_087 = oNode_webblik.selectSingleNode("txt_87").Text
          
          tsa_txt_088 = oNode_webblik.selectSingleNode("txt_88").Text
          tsa_txt_089 = oNode_webblik.selectSingleNode("txt_89").Text
          tsa_txt_090 = oNode_webblik.selectSingleNode("txt_90").Text
          
          tsa_txt_091 = oNode_webblik.selectSingleNode("txt_91").Text
          tsa_txt_092 = oNode_webblik.selectSingleNode("txt_92").Text
          tsa_txt_093 = oNode_webblik.selectSingleNode("txt_93").Text
          tsa_txt_094 = oNode_webblik.selectSingleNode("txt_94").Text
          tsa_txt_095 = oNode_webblik.selectSingleNode("txt_95").Text
          tsa_txt_096 = oNode_webblik.selectSingleNode("txt_96").Text
          tsa_txt_097 = oNode_webblik.selectSingleNode("txt_97").Text
          tsa_txt_098 = oNode_webblik.selectSingleNode("txt_98").Text
          tsa_txt_099 = oNode_webblik.selectSingleNode("txt_99").Text
          
          tsa_txt_100 = oNode_webblik.selectSingleNode("txt_100").Text
          tsa_txt_101 = oNode_webblik.selectSingleNode("txt_101").Text
          tsa_txt_102 = oNode_webblik.selectSingleNode("txt_102").Text
          tsa_txt_103 = oNode_webblik.selectSingleNode("txt_103").Text
          
          tsa_txt_104 = oNode_webblik.selectSingleNode("txt_104").Text
          tsa_txt_105 = oNode_webblik.selectSingleNode("txt_105").Text
          tsa_txt_106 = oNode_webblik.selectSingleNode("txt_106").Text
         
          
          tsa_txt_116 = oNode_webblik.selectSingleNode("txt_116").Text
          tsa_txt_117 = oNode_webblik.selectSingleNode("txt_117").Text
          tsa_txt_118 = oNode_webblik.selectSingleNode("txt_118").Text
          tsa_txt_119 = oNode_webblik.selectSingleNode("txt_119").Text
         
          tsa_txt_120 = oNode_webblik.selectSingleNode("txt_120").Text
          tsa_txt_121 = oNode_webblik.selectSingleNode("txt_121").Text
          tsa_txt_122 = oNode_webblik.selectSingleNode("txt_122").Text
          tsa_txt_123 = oNode_webblik.selectSingleNode("txt_123").Text
          
          tsa_txt_124 = oNode_webblik.selectSingleNode("txt_124").Text
           tsa_txt_125 = oNode_webblik.selectSingleNode("txt_125").Text
            tsa_txt_126 = oNode_webblik.selectSingleNode("txt_126").Text
            tsa_txt_127 = oNode_webblik.selectSingleNode("txt_127").Text
            tsa_txt_128 = oNode_webblik.selectSingleNode("txt_128").Text
            
            tsa_txt_129 = oNode_webblik.selectSingleNode("txt_129").Text
            tsa_txt_130 = oNode_webblik.selectSingleNode("txt_130").Text
            tsa_txt_131 = oNode_webblik.selectSingleNode("txt_131").Text
            tsa_txt_132 = oNode_webblik.selectSingleNode("txt_132").Text
            tsa_txt_133 = oNode_webblik.selectSingleNode("txt_133").Text
            
            tsa_txt_134 = oNode_webblik.selectSingleNode("txt_134").Text
            tsa_txt_135 = oNode_webblik.selectSingleNode("txt_135").Text
            tsa_txt_136 = oNode_webblik.selectSingleNode("txt_136").Text
            tsa_txt_137 = oNode_webblik.selectSingleNode("txt_137").Text
            tsa_txt_138 = oNode_webblik.selectSingleNode("txt_138").Text
            tsa_txt_139 = oNode_webblik.selectSingleNode("txt_139").Text
            
            tsa_txt_140 = oNode_webblik.selectSingleNode("txt_140").Text
            tsa_txt_141 = oNode_webblik.selectSingleNode("txt_141").Text
            
            tsa_txt_142 = oNode_webblik.selectSingleNode("txt_142").Text
            tsa_txt_143 = oNode_webblik.selectSingleNode("txt_143").Text
            tsa_txt_144 = oNode_webblik.selectSingleNode("txt_144").Text
            tsa_txt_145 = oNode_webblik.selectSingleNode("txt_145").Text
            tsa_txt_146 = oNode_webblik.selectSingleNode("txt_146").Text
            
            tsa_txt_147 = oNode_webblik.selectSingleNode("txt_147").Text
            tsa_txt_148 = oNode_webblik.selectSingleNode("txt_148").Text
            tsa_txt_149 = oNode_webblik.selectSingleNode("txt_149").Text
            
            tsa_txt_150 = oNode_webblik.selectSingleNode("txt_150").Text
            tsa_txt_151 = oNode_webblik.selectSingleNode("txt_151").Text
            tsa_txt_152 = oNode_webblik.selectSingleNode("txt_152").Text
            tsa_txt_153 = oNode_webblik.selectSingleNode("txt_153").Text
            tsa_txt_154 = oNode_webblik.selectSingleNode("txt_154").Text
            tsa_txt_155 = oNode_webblik.selectSingleNode("txt_155").Text
            tsa_txt_156 = oNode_webblik.selectSingleNode("txt_156").Text
            tsa_txt_157 = oNode_webblik.selectSingleNode("txt_157").Text
            tsa_txt_158 = oNode_webblik.selectSingleNode("txt_158").Text
            tsa_txt_159 = oNode_webblik.selectSingleNode("txt_159").Text
            
            tsa_txt_160 = oNode_webblik.selectSingleNode("txt_160").Text
            tsa_txt_161 = oNode_webblik.selectSingleNode("txt_161").Text
            tsa_txt_162 = oNode_webblik.selectSingleNode("txt_162").Text
            tsa_txt_163 = oNode_webblik.selectSingleNode("txt_163").Text
            tsa_txt_164 = oNode_webblik.selectSingleNode("txt_164").Text
            tsa_txt_165 = oNode_webblik.selectSingleNode("txt_165").Text
            tsa_txt_166 = oNode_webblik.selectSingleNode("txt_166").Text
            tsa_txt_167 = oNode_webblik.selectSingleNode("txt_167").Text
            tsa_txt_168 = oNode_webblik.selectSingleNode("txt_168").Text
            tsa_txt_169 = oNode_webblik.selectSingleNode("txt_169").Text
            
            tsa_txt_170 = oNode_webblik.selectSingleNode("txt_170").Text
            tsa_txt_171 = oNode_webblik.selectSingleNode("txt_171").Text
            tsa_txt_172 = oNode_webblik.selectSingleNode("txt_172").Text
            tsa_txt_173 = oNode_webblik.selectSingleNode("txt_173").Text
            tsa_txt_174 = oNode_webblik.selectSingleNode("txt_174").Text
            tsa_txt_175 = oNode_webblik.selectSingleNode("txt_175").Text
            tsa_txt_176 = oNode_webblik.selectSingleNode("txt_176").Text
            tsa_txt_177 = oNode_webblik.selectSingleNode("txt_177").Text
            tsa_txt_178 = oNode_webblik.selectSingleNode("txt_178").Text
            tsa_txt_179 = oNode_webblik.selectSingleNode("txt_179").Text
            
            tsa_txt_180 = oNode_webblik.selectSingleNode("txt_180").Text
            tsa_txt_181 = oNode_webblik.selectSingleNode("txt_181").Text
            tsa_txt_182 = oNode_webblik.selectSingleNode("txt_182").Text
            tsa_txt_183 = oNode_webblik.selectSingleNode("txt_183").Text
            tsa_txt_184 = oNode_webblik.selectSingleNode("txt_184").Text
            tsa_txt_185 = oNode_webblik.selectSingleNode("txt_185").Text
            tsa_txt_186 = oNode_webblik.selectSingleNode("txt_186").Text
            tsa_txt_187 = oNode_webblik.selectSingleNode("txt_187").Text
            
            tsa_txt_188 = oNode_webblik.selectSingleNode("txt_188").Text
            tsa_txt_189 = oNode_webblik.selectSingleNode("txt_189").Text
            
            tsa_txt_190 = oNode_webblik.selectSingleNode("txt_190").Text
            tsa_txt_191 = oNode_webblik.selectSingleNode("txt_191").Text
            tsa_txt_192 = oNode_webblik.selectSingleNode("txt_192").Text
            tsa_txt_193 = oNode_webblik.selectSingleNode("txt_193").Text
            tsa_txt_194 = oNode_webblik.selectSingleNode("txt_194").Text
            tsa_txt_195 = oNode_webblik.selectSingleNode("txt_195").Text
            tsa_txt_196 = oNode_webblik.selectSingleNode("txt_196").Text
            tsa_txt_197 = oNode_webblik.selectSingleNode("txt_197").Text
            tsa_txt_198 = oNode_webblik.selectSingleNode("txt_198").Text
          
    next




'Response.Write "week_txt_001: " & week_txt_001 & "<br>"
'Response.Write "week_txt_002: " & week_txt_002 & "<br>"


%>