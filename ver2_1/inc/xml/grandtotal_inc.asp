
<% 
Dim objXMLHTTP_grand, objXMLDOM_grand, i_grand, strHTML_grand

Set objXMLDom_grand = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_grand = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_grand.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/grandtotal_sprog.xml", False
'objXmlHttp_grand.open "GET", "http://localhost/inc/xml/grandtotal_sprog.xml", False
'objXmlHttp_grand.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/grand_sprog.xml", False
objXmlHttp_grand.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/grandtotal_sprog.xml", False
'objXmlHttp_grand.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/grandtotal_sprog.xml", False
'objXmlHttp_grand.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/grandtotal_sprog.xml", False

objXmlHttp_grand.send


Set objXmlDom_grand = objXmlHttp_grand.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_grand = Nothing



Dim Address_grand, Latitude_grand, Longitude_grand
Dim oNode_grand, oNodes_grand
Dim sXPathQuery_grand

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
sXPathQuery_grand = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_grand = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_grand = "//sprog/se"
Session.LCID = 1053
case 4
sXPathQuery_grand = "//sprog/no"
Session.LCID = 2068
case 5
sXPathQuery_grand = "//sprog/es"
Session.LCID = 1034
case 6
sXPathQuery_grand = "//sprog/de"
Session.LCID = 1031
case 7
sXPathQuery_grand = "//sprog/fr"
Session.LCID = 1036
case else
sXPathQuery_grand = "//sprog/dk"
Session.LCID = 1030
end select





Set oNode_grand = objXmlDom_grand.documentElement.selectSingleNode(sXPathQuery_grand)
Address_grand = oNode_grand.Text

Set oNodes_grand = objXmlDom_grand.documentElement.selectNodes(sXPathQuery_grand)

    For Each oNode_grand in oNodes_grand

    grand_txt_001 = oNode_grand.selectSingleNode("txt_1").Text
    grand_txt_002 = oNode_grand.selectSingleNode("txt_2").Text
    grand_txt_003 = oNode_grand.selectSingleNode("txt_3").Text
    grand_txt_003 = oNode_grand.selectSingleNode("txt_3").Text
    grand_txt_004 = oNode_grand.selectSingleNode("txt_4").Text
    grand_txt_005 = oNode_grand.selectSingleNode("txt_5").Text

   grand_txt_006 = oNode_grand.selectSingleNode("txt_6").Text
   grand_txt_007 = oNode_grand.selectSingleNode("txt_7").Text
   grand_txt_008 = oNode_grand.selectSingleNode("txt_8").Text
   grand_txt_009 = oNode_grand.selectSingleNode("txt_9").Text
   grand_txt_010 = oNode_grand.selectSingleNode("txt_10").Text
   grand_txt_011 = oNode_grand.selectSingleNode("txt_11").Text
    
   grand_txt_012 = oNode_grand.selectSingleNode("txt_12").Text
   grand_txt_013 = oNode_grand.selectSingleNode("txt_13").Text
   grand_txt_014 = oNode_grand.selectSingleNode("txt_14").Text
   grand_txt_015 = oNode_grand.selectSingleNode("txt_15").Text
   grand_txt_016 = oNode_grand.selectSingleNode("txt_16").Text
   grand_txt_017 = oNode_grand.selectSingleNode("txt_17").Text
    
   grand_txt_018 = oNode_grand.selectSingleNode("txt_18").Text
   grand_txt_019 = oNode_grand.selectSingleNode("txt_19").Text
   grand_txt_020 = oNode_grand.selectSingleNode("txt_20").Text
   grand_txt_021 = oNode_grand.selectSingleNode("txt_21").Text
   grand_txt_022 = oNode_grand.selectSingleNode("txt_22").Text
   grand_txt_023 = oNode_grand.selectSingleNode("txt_23").Text
   grand_txt_024 = oNode_grand.selectSingleNode("txt_24").Text
    
   grand_txt_025 = oNode_grand.selectSingleNode("txt_25").Text
   grand_txt_026 = oNode_grand.selectSingleNode("txt_26").Text
   grand_txt_027 = oNode_grand.selectSingleNode("txt_27").Text
   grand_txt_028 = oNode_grand.selectSingleNode("txt_28").Text
   grand_txt_029 = oNode_grand.selectSingleNode("txt_29").Text
   grand_txt_030 = oNode_grand.selectSingleNode("txt_30").Text
   grand_txt_031 = oNode_grand.selectSingleNode("txt_31").Text
   grand_txt_032 = oNode_grand.selectSingleNode("txt_32").Text
   grand_txt_033 = oNode_grand.selectSingleNode("txt_33").Text
   grand_txt_034 = oNode_grand.selectSingleNode("txt_34").Text
    
    grand_txt_035 = oNode_grand.selectSingleNode("txt_35").Text
     grand_txt_036 = oNode_grand.selectSingleNode("txt_36").Text
      grand_txt_037 = oNode_grand.selectSingleNode("txt_37").Text
       
       grand_txt_038 = oNode_grand.selectSingleNode("txt_38").Text
        grand_txt_039 = oNode_grand.selectSingleNode("txt_39").Text
        grand_txt_040 = oNode_grand.selectSingleNode("txt_40").Text
        grand_txt_041 = oNode_grand.selectSingleNode("txt_41").Text
        grand_txt_042 = oNode_grand.selectSingleNode("txt_42").Text
         grand_txt_043 = oNode_grand.selectSingleNode("txt_43").Text
         grand_txt_044 = oNode_grand.selectSingleNode("txt_44").Text
          
         grand_txt_045 = oNode_grand.selectSingleNode("txt_45").Text
         grand_txt_046 = oNode_grand.selectSingleNode("txt_46").Text
         grand_txt_047 = oNode_grand.selectSingleNode("txt_47").Text
         grand_txt_048 = oNode_grand.selectSingleNode("txt_48").Text
         grand_txt_049 = oNode_grand.selectSingleNode("txt_49").Text
         grand_txt_050 = oNode_grand.selectSingleNode("txt_50").Text
         grand_txt_051 = oNode_grand.selectSingleNode("txt_51").Text
         grand_txt_052 = oNode_grand.selectSingleNode("txt_52").Text
         grand_txt_053 = oNode_grand.selectSingleNode("txt_53").Text
         grand_txt_054 = oNode_grand.selectSingleNode("txt_54").Text
         grand_txt_055 = oNode_grand.selectSingleNode("txt_55").Text
         grand_txt_056 = oNode_grand.selectSingleNode("txt_56").Text
         grand_txt_057 = oNode_grand.selectSingleNode("txt_57").Text
        
        grand_txt_058 = oNode_grand.selectSingleNode("txt_58").Text
        grand_txt_059 = oNode_grand.selectSingleNode("txt_59").Text
    

   grand_txt_060 = oNode_grand.selectSingleNode("txt_60").Text
   grand_txt_061 = oNode_grand.selectSingleNode("txt_61").Text
   grand_txt_062 = oNode_grand.selectSingleNode("txt_62").Text
   grand_txt_063 = oNode_grand.selectSingleNode("txt_63").Text
   grand_txt_064 = oNode_grand.selectSingleNode("txt_64").Text
   grand_txt_065 = oNode_grand.selectSingleNode("txt_65").Text
   grand_txt_066 = oNode_grand.selectSingleNode("txt_66").Text
   grand_txt_067 = oNode_grand.selectSingleNode("txt_67").Text
   grand_txt_068 = oNode_grand.selectSingleNode("txt_68").Text
   grand_txt_069 = oNode_grand.selectSingleNode("txt_69").Text

   grand_txt_070 = oNode_grand.selectSingleNode("txt_70").Text
   grand_txt_071 = oNode_grand.selectSingleNode("txt_71").Text
   grand_txt_072 = oNode_grand.selectSingleNode("txt_72").Text
   grand_txt_070 = oNode_grand.selectSingleNode("txt_73").Text
   grand_txt_074 = oNode_grand.selectSingleNode("txt_74").Text
   grand_txt_075 = oNode_grand.selectSingleNode("txt_75").Text
   grand_txt_076 = oNode_grand.selectSingleNode("txt_76").Text
   grand_txt_077 = oNode_grand.selectSingleNode("txt_77").Text
   grand_txt_078 = oNode_grand.selectSingleNode("txt_78").Text
   grand_txt_079 = oNode_grand.selectSingleNode("txt_79").Text
    
   grand_txt_080 = oNode_grand.selectSingleNode("txt_80").Text
   grand_txt_081 = oNode_grand.selectSingleNode("txt_81").Text
   grand_txt_082 = oNode_grand.selectSingleNode("txt_82").Text
   grand_txt_083 = oNode_grand.selectSingleNode("txt_83").Text
   grand_txt_084 = oNode_grand.selectSingleNode("txt_84").Text
   grand_txt_085 = oNode_grand.selectSingleNode("txt_85").Text
   grand_txt_086 = oNode_grand.selectSingleNode("txt_86").Text
   grand_txt_087 = oNode_grand.selectSingleNode("txt_87").Text
   grand_txt_088 = oNode_grand.selectSingleNode("txt_88").Text
   grand_txt_089 = oNode_grand.selectSingleNode("txt_89").Text

   grand_txt_090 = oNode_grand.selectSingleNode("txt_90").Text
   grand_txt_091 = oNode_grand.selectSingleNode("txt_91").Text
   grand_txt_092 = oNode_grand.selectSingleNode("txt_92").Text
   grand_txt_093 = oNode_grand.selectSingleNode("txt_93").Text
   grand_txt_094 = oNode_grand.selectSingleNode("txt_94").Text
   grand_txt_095 = oNode_grand.selectSingleNode("txt_95").Text
   grand_txt_096 = oNode_grand.selectSingleNode("txt_96").Text
   grand_txt_097 = oNode_grand.selectSingleNode("txt_97").Text
   grand_txt_098 = oNode_grand.selectSingleNode("txt_98").Text
   grand_txt_099 = oNode_grand.selectSingleNode("txt_99").Text

   grand_txt_100 = oNode_grand.selectSingleNode("txt_100").Text
   grand_txt_101 = oNode_grand.selectSingleNode("txt_101").Text
   grand_txt_102 = oNode_grand.selectSingleNode("txt_102").Text
   grand_txt_103 = oNode_grand.selectSingleNode("txt_103").Text
   grand_txt_104 = oNode_grand.selectSingleNode("txt_104").Text
   grand_txt_105 = oNode_grand.selectSingleNode("txt_105").Text
   grand_txt_106 = oNode_grand.selectSingleNode("txt_106").Text
   grand_txt_107 = oNode_grand.selectSingleNode("txt_107").Text
   grand_txt_108 = oNode_grand.selectSingleNode("txt_108").Text
   grand_txt_109 = oNode_grand.selectSingleNode("txt_109").Text

   grand_txt_110 = oNode_grand.selectSingleNode("txt_110").Text
   grand_txt_111 = oNode_grand.selectSingleNode("txt_111").Text
   grand_txt_112 = oNode_grand.selectSingleNode("txt_112").Text
   grand_txt_113 = oNode_grand.selectSingleNode("txt_113").Text
   grand_txt_114 = oNode_grand.selectSingleNode("txt_114").Text
   grand_txt_115 = oNode_grand.selectSingleNode("txt_115").Text
   grand_txt_116 = oNode_grand.selectSingleNode("txt_116").Text
   grand_txt_117 = oNode_grand.selectSingleNode("txt_117").Text
   grand_txt_118 = oNode_grand.selectSingleNode("txt_118").Text
   grand_txt_119 = oNode_grand.selectSingleNode("txt_119").Text
    
   grand_txt_120 = oNode_grand.selectSingleNode("txt_120").Text
   grand_txt_121 = oNode_grand.selectSingleNode("txt_121").Text
   grand_txt_122 = oNode_grand.selectSingleNode("txt_122").Text
   grand_txt_123 = oNode_grand.selectSingleNode("txt_123").Text
   grand_txt_124 = oNode_grand.selectSingleNode("txt_124").Text
   grand_txt_125 = oNode_grand.selectSingleNode("txt_125").Text
   grand_txt_126 = oNode_grand.selectSingleNode("txt_126").Text
   grand_txt_127 = oNode_grand.selectSingleNode("txt_127").Text
   grand_txt_128 = oNode_grand.selectSingleNode("txt_128").Text
   grand_txt_129 = oNode_grand.selectSingleNode("txt_129").Text

   grand_txt_130 = oNode_grand.selectSingleNode("txt_130").Text
   grand_txt_131 = oNode_grand.selectSingleNode("txt_131").Text
   grand_txt_132 = oNode_grand.selectSingleNode("txt_132").Text
   grand_txt_133 = oNode_grand.selectSingleNode("txt_133").Text
   grand_txt_134 = oNode_grand.selectSingleNode("txt_134").Text
   grand_txt_135 = oNode_grand.selectSingleNode("txt_135").Text
   grand_txt_136 = oNode_grand.selectSingleNode("txt_136").Text
   grand_txt_137 = oNode_grand.selectSingleNode("txt_137").Text
   grand_txt_138 = oNode_grand.selectSingleNode("txt_138").Text
   grand_txt_139 = oNode_grand.selectSingleNode("txt_139").Text
   grand_txt_140 = oNode_grand.selectSingleNode("txt_140").Text

    grand_txt_141 = oNode_grand.selectSingleNode("txt_141").Text
            
    grand_txt_142 = oNode_grand.selectSingleNode("txt_142").Text
    grand_txt_143 = oNode_grand.selectSingleNode("txt_143").Text
    grand_txt_144 = oNode_grand.selectSingleNode("txt_144").Text
    grand_txt_145 = oNode_grand.selectSingleNode("txt_145").Text
    grand_txt_146 = oNode_grand.selectSingleNode("txt_146").Text
            
    grand_txt_147 = oNode_grand.selectSingleNode("txt_147").Text
    grand_txt_148 = oNode_grand.selectSingleNode("txt_148").Text
    grand_txt_149 = oNode_grand.selectSingleNode("txt_149").Text
            
    grand_txt_150 = oNode_grand.selectSingleNode("txt_150").Text
    grand_txt_151 = oNode_grand.selectSingleNode("txt_151").Text
    grand_txt_152 = oNode_grand.selectSingleNode("txt_152").Text
    grand_txt_153 = oNode_grand.selectSingleNode("txt_153").Text
    grand_txt_154 = oNode_grand.selectSingleNode("txt_154").Text
    grand_txt_155 = oNode_grand.selectSingleNode("txt_155").Text
    grand_txt_156 = oNode_grand.selectSingleNode("txt_156").Text
    grand_txt_157 = oNode_grand.selectSingleNode("txt_157").Text
    grand_txt_158 = oNode_grand.selectSingleNode("txt_158").Text
    grand_txt_159 = oNode_grand.selectSingleNode("txt_159").Text
            
    grand_txt_160 = oNode_grand.selectSingleNode("txt_160").Text
    grand_txt_161 = oNode_grand.selectSingleNode("txt_161").Text
    grand_txt_162 = oNode_grand.selectSingleNode("txt_162").Text
    grand_txt_163 = oNode_grand.selectSingleNode("txt_163").Text
    grand_txt_164 = oNode_grand.selectSingleNode("txt_164").Text
    grand_txt_165 = oNode_grand.selectSingleNode("txt_165").Text
    grand_txt_166 = oNode_grand.selectSingleNode("txt_166").Text
    grand_txt_167 = oNode_grand.selectSingleNode("txt_167").Text
    grand_txt_168 = oNode_grand.selectSingleNode("txt_168").Text
    grand_txt_169 = oNode_grand.selectSingleNode("txt_169").Text
            
    grand_txt_170 = oNode_grand.selectSingleNode("txt_170").Text
    grand_txt_171 = oNode_grand.selectSingleNode("txt_171").Text
    grand_txt_172 = oNode_grand.selectSingleNode("txt_172").Text
    grand_txt_173 = oNode_grand.selectSingleNode("txt_173").Text
    grand_txt_174 = oNode_grand.selectSingleNode("txt_174").Text
    grand_txt_175 = oNode_grand.selectSingleNode("txt_175").Text
    grand_txt_176 = oNode_grand.selectSingleNode("txt_176").Text
    grand_txt_177 = oNode_grand.selectSingleNode("txt_177").Text
    grand_txt_178 = oNode_grand.selectSingleNode("txt_178").Text
    grand_txt_179 = oNode_grand.selectSingleNode("txt_179").Text
            
    grand_txt_180 = oNode_grand.selectSingleNode("txt_180").Text
    grand_txt_181 = oNode_grand.selectSingleNode("txt_181").Text
    grand_txt_182 = oNode_grand.selectSingleNode("txt_182").Text
    grand_txt_183 = oNode_grand.selectSingleNode("txt_183").Text
    grand_txt_184 = oNode_grand.selectSingleNode("txt_184").Text
    grand_txt_185 = oNode_grand.selectSingleNode("txt_185").Text
    grand_txt_186 = oNode_grand.selectSingleNode("txt_186").Text
    grand_txt_187 = oNode_grand.selectSingleNode("txt_187").Text
            
    grand_txt_188 = oNode_grand.selectSingleNode("txt_188").Text
    grand_txt_189 = oNode_grand.selectSingleNode("txt_189").Text
    grand_txt_190 = oNode_grand.selectSingleNode("txt_190").Text
    grand_txt_191 = oNode_grand.selectSingleNode("txt_191").Text
    grand_txt_192 = oNode_grand.selectSingleNode("txt_192").Text

  
          
    next




'Response.Write "grand_txt_001: " & grand_txt_001 & "<br>"
'Response.Write "grand_txt_002: " & grand_txt_002 & "<br>"


%>