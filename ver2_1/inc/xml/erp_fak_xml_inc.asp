
<% 
'if thisfile <> "erp_oprfak_fs" then 'fejler i _fs da de åbenbart bliver declareret 2 gange i forbindelse med faktura_godkendt visning **'
Dim objXMLHTTP_erp, objXMLDOM_erp, z, strHTML_erp
Dim Address_erp, Latitude_erp, Longitude_erp
Dim oNode_erp, oNodes_erp
Dim sXPathQuery_erp
'end if

Set objXMLDOM_erp = Server.CreateObject("Microsoft.XMLDOM")
Set objXMLHTTP_erp = Server.CreateObject("Msxml2.ServerXMLHTTP")
objXMLHTTP_erp.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/erp_fak_sprog.xml", False
'objXMLHTTP_erp.open "GET", "http://localhost/inc/xml/erp_fak_sprog.xml", False
'objXMLHTTP_erp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_1/inc/xml/erp_fak_sprog.xml", False
'objXMLHTTP_erp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/erp_fak_sprog.xml", False
'objXMLHTTP_erp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/erp_fak_sprog.xml", False
'objXMLHTTP_erp.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/erp_fak_sprog.xml", False    
'objXMLHTTP_erp.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver4_22/inc/xml/erp_fak_sprog.xml", False
objXMLHTTP_erp.send


Set objXMLDOM_erp = objXMLHTTP_erp.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM_erp.getElementsByTagName("se")


Set objXMLHTTP_erp = Nothing


'strTxt_1 = objModuler(0).text   'Timeregistrering 
    
        if len(trim(request("visfaktura"))) <> 0 then
        visfakturaSprog = request("visfaktura")
        else
        visfakturaSprog = 0
        end if

        'Response.Write "HER: " & thisfile & " visfaktura: "& visfakturaSprog

        'response.write "thisfile" & thisfile
        
        '**** Faktura oprettelse følger sprog på medarbejder ***' 
        sprog = 1
        if len(trim(session("mid"))) <> 0 then
        
        strSQL = "SELECT sprog FROM medarbejdere WHERE mid = " & session("mid")
        oRec.open strSQL, oConn, 3
        if not oRec.EOF then
        sprog = oRec("sprog")
        end if
        oRec.close
        end if 	
        
        
        '*** HVIS DER vises faktura layout overruler valgt sprog på faktura 
        if len(trim(request("id"))) <> 0 AND instr(request("id"), ",") = 0 then
        thisFakid = request("id")
        else
        thisFakid = 0
        end if

        
        '** Sprog valgt på faktura overruler sprog settings på medarbejder KUN FAKTURA LAYOUT **'
        if thisFakid <> "0" AND len(trim(thisFakid)) <> 0 AND cint(visfakturaSprog) = 2 then

	        strSQLFak = "SELECT f.sprog FROM fakturaer AS f WHERE fid = "& thisFakid

            'response.write strSQLFak
            'response.flush

            oRec6.open strSQLFak, oConn, 3
            if not oRec6.EOF then

                sprog = oRec6("sprog")

            end if
            oRec6.close
        
        end if

'sprog = 2
select case sprog
case 1
sXPathQuery_erp = "//sprog/dk"
Session.LCID = 1030
case 2
sXPathQuery_erp = "//sprog/uk"
Session.LCID = 2057
case 3
sXPathQuery_erp = "//sprog/se"
Session.LCID = 1053
case 4
sXPathQuery_erp = "//sprog/no"
Session.LCID = 2068
case 5
sXPathQuery_erp = "//sprog/es"
Session.LCID = 1034
case 6
sXPathQuery_erp = "//sprog/de"
Session.LCID = 1031
case 7
sXPathQuery_erp = "//sprog/fr"
Session.LCID = 1036
case else
sXPathQuery_erp = "//sprog/dk"
Session.LCID = 1030
end select



Set oNode_erp = objXMLDOM_erp.documentElement.selectSingleNode(sXPathQuery_erp)
Address_erp = oNode_erp.Text

Set oNodes_erp = objXMLDOM_erp.documentElement.selectNodes(sXPathQuery_erp)

For Each oNode_erp in oNodes_erp

    erp_txt_001 = oNode_erp.selectSingleNode("txt_1").Text
    erp_txt_002 = oNode_erp.selectSingleNode("txt_2").Text
    erp_txt_003 = oNode_erp.selectSingleNode("txt_3").Text
    erp_txt_003 = oNode_erp.selectSingleNode("txt_3").Text
   erp_txt_004 = oNode_erp.selectSingleNode("txt_4").Text
   erp_txt_005 = oNode_erp.selectSingleNode("txt_5").Text

   erp_txt_006 = oNode_erp.selectSingleNode("txt_6").Text
   erp_txt_007 = oNode_erp.selectSingleNode("txt_7").Text
   erp_txt_008 = oNode_erp.selectSingleNode("txt_8").Text
   erp_txt_009 = oNode_erp.selectSingleNode("txt_9").Text
   erp_txt_010 = oNode_erp.selectSingleNode("txt_10").Text
   erp_txt_011 = oNode_erp.selectSingleNode("txt_11").Text
    
   erp_txt_012 = oNode_erp.selectSingleNode("txt_12").Text
   erp_txt_013 = oNode_erp.selectSingleNode("txt_13").Text
   erp_txt_014 = oNode_erp.selectSingleNode("txt_14").Text
   erp_txt_015 = oNode_erp.selectSingleNode("txt_15").Text
   erp_txt_016 = oNode_erp.selectSingleNode("txt_16").Text
   erp_txt_017 = oNode_erp.selectSingleNode("txt_17").Text
    
   erp_txt_018 = oNode_erp.selectSingleNode("txt_18").Text
   erp_txt_019 = oNode_erp.selectSingleNode("txt_19").Text
   erp_txt_020 = oNode_erp.selectSingleNode("txt_20").Text
   erp_txt_021 = oNode_erp.selectSingleNode("txt_21").Text
   erp_txt_022 = oNode_erp.selectSingleNode("txt_22").Text
   erp_txt_023 = oNode_erp.selectSingleNode("txt_23").Text
   erp_txt_024 = oNode_erp.selectSingleNode("txt_24").Text
    
   erp_txt_025 = oNode_erp.selectSingleNode("txt_25").Text
   erp_txt_026 = oNode_erp.selectSingleNode("txt_26").Text
   erp_txt_027 = oNode_erp.selectSingleNode("txt_27").Text
   erp_txt_028 = oNode_erp.selectSingleNode("txt_28").Text
   erp_txt_029 = oNode_erp.selectSingleNode("txt_29").Text
   erp_txt_030 = oNode_erp.selectSingleNode("txt_30").Text
   erp_txt_031 = oNode_erp.selectSingleNode("txt_31").Text
   erp_txt_032 = oNode_erp.selectSingleNode("txt_32").Text
   erp_txt_033 = oNode_erp.selectSingleNode("txt_33").Text
   erp_txt_034 = oNode_erp.selectSingleNode("txt_34").Text
    
    erp_txt_035 = oNode_erp.selectSingleNode("txt_35").Text
     erp_txt_036 = oNode_erp.selectSingleNode("txt_36").Text
      erp_txt_037 = oNode_erp.selectSingleNode("txt_37").Text
       
       erp_txt_038 = oNode_erp.selectSingleNode("txt_38").Text
        erp_txt_039 = oNode_erp.selectSingleNode("txt_39").Text
        erp_txt_040 = oNode_erp.selectSingleNode("txt_40").Text
        erp_txt_041 = oNode_erp.selectSingleNode("txt_41").Text
        erp_txt_042 = oNode_erp.selectSingleNode("txt_42").Text
         erp_txt_043 = oNode_erp.selectSingleNode("txt_43").Text
         erp_txt_044 = oNode_erp.selectSingleNode("txt_44").Text
          
         erp_txt_045 = oNode_erp.selectSingleNode("txt_45").Text
         erp_txt_046 = oNode_erp.selectSingleNode("txt_46").Text
         erp_txt_047 = oNode_erp.selectSingleNode("txt_47").Text
         erp_txt_048 = oNode_erp.selectSingleNode("txt_48").Text
         erp_txt_049 = oNode_erp.selectSingleNode("txt_49").Text
         erp_txt_050 = oNode_erp.selectSingleNode("txt_50").Text
         erp_txt_051 = oNode_erp.selectSingleNode("txt_51").Text
         erp_txt_052 = oNode_erp.selectSingleNode("txt_52").Text
         erp_txt_053 = oNode_erp.selectSingleNode("txt_53").Text
         erp_txt_054 = oNode_erp.selectSingleNode("txt_54").Text
         erp_txt_055 = oNode_erp.selectSingleNode("txt_55").Text
         erp_txt_056 = oNode_erp.selectSingleNode("txt_56").Text
         erp_txt_057 = oNode_erp.selectSingleNode("txt_57").Text
        
        erp_txt_058 = oNode_erp.selectSingleNode("txt_58").Text
        erp_txt_059 = oNode_erp.selectSingleNode("txt_59").Text
    

   erp_txt_060 = oNode_erp.selectSingleNode("txt_60").Text
   erp_txt_061 = oNode_erp.selectSingleNode("txt_61").Text
   erp_txt_062 = oNode_erp.selectSingleNode("txt_62").Text
   erp_txt_063 = oNode_erp.selectSingleNode("txt_63").Text
   erp_txt_064 = oNode_erp.selectSingleNode("txt_64").Text
   erp_txt_065 = oNode_erp.selectSingleNode("txt_65").Text
   erp_txt_066 = oNode_erp.selectSingleNode("txt_66").Text
   erp_txt_067 = oNode_erp.selectSingleNode("txt_67").Text
   erp_txt_068 = oNode_erp.selectSingleNode("txt_68").Text
   erp_txt_069 = oNode_erp.selectSingleNode("txt_69").Text

   erp_txt_070 = oNode_erp.selectSingleNode("txt_70").Text
   erp_txt_071 = oNode_erp.selectSingleNode("txt_71").Text
   erp_txt_072 = oNode_erp.selectSingleNode("txt_72").Text
   erp_txt_070 = oNode_erp.selectSingleNode("txt_73").Text
   erp_txt_074 = oNode_erp.selectSingleNode("txt_74").Text
   erp_txt_075 = oNode_erp.selectSingleNode("txt_75").Text
   erp_txt_076 = oNode_erp.selectSingleNode("txt_76").Text
   erp_txt_077 = oNode_erp.selectSingleNode("txt_77").Text
   erp_txt_078 = oNode_erp.selectSingleNode("txt_78").Text
   erp_txt_079 = oNode_erp.selectSingleNode("txt_79").Text
    
   erp_txt_080 = oNode_erp.selectSingleNode("txt_80").Text
   erp_txt_081 = oNode_erp.selectSingleNode("txt_81").Text
   erp_txt_082 = oNode_erp.selectSingleNode("txt_82").Text
   erp_txt_083 = oNode_erp.selectSingleNode("txt_83").Text
   erp_txt_084 = oNode_erp.selectSingleNode("txt_84").Text
   erp_txt_085 = oNode_erp.selectSingleNode("txt_85").Text
   erp_txt_086 = oNode_erp.selectSingleNode("txt_86").Text
   erp_txt_087 = oNode_erp.selectSingleNode("txt_87").Text
   erp_txt_088 = oNode_erp.selectSingleNode("txt_88").Text
   erp_txt_089 = oNode_erp.selectSingleNode("txt_89").Text

   erp_txt_090 = oNode_erp.selectSingleNode("txt_90").Text
   erp_txt_091 = oNode_erp.selectSingleNode("txt_91").Text
   erp_txt_092 = oNode_erp.selectSingleNode("txt_92").Text
   erp_txt_093 = oNode_erp.selectSingleNode("txt_93").Text
   erp_txt_094 = oNode_erp.selectSingleNode("txt_94").Text
   erp_txt_095 = oNode_erp.selectSingleNode("txt_95").Text
   erp_txt_096 = oNode_erp.selectSingleNode("txt_96").Text
   erp_txt_097 = oNode_erp.selectSingleNode("txt_97").Text
   erp_txt_098 = oNode_erp.selectSingleNode("txt_98").Text
   erp_txt_099 = oNode_erp.selectSingleNode("txt_99").Text

   erp_txt_100 = oNode_erp.selectSingleNode("txt_100").Text
   erp_txt_101 = oNode_erp.selectSingleNode("txt_101").Text
   erp_txt_102 = oNode_erp.selectSingleNode("txt_102").Text
   erp_txt_103 = oNode_erp.selectSingleNode("txt_103").Text
   erp_txt_104 = oNode_erp.selectSingleNode("txt_104").Text
   erp_txt_105 = oNode_erp.selectSingleNode("txt_105").Text
   erp_txt_106 = oNode_erp.selectSingleNode("txt_106").Text
   erp_txt_107 = oNode_erp.selectSingleNode("txt_107").Text
   erp_txt_108 = oNode_erp.selectSingleNode("txt_108").Text
   erp_txt_109 = oNode_erp.selectSingleNode("txt_109").Text

   erp_txt_110 = oNode_erp.selectSingleNode("txt_110").Text
   erp_txt_111 = oNode_erp.selectSingleNode("txt_111").Text
   erp_txt_112 = oNode_erp.selectSingleNode("txt_112").Text
   erp_txt_113 = oNode_erp.selectSingleNode("txt_113").Text
   erp_txt_114 = oNode_erp.selectSingleNode("txt_114").Text
   erp_txt_115 = oNode_erp.selectSingleNode("txt_115").Text
   erp_txt_116 = oNode_erp.selectSingleNode("txt_116").Text
   erp_txt_117 = oNode_erp.selectSingleNode("txt_117").Text
   erp_txt_118 = oNode_erp.selectSingleNode("txt_118").Text
   erp_txt_119 = oNode_erp.selectSingleNode("txt_119").Text
    
   erp_txt_120 = oNode_erp.selectSingleNode("txt_120").Text
   erp_txt_121 = oNode_erp.selectSingleNode("txt_121").Text
   erp_txt_122 = oNode_erp.selectSingleNode("txt_122").Text
   erp_txt_123 = oNode_erp.selectSingleNode("txt_123").Text
   erp_txt_124 = oNode_erp.selectSingleNode("txt_124").Text
   erp_txt_125 = oNode_erp.selectSingleNode("txt_125").Text
   erp_txt_126 = oNode_erp.selectSingleNode("txt_126").Text
   erp_txt_127 = oNode_erp.selectSingleNode("txt_127").Text
   erp_txt_128 = oNode_erp.selectSingleNode("txt_128").Text
   erp_txt_129 = oNode_erp.selectSingleNode("txt_129").Text

   erp_txt_130 = oNode_erp.selectSingleNode("txt_130").Text
   erp_txt_131 = oNode_erp.selectSingleNode("txt_131").Text
   erp_txt_132 = oNode_erp.selectSingleNode("txt_132").Text
   erp_txt_133 = oNode_erp.selectSingleNode("txt_133").Text
   erp_txt_134 = oNode_erp.selectSingleNode("txt_134").Text
   erp_txt_135 = oNode_erp.selectSingleNode("txt_135").Text
   erp_txt_136 = oNode_erp.selectSingleNode("txt_136").Text
   erp_txt_137 = oNode_erp.selectSingleNode("txt_137").Text
   erp_txt_138 = oNode_erp.selectSingleNode("txt_138").Text
   erp_txt_139 = oNode_erp.selectSingleNode("txt_139").Text

   erp_txt_140 = oNode_erp.selectSingleNode("txt_140").Text
   erp_txt_141 = oNode_erp.selectSingleNode("txt_141").Text
   erp_txt_142 = oNode_erp.selectSingleNode("txt_142").Text
   erp_txt_143 = oNode_erp.selectSingleNode("txt_143").Text
   erp_txt_144 = oNode_erp.selectSingleNode("txt_144").Text
   erp_txt_145 = oNode_erp.selectSingleNode("txt_145").Text
   erp_txt_146 = oNode_erp.selectSingleNode("txt_146").Text
   erp_txt_147 = oNode_erp.selectSingleNode("txt_147").Text
   erp_txt_148 = oNode_erp.selectSingleNode("txt_148").Text
   erp_txt_149 = oNode_erp.selectSingleNode("txt_149").Text

   erp_txt_150 = oNode_erp.selectSingleNode("txt_150").Text
   erp_txt_151 = oNode_erp.selectSingleNode("txt_151").Text
   erp_txt_152 = oNode_erp.selectSingleNode("txt_152").Text
   erp_txt_153 = oNode_erp.selectSingleNode("txt_153").Text
   erp_txt_154 = oNode_erp.selectSingleNode("txt_154").Text
   erp_txt_155 = oNode_erp.selectSingleNode("txt_155").Text
   erp_txt_156 = oNode_erp.selectSingleNode("txt_156").Text
   erp_txt_157 = oNode_erp.selectSingleNode("txt_157").Text
   erp_txt_158 = oNode_erp.selectSingleNode("txt_158").Text
   erp_txt_159 = oNode_erp.selectSingleNode("txt_159").Text

   erp_txt_160 = oNode_erp.selectSingleNode("txt_160").Text
   erp_txt_161 = oNode_erp.selectSingleNode("txt_161").Text
   erp_txt_162 = oNode_erp.selectSingleNode("txt_162").Text
   erp_txt_163 = oNode_erp.selectSingleNode("txt_163").Text
   erp_txt_164 = oNode_erp.selectSingleNode("txt_164").Text
   erp_txt_165 = oNode_erp.selectSingleNode("txt_165").Text
   erp_txt_166 = oNode_erp.selectSingleNode("txt_166").Text
   erp_txt_167 = oNode_erp.selectSingleNode("txt_167").Text
   erp_txt_168 = oNode_erp.selectSingleNode("txt_168").Text
   erp_txt_169 = oNode_erp.selectSingleNode("txt_160").Text

   erp_txt_170 = oNode_erp.selectSingleNode("txt_170").Text
   erp_txt_171 = oNode_erp.selectSingleNode("txt_171").Text
   erp_txt_172 = oNode_erp.selectSingleNode("txt_172").Text
   erp_txt_173 = oNode_erp.selectSingleNode("txt_173").Text
   erp_txt_174 = oNode_erp.selectSingleNode("txt_174").Text
   erp_txt_175 = oNode_erp.selectSingleNode("txt_175").Text
   erp_txt_176 = oNode_erp.selectSingleNode("txt_176").Text
   erp_txt_177 = oNode_erp.selectSingleNode("txt_177").Text
   erp_txt_178 = oNode_erp.selectSingleNode("txt_178").Text
   erp_txt_179 = oNode_erp.selectSingleNode("txt_179").Text

   erp_txt_180 = oNode_erp.selectSingleNode("txt_180").Text
   erp_txt_181 = oNode_erp.selectSingleNode("txt_181").Text
   erp_txt_182 = oNode_erp.selectSingleNode("txt_182").Text
   erp_txt_183 = oNode_erp.selectSingleNode("txt_183").Text
   erp_txt_184 = oNode_erp.selectSingleNode("txt_184").Text
   erp_txt_185 = oNode_erp.selectSingleNode("txt_185").Text
   erp_txt_186 = oNode_erp.selectSingleNode("txt_186").Text
   erp_txt_187 = oNode_erp.selectSingleNode("txt_187").Text
   erp_txt_188 = oNode_erp.selectSingleNode("txt_188").Text
   erp_txt_189 = oNode_erp.selectSingleNode("txt_189").Text

   erp_txt_190 = oNode_erp.selectSingleNode("txt_190").Text
   erp_txt_191 = oNode_erp.selectSingleNode("txt_191").Text
   erp_txt_192 = oNode_erp.selectSingleNode("txt_192").Text
   erp_txt_193 = oNode_erp.selectSingleNode("txt_193").Text
   erp_txt_194 = oNode_erp.selectSingleNode("txt_194").Text
   erp_txt_195 = oNode_erp.selectSingleNode("txt_195").Text
   erp_txt_196 = oNode_erp.selectSingleNode("txt_196").Text
   erp_txt_197 = oNode_erp.selectSingleNode("txt_197").Text
   erp_txt_198 = oNode_erp.selectSingleNode("txt_198").Text
   erp_txt_199 = oNode_erp.selectSingleNode("txt_199").Text
    
   erp_txt_200 = oNode_erp.selectSingleNode("txt_200").Text
   erp_txt_201 = oNode_erp.selectSingleNode("txt_201").Text
   erp_txt_202 = oNode_erp.selectSingleNode("txt_202").Text
   erp_txt_203 = oNode_erp.selectSingleNode("txt_203").Text
   erp_txt_204 = oNode_erp.selectSingleNode("txt_204").Text
   erp_txt_205 = oNode_erp.selectSingleNode("txt_205").Text
   erp_txt_206 = oNode_erp.selectSingleNode("txt_206").Text
   erp_txt_207 = oNode_erp.selectSingleNode("txt_207").Text
   erp_txt_208 = oNode_erp.selectSingleNode("txt_208").Text
   erp_txt_209 = oNode_erp.selectSingleNode("txt_209").Text
    
   erp_txt_210 = oNode_erp.selectSingleNode("txt_210").Text
   erp_txt_211 = oNode_erp.selectSingleNode("txt_211").Text
   erp_txt_212 = oNode_erp.selectSingleNode("txt_212").Text
   erp_txt_213 = oNode_erp.selectSingleNode("txt_213").Text
   erp_txt_214 = oNode_erp.selectSingleNode("txt_214").Text
   erp_txt_215 = oNode_erp.selectSingleNode("txt_215").Text
   erp_txt_216 = oNode_erp.selectSingleNode("txt_216").Text
   erp_txt_217 = oNode_erp.selectSingleNode("txt_217").Text
   erp_txt_218 = oNode_erp.selectSingleNode("txt_218").Text
   erp_txt_219 = oNode_erp.selectSingleNode("txt_219").Text

   erp_txt_220 = oNode_erp.selectSingleNode("txt_220").Text
   erp_txt_221 = oNode_erp.selectSingleNode("txt_221").Text

       erp_txt_222 = oNode_erp.selectSingleNode("txt_222").Text
       erp_txt_223 = oNode_erp.selectSingleNode("txt_223").Text
       erp_txt_224 = oNode_erp.selectSingleNode("txt_224").Text
       erp_txt_225 = oNode_erp.selectSingleNode("txt_225").Text
       erp_txt_226 = oNode_erp.selectSingleNode("txt_226").Text
       erp_txt_227 = oNode_erp.selectSingleNode("txt_227").Text
       erp_txt_228 = oNode_erp.selectSingleNode("txt_228").Text
       erp_txt_229 = oNode_erp.selectSingleNode("txt_229").Text

       erp_txt_230 = oNode_erp.selectSingleNode("txt_230").Text
       erp_txt_231 = oNode_erp.selectSingleNode("txt_231").Text
       erp_txt_232 = oNode_erp.selectSingleNode("txt_232").Text
       erp_txt_233 = oNode_erp.selectSingleNode("txt_233").Text
       erp_txt_234 = oNode_erp.selectSingleNode("txt_234").Text
       erp_txt_235 = oNode_erp.selectSingleNode("txt_235").Text
       erp_txt_236 = oNode_erp.selectSingleNode("txt_236").Text
       erp_txt_237 = oNode_erp.selectSingleNode("txt_237").Text
       erp_txt_238 = oNode_erp.selectSingleNode("txt_238").Text
       erp_txt_239 = oNode_erp.selectSingleNode("txt_239").Text

       erp_txt_240 = oNode_erp.selectSingleNode("txt_240").Text
       erp_txt_241 = oNode_erp.selectSingleNode("txt_241").Text
       erp_txt_242 = oNode_erp.selectSingleNode("txt_242").Text
       erp_txt_243 = oNode_erp.selectSingleNode("txt_243").Text
       erp_txt_244 = oNode_erp.selectSingleNode("txt_244").Text
       erp_txt_245 = oNode_erp.selectSingleNode("txt_245").Text
       erp_txt_246 = oNode_erp.selectSingleNode("txt_246").Text
       erp_txt_247 = oNode_erp.selectSingleNode("txt_247").Text
       erp_txt_248 = oNode_erp.selectSingleNode("txt_248").Text
       erp_txt_249 = oNode_erp.selectSingleNode("txt_249").Text

       erp_txt_250 = oNode_erp.selectSingleNode("txt_250").Text
       erp_txt_251 = oNode_erp.selectSingleNode("txt_251").Text
       erp_txt_252 = oNode_erp.selectSingleNode("txt_252").Text
       erp_txt_253 = oNode_erp.selectSingleNode("txt_253").Text
       erp_txt_254 = oNode_erp.selectSingleNode("txt_254").Text
       erp_txt_255 = oNode_erp.selectSingleNode("txt_255").Text
       erp_txt_256 = oNode_erp.selectSingleNode("txt_256").Text
       erp_txt_257 = oNode_erp.selectSingleNode("txt_257").Text
       erp_txt_258 = oNode_erp.selectSingleNode("txt_258").Text
       erp_txt_259 = oNode_erp.selectSingleNode("txt_259").Text
       
       erp_txt_260 = oNode_erp.selectSingleNode("txt_260").Text
       erp_txt_261 = oNode_erp.selectSingleNode("txt_261").Text
       erp_txt_262 = oNode_erp.selectSingleNode("txt_262").Text
       erp_txt_263 = oNode_erp.selectSingleNode("txt_263").Text
       erp_txt_264 = oNode_erp.selectSingleNode("txt_264").Text
       erp_txt_265 = oNode_erp.selectSingleNode("txt_265").Text
       erp_txt_266 = oNode_erp.selectSingleNode("txt_266").Text
       erp_txt_267 = oNode_erp.selectSingleNode("txt_267").Text
       erp_txt_268 = oNode_erp.selectSingleNode("txt_268").Text
       erp_txt_269 = oNode_erp.selectSingleNode("txt_269").Text

       erp_txt_270 = oNode_erp.selectSingleNode("txt_270").Text
       erp_txt_271 = oNode_erp.selectSingleNode("txt_271").Text
       erp_txt_272 = oNode_erp.selectSingleNode("txt_272").Text
       erp_txt_273 = oNode_erp.selectSingleNode("txt_273").Text
       erp_txt_274 = oNode_erp.selectSingleNode("txt_274").Text
       erp_txt_275 = oNode_erp.selectSingleNode("txt_275").Text
       erp_txt_276 = oNode_erp.selectSingleNode("txt_276").Text
       erp_txt_277 = oNode_erp.selectSingleNode("txt_277").Text
       erp_txt_278 = oNode_erp.selectSingleNode("txt_278").Text
       erp_txt_279 = oNode_erp.selectSingleNode("txt_279").Text
    
       erp_txt_280 = oNode_erp.selectSingleNode("txt_280").Text
       erp_txt_281 = oNode_erp.selectSingleNode("txt_281").Text
       erp_txt_282 = oNode_erp.selectSingleNode("txt_282").Text
       erp_txt_283 = oNode_erp.selectSingleNode("txt_283").Text
       erp_txt_284 = oNode_erp.selectSingleNode("txt_284").Text
       erp_txt_285 = oNode_erp.selectSingleNode("txt_285").Text
       erp_txt_286 = oNode_erp.selectSingleNode("txt_286").Text
       erp_txt_287 = oNode_erp.selectSingleNode("txt_287").Text
       
       erp_txt_288 = oNode_erp.selectSingleNode("txt_288").Text
       erp_txt_289 = oNode_erp.selectSingleNode("txt_289").Text
       erp_txt_290 = oNode_erp.selectSingleNode("txt_290").Text

        erp_txt_291 = oNode_erp.selectSingleNode("txt_291").Text
        erp_txt_292 = oNode_erp.selectSingleNode("txt_292").Text
        erp_txt_293 = oNode_erp.selectSingleNode("txt_293").Text
        erp_txt_294 = oNode_erp.selectSingleNode("txt_294").Text
        erp_txt_295 = oNode_erp.selectSingleNode("txt_295").Text
        erp_txt_296 = oNode_erp.selectSingleNode("txt_296").Text
        erp_txt_297 = oNode_erp.selectSingleNode("txt_297").Text
        erp_txt_298 = oNode_erp.selectSingleNode("txt_298").Text
        erp_txt_299 = oNode_erp.selectSingleNode("txt_299").Text
        erp_txt_300 = oNode_erp.selectSingleNode("txt_300").Text
        erp_txt_301 = oNode_erp.selectSingleNode("txt_301").Text
        erp_txt_302 = oNode_erp.selectSingleNode("txt_302").Text
        erp_txt_303 = oNode_erp.selectSingleNode("txt_303").Text
        erp_txt_304 = oNode_erp.selectSingleNode("txt_304").Text
        erp_txt_305 = oNode_erp.selectSingleNode("txt_305").Text
        erp_txt_306 = oNode_erp.selectSingleNode("txt_306").Text
        erp_txt_307 = oNode_erp.selectSingleNode("txt_307").Text
        erp_txt_308 = oNode_erp.selectSingleNode("txt_308").Text
        erp_txt_309 = oNode_erp.selectSingleNode("txt_309").Text
        erp_txt_310 = oNode_erp.selectSingleNode("txt_310").Text
        erp_txt_311 = oNode_erp.selectSingleNode("txt_311").Text
        erp_txt_312 = oNode_erp.selectSingleNode("txt_312").Text
        erp_txt_313 = oNode_erp.selectSingleNode("txt_313").Text
        erp_txt_314 = oNode_erp.selectSingleNode("txt_314").Text
        erp_txt_315 = oNode_erp.selectSingleNode("txt_315").Text
        erp_txt_316 = oNode_erp.selectSingleNode("txt_316").Text
             
        erp_txt_317 = oNode_erp.selectSingleNode("txt_317").Text
        erp_txt_318 = oNode_erp.selectSingleNode("txt_318").Text
             
        erp_txt_319 = oNode_erp.selectSingleNode("txt_319").Text
        erp_txt_320 = oNode_erp.selectSingleNode("txt_320").Text
        erp_txt_321 = oNode_erp.selectSingleNode("txt_321").Text
        erp_txt_322 = oNode_erp.selectSingleNode("txt_322").Text
        erp_txt_323 = oNode_erp.selectSingleNode("txt_323").Text
        erp_txt_324 = oNode_erp.selectSingleNode("txt_324").Text
        erp_txt_325 = oNode_erp.selectSingleNode("txt_325").Text
        erp_txt_326 = oNode_erp.selectSingleNode("txt_326").Text
        erp_txt_327 = oNode_erp.selectSingleNode("txt_327").Text
        erp_txt_328 = oNode_erp.selectSingleNode("txt_328").Text
        erp_txt_329 = oNode_erp.selectSingleNode("txt_329").Text
              
        erp_txt_330 = oNode_erp.selectSingleNode("txt_330").Text
        erp_txt_331 = oNode_erp.selectSingleNode("txt_331").Text
        erp_txt_332 = oNode_erp.selectSingleNode("txt_332").Text
        erp_txt_333 = oNode_erp.selectSingleNode("txt_333").Text
        erp_txt_334 = oNode_erp.selectSingleNode("txt_334").Text
        erp_txt_335 = oNode_erp.selectSingleNode("txt_335").Text
        erp_txt_336 = oNode_erp.selectSingleNode("txt_336").Text
        erp_txt_337 = oNode_erp.selectSingleNode("txt_337").Text
        erp_txt_338 = oNode_erp.selectSingleNode("txt_338").Text
        erp_txt_339 = oNode_erp.selectSingleNode("txt_339").Text
        erp_txt_340 = oNode_erp.selectSingleNode("txt_340").Text
              
        erp_txt_341 = oNode_erp.selectSingleNode("txt_341").Text
        erp_txt_342 = oNode_erp.selectSingleNode("txt_342").Text
        erp_txt_343 = oNode_erp.selectSingleNode("txt_343").Text
        erp_txt_344 = oNode_erp.selectSingleNode("txt_344").Text
        erp_txt_345 = oNode_erp.selectSingleNode("txt_345").Text
        erp_txt_346 = oNode_erp.selectSingleNode("txt_346").Text
              
              erp_txt_347 = oNode_erp.selectSingleNode("txt_347").Text
              erp_txt_348 = oNode_erp.selectSingleNode("txt_348").Text
              erp_txt_349 = oNode_erp.selectSingleNode("txt_349").Text
              erp_txt_350 = oNode_erp.selectSingleNode("txt_350").Text
              erp_txt_351 = oNode_erp.selectSingleNode("txt_351").Text
              erp_txt_352 = oNode_erp.selectSingleNode("txt_352").Text
              erp_txt_353 = oNode_erp.selectSingleNode("txt_353").Text
              
              erp_txt_354 = oNode_erp.selectSingleNode("txt_354").Text
              erp_txt_355 = oNode_erp.selectSingleNode("txt_355").Text
              erp_txt_356 = oNode_erp.selectSingleNode("txt_356").Text
              erp_txt_357 = oNode_erp.selectSingleNode("txt_357").Text
              erp_txt_358 = oNode_erp.selectSingleNode("txt_358").Text
              erp_txt_359 = oNode_erp.selectSingleNode("txt_359").Text
              erp_txt_360 = oNode_erp.selectSingleNode("txt_360").Text
              erp_txt_361 = oNode_erp.selectSingleNode("txt_361").Text
              erp_txt_362 = oNode_erp.selectSingleNode("txt_362").Text
              erp_txt_363 = oNode_erp.selectSingleNode("txt_363").Text
              erp_txt_364 = oNode_erp.selectSingleNode("txt_364").Text
              
              erp_txt_365 = oNode_erp.selectSingleNode("txt_365").Text
              erp_txt_366 = oNode_erp.selectSingleNode("txt_366").Text
              erp_txt_367 = oNode_erp.selectSingleNode("txt_367").Text
              erp_txt_368 = oNode_erp.selectSingleNode("txt_368").Text
              erp_txt_369 = oNode_erp.selectSingleNode("txt_369").Text
              erp_txt_370 = oNode_erp.selectSingleNode("txt_370").Text
              erp_txt_371 = oNode_erp.selectSingleNode("txt_371").Text
              erp_txt_372 = oNode_erp.selectSingleNode("txt_372").Text
              erp_txt_373 = oNode_erp.selectSingleNode("txt_373").Text
              erp_txt_374 = oNode_erp.selectSingleNode("txt_374").Text
              erp_txt_375 = oNode_erp.selectSingleNode("txt_375").Text
              erp_txt_376 = oNode_erp.selectSingleNode("txt_376").Text

              erp_txt_377 = oNode_erp.selectSingleNode("txt_377").Text
              erp_txt_378 = oNode_erp.selectSingleNode("txt_378").Text
              erp_txt_379 = oNode_erp.selectSingleNode("txt_379").Text
              erp_txt_380 = oNode_erp.selectSingleNode("txt_380").Text
              erp_txt_381 = oNode_erp.selectSingleNode("txt_381").Text
              erp_txt_382 = oNode_erp.selectSingleNode("txt_382").Text
              erp_txt_383 = oNode_erp.selectSingleNode("txt_383").Text
                
              erp_txt_384 = oNode_erp.selectSingleNode("txt_384").Text
              erp_txt_385 = oNode_erp.selectSingleNode("txt_385").Text
              erp_txt_386 = oNode_erp.selectSingleNode("txt_386").Text
              erp_txt_387 = oNode_erp.selectSingleNode("txt_387").Text
              erp_txt_388 = oNode_erp.selectSingleNode("txt_388").Text

              erp_txt_389 = oNode_erp.selectSingleNode("txt_389").Text
              erp_txt_390 = oNode_erp.selectSingleNode("txt_390").Text

              erp_txt_391 = oNode_erp.selectSingleNode("txt_391").Text
              erp_txt_392 = oNode_erp.selectSingleNode("txt_392").Text
              erp_txt_393 = oNode_erp.selectSingleNode("txt_393").Text

              erp_txt_394 = oNode_erp.selectSingleNode("txt_394").Text
              erp_txt_395 = oNode_erp.selectSingleNode("txt_395").Text
              erp_txt_396 = oNode_erp.selectSingleNode("txt_396").Text
        
                erp_txt_397 = oNode_erp.selectSingleNode("txt_397").Text
                erp_txt_398 = oNode_erp.selectSingleNode("txt_398").Text
                erp_txt_399 = oNode_erp.selectSingleNode("txt_399").Text
                erp_txt_400 = oNode_erp.selectSingleNode("txt_400").Text
                erp_txt_401 = oNode_erp.selectSingleNode("txt_401").Text

              erp_txt_402 = oNode_erp.selectSingleNode("txt_402").Text

             erp_txt_403 = oNode_erp.selectSingleNode("txt_403").Text
             erp_txt_404 = oNode_erp.selectSingleNode("txt_404").Text

             erp_txt_405 = oNode_erp.selectSingleNode("txt_405").Text
             erp_txt_406 = oNode_erp.selectSingleNode("txt_406").Text
             erp_txt_407 = oNode_erp.selectSingleNode("txt_407").Text
             erp_txt_408 = oNode_erp.selectSingleNode("txt_408").Text
             erp_txt_409 = oNode_erp.selectSingleNode("txt_409").Text
             erp_txt_410 = oNode_erp.selectSingleNode("txt_410").Text
            erp_txt_411 = oNode_erp.selectSingleNode("txt_411").Text
    
            erp_txt_412 = oNode_erp.selectSingleNode("txt_412").Text
            erp_txt_413 = oNode_erp.selectSingleNode("txt_413").Text
            erp_txt_414 = oNode_erp.selectSingleNode("txt_414").Text
            erp_txt_415 = oNode_erp.selectSingleNode("txt_415").Text
            erp_txt_416 = oNode_erp.selectSingleNode("txt_416").Text
            erp_txt_417 = oNode_erp.selectSingleNode("txt_417").Text
            erp_txt_418 = oNode_erp.selectSingleNode("txt_418").Text
            erp_txt_419 = oNode_erp.selectSingleNode("txt_419").Text
            erp_txt_420 = oNode_erp.selectSingleNode("txt_420").Text

             erp_txt_421 = oNode_erp.selectSingleNode("txt_421").Text
            erp_txt_422 = oNode_erp.selectSingleNode("txt_422").Text
            erp_txt_423 = oNode_erp.selectSingleNode("txt_423").Text
            erp_txt_424 = oNode_erp.selectSingleNode("txt_424").Text
            erp_txt_425 = oNode_erp.selectSingleNode("txt_425").Text

            erp_txt_426 = oNode_erp.selectSingleNode("txt_426").Text

            erp_txt_427 = oNode_erp.selectSingleNode("txt_427").Text
            erp_txt_428 = oNode_erp.selectSingleNode("txt_428").Text
            erp_txt_429 = oNode_erp.selectSingleNode("txt_429").Text
            erp_txt_430 = oNode_erp.selectSingleNode("txt_430").Text

            erp_txt_431 = oNode_erp.selectSingleNode("txt_431").Text
            erp_txt_432 = oNode_erp.selectSingleNode("txt_432").Text

            erp_txt_433 = oNode_erp.selectSingleNode("txt_433").Text
            erp_txt_434 = oNode_erp.selectSingleNode("txt_434").Text
            erp_txt_435 = oNode_erp.selectSingleNode("txt_435").Text
            erp_txt_436 = oNode_erp.selectSingleNode("txt_436").Text
            erp_txt_437 = oNode_erp.selectSingleNode("txt_437").Text
            erp_txt_438 = oNode_erp.selectSingleNode("txt_438").Text
            erp_txt_439 = oNode_erp.selectSingleNode("txt_439").Text

            erp_txt_440 = oNode_erp.selectSingleNode("txt_440").Text
            erp_txt_441 = oNode_erp.selectSingleNode("txt_441").Text
            erp_txt_442 = oNode_erp.selectSingleNode("txt_442").Text
            erp_txt_443 = oNode_erp.selectSingleNode("txt_443").Text
            erp_txt_444 = oNode_erp.selectSingleNode("txt_444").Text
            erp_txt_445 = oNode_erp.selectSingleNode("txt_445").Text
            erp_txt_446 = oNode_erp.selectSingleNode("txt_446").Text
            erp_txt_447 = oNode_erp.selectSingleNode("txt_447").Text
            erp_txt_448 = oNode_erp.selectSingleNode("txt_448").Text
            erp_txt_449 = oNode_erp.selectSingleNode("txt_449").Text

            erp_txt_450 = oNode_erp.selectSingleNode("txt_450").Text
            erp_txt_451 = oNode_erp.selectSingleNode("txt_451").Text
            erp_txt_452 = oNode_erp.selectSingleNode("txt_452").Text
            erp_txt_453 = oNode_erp.selectSingleNode("txt_453").Text
            erp_txt_454 = oNode_erp.selectSingleNode("txt_454").Text
            erp_txt_455 = oNode_erp.selectSingleNode("txt_455").Text
            erp_txt_456 = oNode_erp.selectSingleNode("txt_456").Text
            erp_txt_457 = oNode_erp.selectSingleNode("txt_457").Text
            erp_txt_458 = oNode_erp.selectSingleNode("txt_458").Text
            erp_txt_459 = oNode_erp.selectSingleNode("txt_459").Text

            erp_txt_460 = oNode_erp.selectSingleNode("txt_460").Text
            erp_txt_461 = oNode_erp.selectSingleNode("txt_461").Text
            erp_txt_462 = oNode_erp.selectSingleNode("txt_462").Text
    
            erp_txt_463 = oNode_erp.selectSingleNode("txt_463").Text
            erp_txt_464 = oNode_erp.selectSingleNode("txt_464").Text
            erp_txt_465 = oNode_erp.selectSingleNode("txt_465").Text
            erp_txt_466 = oNode_erp.selectSingleNode("txt_466").Text
            erp_txt_467 = oNode_erp.selectSingleNode("txt_467").Text
            erp_txt_468 = oNode_erp.selectSingleNode("txt_468").Text
            erp_txt_469 = oNode_erp.selectSingleNode("txt_469").Text

            erp_txt_470 = oNode_erp.selectSingleNode("txt_470").Text
            erp_txt_471 = oNode_erp.selectSingleNode("txt_471").Text
            erp_txt_472 = oNode_erp.selectSingleNode("txt_472").Text
            erp_txt_473 = oNode_erp.selectSingleNode("txt_473").Text
            erp_txt_474 = oNode_erp.selectSingleNode("txt_474").Text
            erp_txt_475 = oNode_erp.selectSingleNode("txt_475").Text
            erp_txt_476 = oNode_erp.selectSingleNode("txt_476").Text
            erp_txt_477 = oNode_erp.selectSingleNode("txt_477").Text
            erp_txt_478 = oNode_erp.selectSingleNode("txt_478").Text
            erp_txt_479 = oNode_erp.selectSingleNode("txt_479").Text

            erp_txt_480 = oNode_erp.selectSingleNode("txt_480").Text
            erp_txt_481 = oNode_erp.selectSingleNode("txt_481").Text
            erp_txt_482 = oNode_erp.selectSingleNode("txt_482").Text
            erp_txt_483 = oNode_erp.selectSingleNode("txt_483").Text
            erp_txt_484 = oNode_erp.selectSingleNode("txt_484").Text
            erp_txt_485 = oNode_erp.selectSingleNode("txt_485").Text
            erp_txt_486 = oNode_erp.selectSingleNode("txt_486").Text
            erp_txt_487 = oNode_erp.selectSingleNode("txt_487").Text
            erp_txt_488 = oNode_erp.selectSingleNode("txt_488").Text
            erp_txt_489 = oNode_erp.selectSingleNode("txt_489").Text

            erp_txt_490 = oNode_erp.selectSingleNode("txt_490").Text
            erp_txt_491 = oNode_erp.selectSingleNode("txt_491").Text
            erp_txt_492 = oNode_erp.selectSingleNode("txt_492").Text
            erp_txt_493 = oNode_erp.selectSingleNode("txt_493").Text
            erp_txt_494 = oNode_erp.selectSingleNode("txt_494").Text
            erp_txt_495 = oNode_erp.selectSingleNode("txt_495").Text
            erp_txt_496 = oNode_erp.selectSingleNode("txt_496").Text
            erp_txt_497 = oNode_erp.selectSingleNode("txt_497").Text
            erp_txt_498 = oNode_erp.selectSingleNode("txt_498").Text
            erp_txt_499 = oNode_erp.selectSingleNode("txt_499").Text

            erp_txt_500 = oNode_erp.selectSingleNode("txt_500").Text
            erp_txt_501 = oNode_erp.selectSingleNode("txt_501").Text
            erp_txt_502 = oNode_erp.selectSingleNode("txt_502").Text
            erp_txt_503 = oNode_erp.selectSingleNode("txt_503").Text
            erp_txt_504 = oNode_erp.selectSingleNode("txt_504").Text
            erp_txt_505 = oNode_erp.selectSingleNode("txt_505").Text
            erp_txt_506 = oNode_erp.selectSingleNode("txt_506").Text 
            erp_txt_507 = oNode_erp.selectSingleNode("txt_507").Text
            erp_txt_508 = oNode_erp.selectSingleNode("txt_508").Text
            erp_txt_509 = oNode_erp.selectSingleNode("txt_509").Text
            erp_txt_510 = oNode_erp.selectSingleNode("txt_510").Text
            erp_txt_511 = oNode_erp.selectSingleNode("txt_511").Text
            erp_txt_512 = oNode_erp.selectSingleNode("txt_512").Text
            erp_txt_513 = oNode_erp.selectSingleNode("txt_513").Text
            erp_txt_514 = oNode_erp.selectSingleNode("txt_514").Text
            erp_txt_515 = oNode_erp.selectSingleNode("txt_515").Text
            erp_txt_516 = oNode_erp.selectSingleNode("txt_516").Text
            erp_txt_517 = oNode_erp.selectSingleNode("txt_517").Text
            erp_txt_518 = oNode_erp.selectSingleNode("txt_518").Text
            erp_txt_519 = oNode_erp.selectSingleNode("txt_519").Text
            erp_txt_520 = oNode_erp.selectSingleNode("txt_520").Text
            erp_txt_521 = oNode_erp.selectSingleNode("txt_521").Text
            erp_txt_522 = oNode_erp.selectSingleNode("txt_522").Text
           

next



'Response.Write "txt_001: " &erp_txt_001 & "<br>"
'Response.Write "txt_002: " &erp_txt_002 & "<br>"


%>