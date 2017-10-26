
<% 
Dim objXMLHTTP_job, objXMLDOM_job, i_job, strHTML_job

Set objXMLDom_job = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_job = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_job.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/job_sprog.xml", False
'objXmlHttp_job.open "GET", "http://localhost/inc/xml/job_sprog.xml", False
'objXmlHttp_job.open "GET", "http://outzource.dk/timeout_xp/wwwroot/ver2_10/inc/xml/job_sprog.xml", False
'objXmlHttp_job.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/job_sprog.xml", False
'objXmlHttp_job.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/job_sprog.xml", False
objXmlHttp_job.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/job_sprog.xml", False

objXmlHttp_job.send


Set objXmlDom_job = objXmlHttp_job.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_job = Nothing



Dim Address_job, Latitude_job, Longitude_job
Dim oNode_job, oNodes_job
Dim sXPathQuery_job

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
sXPathQuery_job = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_job = "//sprog/uk"
'Session.LCID = 2057
case 3
sXPathQuery_job = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_job = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_job = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_job = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_job = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_job = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030



Set oNode_job = objXmlDom_job.documentElement.selectSingleNode(sXPathQuery_job)
Address_job = oNode_job.Text

Set oNodes_job = objXmlDom_job.documentElement.selectNodes(sXPathQuery_job)

    For Each oNode_job in oNodes_job

    job_txt_001 = oNode_job.selectSingleNode("txt_1").Text
    job_txt_002 = oNode_job.selectSingleNode("txt_2").Text
    job_txt_003 = oNode_job.selectSingleNode("txt_3").Text
    job_txt_004 = oNode_job.selectSingleNode("txt_4").Text
    job_txt_005 = oNode_job.selectSingleNode("txt_5").Text
    job_txt_006 = oNode_job.selectSingleNode("txt_6").Text
    job_txt_007 = oNode_job.selectSingleNode("txt_7").Text
    job_txt_008 = oNode_job.selectSingleNode("txt_8").Text
    job_txt_009 = oNode_job.selectSingleNode("txt_9").Text
    job_txt_010 = oNode_job.selectSingleNode("txt_10").Text

    job_txt_011 = oNode_job.selectSingleNode("txt_11").Text
    job_txt_012 = oNode_job.selectSingleNode("txt_12").Text
    job_txt_013 = oNode_job.selectSingleNode("txt_13").Text
    job_txt_014 = oNode_job.selectSingleNode("txt_14").Text
    job_txt_015 = oNode_job.selectSingleNode("txt_15").Text
    job_txt_016 = oNode_job.selectSingleNode("txt_16").Text
    job_txt_017 = oNode_job.selectSingleNode("txt_17").Text
    job_txt_018 = oNode_job.selectSingleNode("txt_18").Text
    job_txt_019 = oNode_job.selectSingleNode("txt_19").Text
    job_txt_020 = oNode_job.selectSingleNode("txt_20").Text

    job_txt_021 = oNode_job.selectSingleNode("txt_21").Text
    job_txt_022 = oNode_job.selectSingleNode("txt_22").Text
    job_txt_023 = oNode_job.selectSingleNode("txt_23").Text
    job_txt_024 = oNode_job.selectSingleNode("txt_24").Text
    job_txt_025 = oNode_job.selectSingleNode("txt_25").Text
    job_txt_026 = oNode_job.selectSingleNode("txt_26").Text
    job_txt_027 = oNode_job.selectSingleNode("txt_27").Text
    job_txt_028 = oNode_job.selectSingleNode("txt_28").Text
    job_txt_029 = oNode_job.selectSingleNode("txt_29").Text
    job_txt_030 = oNode_job.selectSingleNode("txt_30").Text

    job_txt_031 = oNode_job.selectSingleNode("txt_31").Text
    job_txt_032 = oNode_job.selectSingleNode("txt_32").Text
    job_txt_033 = oNode_job.selectSingleNode("txt_33").Text
    job_txt_034 = oNode_job.selectSingleNode("txt_34").Text
    job_txt_035 = oNode_job.selectSingleNode("txt_35").Text
    job_txt_036 = oNode_job.selectSingleNode("txt_36").Text
    job_txt_037 = oNode_job.selectSingleNode("txt_37").Text
    job_txt_038 = oNode_job.selectSingleNode("txt_38").Text
    job_txt_039 = oNode_job.selectSingleNode("txt_39").Text
    job_txt_040 = oNode_job.selectSingleNode("txt_40").Text

    job_txt_041 = oNode_job.selectSingleNode("txt_41").Text
    job_txt_042 = oNode_job.selectSingleNode("txt_42").Text
    job_txt_043 = oNode_job.selectSingleNode("txt_43").Text
    job_txt_044 = oNode_job.selectSingleNode("txt_44").Text
    job_txt_045 = oNode_job.selectSingleNode("txt_45").Text
    job_txt_046 = oNode_job.selectSingleNode("txt_46").Text
    job_txt_047 = oNode_job.selectSingleNode("txt_47").Text
    job_txt_048 = oNode_job.selectSingleNode("txt_48").Text
    job_txt_049 = oNode_job.selectSingleNode("txt_49").Text
    job_txt_050 = oNode_job.selectSingleNode("txt_50").Text

    job_txt_051 = oNode_job.selectSingleNode("txt_51").Text
    job_txt_052 = oNode_job.selectSingleNode("txt_52").Text
    job_txt_053 = oNode_job.selectSingleNode("txt_53").Text
    job_txt_054 = oNode_job.selectSingleNode("txt_54").Text
    job_txt_055 = oNode_job.selectSingleNode("txt_55").Text
    job_txt_056 = oNode_job.selectSingleNode("txt_56").Text
    job_txt_057 = oNode_job.selectSingleNode("txt_57").Text
    job_txt_058 = oNode_job.selectSingleNode("txt_58").Text
    job_txt_059 = oNode_job.selectSingleNode("txt_59").Text
    job_txt_060 = oNode_job.selectSingleNode("txt_60").Text

    job_txt_061 = oNode_job.selectSingleNode("txt_61").Text
    job_txt_062 = oNode_job.selectSingleNode("txt_62").Text
    job_txt_063 = oNode_job.selectSingleNode("txt_63").Text
    job_txt_064 = oNode_job.selectSingleNode("txt_64").Text
    job_txt_065 = oNode_job.selectSingleNode("txt_65").Text
    job_txt_066 = oNode_job.selectSingleNode("txt_66").Text
    job_txt_067 = oNode_job.selectSingleNode("txt_67").Text
    job_txt_068 = oNode_job.selectSingleNode("txt_68").Text
    job_txt_069 = oNode_job.selectSingleNode("txt_69").Text
    job_txt_070 = oNode_job.selectSingleNode("txt_70").Text

    job_txt_071 = oNode_job.selectSingleNode("txt_71").Text
    job_txt_072 = oNode_job.selectSingleNode("txt_72").Text
    job_txt_073 = oNode_job.selectSingleNode("txt_73").Text
    job_txt_074 = oNode_job.selectSingleNode("txt_74").Text
    job_txt_075 = oNode_job.selectSingleNode("txt_75").Text
    job_txt_076 = oNode_job.selectSingleNode("txt_76").Text
    job_txt_077 = oNode_job.selectSingleNode("txt_77").Text
    job_txt_078 = oNode_job.selectSingleNode("txt_78").Text
    job_txt_079 = oNode_job.selectSingleNode("txt_79").Text
    job_txt_080 = oNode_job.selectSingleNode("txt_80").Text

    job_txt_081 = oNode_job.selectSingleNode("txt_81").Text
    job_txt_082 = oNode_job.selectSingleNode("txt_82").Text
    job_txt_083 = oNode_job.selectSingleNode("txt_83").Text
    job_txt_084 = oNode_job.selectSingleNode("txt_84").Text
    job_txt_085 = oNode_job.selectSingleNode("txt_85").Text
    job_txt_086 = oNode_job.selectSingleNode("txt_86").Text
    job_txt_087 = oNode_job.selectSingleNode("txt_87").Text
    job_txt_088 = oNode_job.selectSingleNode("txt_88").Text
    job_txt_089 = oNode_job.selectSingleNode("txt_89").Text
    job_txt_090 = oNode_job.selectSingleNode("txt_90").Text

    job_txt_091 = oNode_job.selectSingleNode("txt_91").Text
    job_txt_092 = oNode_job.selectSingleNode("txt_92").Text
    job_txt_093 = oNode_job.selectSingleNode("txt_93").Text
    job_txt_094 = oNode_job.selectSingleNode("txt_94").Text
    job_txt_095 = oNode_job.selectSingleNode("txt_95").Text
    job_txt_096 = oNode_job.selectSingleNode("txt_96").Text
    job_txt_097 = oNode_job.selectSingleNode("txt_97").Text
    job_txt_098 = oNode_job.selectSingleNode("txt_98").Text
    job_txt_099 = oNode_job.selectSingleNode("txt_99").Text
    job_txt_100 = oNode_job.selectSingleNode("txt_100").Text

    job_txt_101 = oNode_job.selectSingleNode("txt_101").Text
    job_txt_102 = oNode_job.selectSingleNode("txt_102").Text
    job_txt_103 = oNode_job.selectSingleNode("txt_103").Text
    job_txt_104 = oNode_job.selectSingleNode("txt_104").Text
    job_txt_105 = oNode_job.selectSingleNode("txt_105").Text
    job_txt_106 = oNode_job.selectSingleNode("txt_106").Text
    job_txt_107 = oNode_job.selectSingleNode("txt_107").Text
    job_txt_108 = oNode_job.selectSingleNode("txt_108").Text
    job_txt_109 = oNode_job.selectSingleNode("txt_109").Text
    job_txt_110 = oNode_job.selectSingleNode("txt_110").Text

    job_txt_111 = oNode_job.selectSingleNode("txt_111").Text
    job_txt_112 = oNode_job.selectSingleNode("txt_112").Text
    job_txt_113 = oNode_job.selectSingleNode("txt_113").Text
    job_txt_114 = oNode_job.selectSingleNode("txt_114").Text
    job_txt_115 = oNode_job.selectSingleNode("txt_115").Text
    job_txt_116 = oNode_job.selectSingleNode("txt_116").Text
    job_txt_117 = oNode_job.selectSingleNode("txt_117").Text
    job_txt_118 = oNode_job.selectSingleNode("txt_118").Text
    job_txt_119 = oNode_job.selectSingleNode("txt_119").Text
    job_txt_120 = oNode_job.selectSingleNode("txt_120").Text

    job_txt_121 = oNode_job.selectSingleNode("txt_121").Text
    job_txt_122 = oNode_job.selectSingleNode("txt_122").Text
    job_txt_123 = oNode_job.selectSingleNode("txt_123").Text
    job_txt_124 = oNode_job.selectSingleNode("txt_124").Text
    job_txt_125 = oNode_job.selectSingleNode("txt_125").Text
    job_txt_126 = oNode_job.selectSingleNode("txt_126").Text
    job_txt_127 = oNode_job.selectSingleNode("txt_127").Text
    job_txt_128 = oNode_job.selectSingleNode("txt_128").Text
    job_txt_129 = oNode_job.selectSingleNode("txt_129").Text
    job_txt_130 = oNode_job.selectSingleNode("txt_130").Text

    job_txt_131 = oNode_job.selectSingleNode("txt_131").Text
    job_txt_132 = oNode_job.selectSingleNode("txt_132").Text
    job_txt_133 = oNode_job.selectSingleNode("txt_133").Text
    job_txt_134 = oNode_job.selectSingleNode("txt_134").Text
    job_txt_135 = oNode_job.selectSingleNode("txt_135").Text
    job_txt_136 = oNode_job.selectSingleNode("txt_136").Text
    job_txt_137 = oNode_job.selectSingleNode("txt_137").Text
    job_txt_138 = oNode_job.selectSingleNode("txt_138").Text
    job_txt_139 = oNode_job.selectSingleNode("txt_139").Text
    job_txt_140 = oNode_job.selectSingleNode("txt_140").Text

    job_txt_141 = oNode_job.selectSingleNode("txt_141").Text
    job_txt_142 = oNode_job.selectSingleNode("txt_142").Text
    job_txt_143 = oNode_job.selectSingleNode("txt_143").Text
    job_txt_144 = oNode_job.selectSingleNode("txt_144").Text
    job_txt_145 = oNode_job.selectSingleNode("txt_145").Text
    job_txt_146 = oNode_job.selectSingleNode("txt_146").Text
    job_txt_147 = oNode_job.selectSingleNode("txt_147").Text
    job_txt_148 = oNode_job.selectSingleNode("txt_148").Text
    job_txt_149 = oNode_job.selectSingleNode("txt_149").Text
    job_txt_150 = oNode_job.selectSingleNode("txt_150").Text

    job_txt_151 = oNode_job.selectSingleNode("txt_151").Text
    job_txt_152 = oNode_job.selectSingleNode("txt_152").Text
    job_txt_153 = oNode_job.selectSingleNode("txt_153").Text
    job_txt_154 = oNode_job.selectSingleNode("txt_154").Text
    job_txt_155 = oNode_job.selectSingleNode("txt_155").Text
    job_txt_156 = oNode_job.selectSingleNode("txt_156").Text
    job_txt_157 = oNode_job.selectSingleNode("txt_157").Text
    job_txt_158 = oNode_job.selectSingleNode("txt_158").Text
    job_txt_159 = oNode_job.selectSingleNode("txt_159").Text
    job_txt_160 = oNode_job.selectSingleNode("txt_160").Text

    job_txt_161 = oNode_job.selectSingleNode("txt_161").Text
    job_txt_162 = oNode_job.selectSingleNode("txt_162").Text
    job_txt_163 = oNode_job.selectSingleNode("txt_163").Text
    job_txt_164 = oNode_job.selectSingleNode("txt_164").Text
    job_txt_165 = oNode_job.selectSingleNode("txt_165").Text
    job_txt_166 = oNode_job.selectSingleNode("txt_166").Text
    job_txt_167 = oNode_job.selectSingleNode("txt_167").Text
    job_txt_168 = oNode_job.selectSingleNode("txt_168").Text
    job_txt_169 = oNode_job.selectSingleNode("txt_169").Text
    job_txt_170 = oNode_job.selectSingleNode("txt_170").Text

    job_txt_171 = oNode_job.selectSingleNode("txt_171").Text
    job_txt_172 = oNode_job.selectSingleNode("txt_172").Text
    job_txt_173 = oNode_job.selectSingleNode("txt_173").Text
    job_txt_174 = oNode_job.selectSingleNode("txt_174").Text
    job_txt_175 = oNode_job.selectSingleNode("txt_175").Text
    job_txt_176 = oNode_job.selectSingleNode("txt_176").Text
    job_txt_177 = oNode_job.selectSingleNode("txt_177").Text
    job_txt_178 = oNode_job.selectSingleNode("txt_178").Text
    job_txt_179 = oNode_job.selectSingleNode("txt_179").Text
    job_txt_180 = oNode_job.selectSingleNode("txt_180").Text

    job_txt_181 = oNode_job.selectSingleNode("txt_181").Text
    job_txt_182 = oNode_job.selectSingleNode("txt_182").Text
    job_txt_183 = oNode_job.selectSingleNode("txt_183").Text
    job_txt_184 = oNode_job.selectSingleNode("txt_184").Text
    job_txt_185 = oNode_job.selectSingleNode("txt_185").Text
    job_txt_186 = oNode_job.selectSingleNode("txt_186").Text
    job_txt_187 = oNode_job.selectSingleNode("txt_187").Text
    job_txt_188 = oNode_job.selectSingleNode("txt_188").Text
    job_txt_189 = oNode_job.selectSingleNode("txt_189").Text
    job_txt_190 = oNode_job.selectSingleNode("txt_190").Text

    job_txt_191 = oNode_job.selectSingleNode("txt_191").Text
    job_txt_192 = oNode_job.selectSingleNode("txt_192").Text
    job_txt_193 = oNode_job.selectSingleNode("txt_193").Text
    job_txt_194 = oNode_job.selectSingleNode("txt_194").Text
    job_txt_195 = oNode_job.selectSingleNode("txt_195").Text
    job_txt_196 = oNode_job.selectSingleNode("txt_196").Text
    job_txt_197 = oNode_job.selectSingleNode("txt_197").Text
    job_txt_198 = oNode_job.selectSingleNode("txt_198").Text
    job_txt_199 = oNode_job.selectSingleNode("txt_199").Text
    job_txt_200 = oNode_job.selectSingleNode("txt_200").Text

    job_txt_201 = oNode_job.selectSingleNode("txt_201").Text
    job_txt_202 = oNode_job.selectSingleNode("txt_202").Text
    job_txt_203 = oNode_job.selectSingleNode("txt_203").Text
    job_txt_204 = oNode_job.selectSingleNode("txt_204").Text
    job_txt_205 = oNode_job.selectSingleNode("txt_205").Text
    job_txt_206 = oNode_job.selectSingleNode("txt_206").Text
    job_txt_207 = oNode_job.selectSingleNode("txt_207").Text
    job_txt_208 = oNode_job.selectSingleNode("txt_208").Text
    job_txt_209 = oNode_job.selectSingleNode("txt_209").Text
    job_txt_210 = oNode_job.selectSingleNode("txt_210").Text

    job_txt_211 = oNode_job.selectSingleNode("txt_211").Text
    job_txt_212 = oNode_job.selectSingleNode("txt_212").Text
    job_txt_213 = oNode_job.selectSingleNode("txt_213").Text
    job_txt_214 = oNode_job.selectSingleNode("txt_214").Text
    job_txt_215 = oNode_job.selectSingleNode("txt_215").Text
    job_txt_216 = oNode_job.selectSingleNode("txt_216").Text
    job_txt_217 = oNode_job.selectSingleNode("txt_217").Text
    job_txt_218 = oNode_job.selectSingleNode("txt_218").Text
    job_txt_219 = oNode_job.selectSingleNode("txt_219").Text
    job_txt_220 = oNode_job.selectSingleNode("txt_220").Text

    job_txt_221 = oNode_job.selectSingleNode("txt_221").Text
    job_txt_222 = oNode_job.selectSingleNode("txt_222").Text
    job_txt_223 = oNode_job.selectSingleNode("txt_223").Text
    job_txt_224 = oNode_job.selectSingleNode("txt_224").Text
    job_txt_225 = oNode_job.selectSingleNode("txt_225").Text
    job_txt_226 = oNode_job.selectSingleNode("txt_226").Text
    job_txt_227 = oNode_job.selectSingleNode("txt_227").Text
    job_txt_228 = oNode_job.selectSingleNode("txt_228").Text
    job_txt_229 = oNode_job.selectSingleNode("txt_229").Text
    job_txt_230 = oNode_job.selectSingleNode("txt_230").Text

    job_txt_231 = oNode_job.selectSingleNode("txt_231").Text
    job_txt_232 = oNode_job.selectSingleNode("txt_232").Text
    job_txt_233 = oNode_job.selectSingleNode("txt_233").Text
    job_txt_234 = oNode_job.selectSingleNode("txt_234").Text
    job_txt_235 = oNode_job.selectSingleNode("txt_235").Text
    job_txt_236 = oNode_job.selectSingleNode("txt_236").Text
    job_txt_237 = oNode_job.selectSingleNode("txt_237").Text
    job_txt_238 = oNode_job.selectSingleNode("txt_238").Text
    job_txt_239 = oNode_job.selectSingleNode("txt_239").Text
    job_txt_240 = oNode_job.selectSingleNode("txt_240").Text

    job_txt_241 = oNode_job.selectSingleNode("txt_241").Text
    job_txt_242 = oNode_job.selectSingleNode("txt_242").Text
    job_txt_243 = oNode_job.selectSingleNode("txt_243").Text
    job_txt_244 = oNode_job.selectSingleNode("txt_244").Text
    job_txt_245 = oNode_job.selectSingleNode("txt_245").Text
    job_txt_246 = oNode_job.selectSingleNode("txt_246").Text
    job_txt_247 = oNode_job.selectSingleNode("txt_247").Text
    job_txt_248 = oNode_job.selectSingleNode("txt_248").Text
    job_txt_249 = oNode_job.selectSingleNode("txt_249").Text
    job_txt_250 = oNode_job.selectSingleNode("txt_250").Text

    job_txt_251 = oNode_job.selectSingleNode("txt_251").Text
    job_txt_252 = oNode_job.selectSingleNode("txt_252").Text
    job_txt_253 = oNode_job.selectSingleNode("txt_253").Text
    job_txt_254 = oNode_job.selectSingleNode("txt_254").Text
    job_txt_255 = oNode_job.selectSingleNode("txt_255").Text
    job_txt_256 = oNode_job.selectSingleNode("txt_256").Text
    job_txt_257 = oNode_job.selectSingleNode("txt_257").Text
    job_txt_258 = oNode_job.selectSingleNode("txt_258").Text
    job_txt_259 = oNode_job.selectSingleNode("txt_259").Text
    job_txt_260 = oNode_job.selectSingleNode("txt_260").Text

    job_txt_261 = oNode_job.selectSingleNode("txt_261").Text
    job_txt_262 = oNode_job.selectSingleNode("txt_262").Text
    job_txt_263 = oNode_job.selectSingleNode("txt_263").Text
    job_txt_264 = oNode_job.selectSingleNode("txt_264").Text
    job_txt_265 = oNode_job.selectSingleNode("txt_265").Text
    job_txt_266 = oNode_job.selectSingleNode("txt_266").Text
    job_txt_267 = oNode_job.selectSingleNode("txt_267").Text
    job_txt_268 = oNode_job.selectSingleNode("txt_268").Text
    job_txt_269 = oNode_job.selectSingleNode("txt_269").Text
    job_txt_270 = oNode_job.selectSingleNode("txt_270").Text

    job_txt_271 = oNode_job.selectSingleNode("txt_271").Text
    job_txt_272 = oNode_job.selectSingleNode("txt_272").Text
    job_txt_273 = oNode_job.selectSingleNode("txt_273").Text
    job_txt_274 = oNode_job.selectSingleNode("txt_274").Text
    job_txt_275 = oNode_job.selectSingleNode("txt_275").Text
    job_txt_276 = oNode_job.selectSingleNode("txt_276").Text
    job_txt_277 = oNode_job.selectSingleNode("txt_277").Text
    job_txt_278 = oNode_job.selectSingleNode("txt_278").Text
    job_txt_279 = oNode_job.selectSingleNode("txt_279").Text
    job_txt_280 = oNode_job.selectSingleNode("txt_280").Text

    job_txt_281 = oNode_job.selectSingleNode("txt_281").Text
    job_txt_282 = oNode_job.selectSingleNode("txt_282").Text
    job_txt_283 = oNode_job.selectSingleNode("txt_283").Text
    job_txt_284 = oNode_job.selectSingleNode("txt_284").Text
    job_txt_285 = oNode_job.selectSingleNode("txt_285").Text
    job_txt_286 = oNode_job.selectSingleNode("txt_286").Text
    job_txt_287 = oNode_job.selectSingleNode("txt_287").Text
    job_txt_288 = oNode_job.selectSingleNode("txt_288").Text
    job_txt_289 = oNode_job.selectSingleNode("txt_289").Text
    job_txt_290 = oNode_job.selectSingleNode("txt_290").Text

    job_txt_291 = oNode_job.selectSingleNode("txt_291").Text
    job_txt_292 = oNode_job.selectSingleNode("txt_292").Text
    job_txt_293 = oNode_job.selectSingleNode("txt_293").Text
    job_txt_294 = oNode_job.selectSingleNode("txt_294").Text
    job_txt_295 = oNode_job.selectSingleNode("txt_295").Text
    job_txt_296 = oNode_job.selectSingleNode("txt_296").Text
    job_txt_297 = oNode_job.selectSingleNode("txt_297").Text
    job_txt_298 = oNode_job.selectSingleNode("txt_298").Text
    job_txt_299 = oNode_job.selectSingleNode("txt_299").Text
    job_txt_300 = oNode_job.selectSingleNode("txt_300").Text

    job_txt_301 = oNode_job.selectSingleNode("txt_301").Text
    job_txt_302 = oNode_job.selectSingleNode("txt_302").Text
    job_txt_303 = oNode_job.selectSingleNode("txt_303").Text
    job_txt_304 = oNode_job.selectSingleNode("txt_304").Text
    job_txt_305 = oNode_job.selectSingleNode("txt_305").Text
    job_txt_306 = oNode_job.selectSingleNode("txt_306").Text
    job_txt_307 = oNode_job.selectSingleNode("txt_307").Text
    job_txt_308 = oNode_job.selectSingleNode("txt_308").Text
    job_txt_309 = oNode_job.selectSingleNode("txt_309").Text
    job_txt_310 = oNode_job.selectSingleNode("txt_310").Text

    job_txt_311 = oNode_job.selectSingleNode("txt_311").Text
    job_txt_312 = oNode_job.selectSingleNode("txt_312").Text
    job_txt_313 = oNode_job.selectSingleNode("txt_313").Text
    job_txt_314 = oNode_job.selectSingleNode("txt_314").Text
    job_txt_315 = oNode_job.selectSingleNode("txt_315").Text
    job_txt_316 = oNode_job.selectSingleNode("txt_316").Text
    job_txt_317 = oNode_job.selectSingleNode("txt_317").Text
    job_txt_318 = oNode_job.selectSingleNode("txt_318").Text
    job_txt_319 = oNode_job.selectSingleNode("txt_319").Text
    job_txt_320 = oNode_job.selectSingleNode("txt_320").Text
    job_txt_321 = oNode_job.selectSingleNode("txt_321").Text
    job_txt_322 = oNode_job.selectSingleNode("txt_322").Text
    job_txt_323 = oNode_job.selectSingleNode("txt_323").Text
    job_txt_324 = oNode_job.selectSingleNode("txt_324").Text
    job_txt_325 = oNode_job.selectSingleNode("txt_325").Text
    job_txt_326 = oNode_job.selectSingleNode("txt_326").Text
    job_txt_327 = oNode_job.selectSingleNode("txt_327").Text
    job_txt_328 = oNode_job.selectSingleNode("txt_328").Text
    job_txt_329 = oNode_job.selectSingleNode("txt_329").Text
    job_txt_330 = oNode_job.selectSingleNode("txt_330").Text
    job_txt_331 = oNode_job.selectSingleNode("txt_331").Text
    job_txt_332 = oNode_job.selectSingleNode("txt_332").Text
    job_txt_333 = oNode_job.selectSingleNode("txt_333").Text
    job_txt_334 = oNode_job.selectSingleNode("txt_334").Text
    job_txt_335 = oNode_job.selectSingleNode("txt_335").Text
    job_txt_336 = oNode_job.selectSingleNode("txt_336").Text
    job_txt_337 = oNode_job.selectSingleNode("txt_337").Text
    job_txt_338 = oNode_job.selectSingleNode("txt_338").Text
    job_txt_339 = oNode_job.selectSingleNode("txt_339").Text
    job_txt_340 = oNode_job.selectSingleNode("txt_340").Text

    job_txt_341 = oNode_job.selectSingleNode("txt_341").Text
    job_txt_342 = oNode_job.selectSingleNode("txt_342").Text
    job_txt_343 = oNode_job.selectSingleNode("txt_343").Text
    job_txt_344 = oNode_job.selectSingleNode("txt_344").Text
    job_txt_345 = oNode_job.selectSingleNode("txt_345").Text
    job_txt_346 = oNode_job.selectSingleNode("txt_346").Text
    job_txt_347 = oNode_job.selectSingleNode("txt_347").Text
    job_txt_348 = oNode_job.selectSingleNode("txt_348").Text
    job_txt_349 = oNode_job.selectSingleNode("txt_349").Text
    job_txt_350 = oNode_job.selectSingleNode("txt_350").Text

    job_txt_351 = oNode_job.selectSingleNode("txt_351").Text
    job_txt_352 = oNode_job.selectSingleNode("txt_352").Text
    job_txt_353 = oNode_job.selectSingleNode("txt_353").Text
    job_txt_354 = oNode_job.selectSingleNode("txt_354").Text
    job_txt_355 = oNode_job.selectSingleNode("txt_355").Text
    job_txt_356 = oNode_job.selectSingleNode("txt_356").Text
    job_txt_357 = oNode_job.selectSingleNode("txt_357").Text
    job_txt_358 = oNode_job.selectSingleNode("txt_358").Text
    job_txt_359 = oNode_job.selectSingleNode("txt_359").Text
    job_txt_360 = oNode_job.selectSingleNode("txt_360").Text

    job_txt_361 = oNode_job.selectSingleNode("txt_361").Text
    job_txt_362 = oNode_job.selectSingleNode("txt_362").Text
    job_txt_363 = oNode_job.selectSingleNode("txt_363").Text
    job_txt_364 = oNode_job.selectSingleNode("txt_364").Text
    job_txt_365 = oNode_job.selectSingleNode("txt_365").Text
    job_txt_366 = oNode_job.selectSingleNode("txt_366").Text
    job_txt_367 = oNode_job.selectSingleNode("txt_367").Text
    job_txt_368 = oNode_job.selectSingleNode("txt_368").Text
    job_txt_369 = oNode_job.selectSingleNode("txt_369").Text
    job_txt_370 = oNode_job.selectSingleNode("txt_370").Text

    job_txt_371 = oNode_job.selectSingleNode("txt_371").Text
    job_txt_372 = oNode_job.selectSingleNode("txt_372").Text
    job_txt_373 = oNode_job.selectSingleNode("txt_373").Text
    job_txt_374 = oNode_job.selectSingleNode("txt_374").Text
    job_txt_375 = oNode_job.selectSingleNode("txt_375").Text
    job_txt_376 = oNode_job.selectSingleNode("txt_376").Text
    job_txt_377 = oNode_job.selectSingleNode("txt_377").Text
    job_txt_378 = oNode_job.selectSingleNode("txt_378").Text
    job_txt_379 = oNode_job.selectSingleNode("txt_379").Text
    job_txt_380 = oNode_job.selectSingleNode("txt_380").Text

    job_txt_381 = oNode_job.selectSingleNode("txt_381").Text
    job_txt_382 = oNode_job.selectSingleNode("txt_382").Text
    job_txt_383 = oNode_job.selectSingleNode("txt_383").Text
    job_txt_384 = oNode_job.selectSingleNode("txt_384").Text
    job_txt_385 = oNode_job.selectSingleNode("txt_385").Text
    job_txt_386 = oNode_job.selectSingleNode("txt_386").Text
    job_txt_387 = oNode_job.selectSingleNode("txt_387").Text
    job_txt_388 = oNode_job.selectSingleNode("txt_388").Text
    job_txt_389 = oNode_job.selectSingleNode("txt_389").Text
    job_txt_390 = oNode_job.selectSingleNode("txt_390").Text

    job_txt_391 = oNode_job.selectSingleNode("txt_391").Text
    job_txt_392 = oNode_job.selectSingleNode("txt_392").Text
    job_txt_393 = oNode_job.selectSingleNode("txt_393").Text
    job_txt_394 = oNode_job.selectSingleNode("txt_394").Text
    job_txt_395 = oNode_job.selectSingleNode("txt_395").Text
    job_txt_396 = oNode_job.selectSingleNode("txt_396").Text
    job_txt_397 = oNode_job.selectSingleNode("txt_397").Text
    job_txt_398 = oNode_job.selectSingleNode("txt_398").Text
    job_txt_399 = oNode_job.selectSingleNode("txt_399").Text
    job_txt_400 = oNode_job.selectSingleNode("txt_400").Text

    job_txt_401 = oNode_job.selectSingleNode("txt_401").Text
    job_txt_402 = oNode_job.selectSingleNode("txt_402").Text
    job_txt_403 = oNode_job.selectSingleNode("txt_403").Text
    job_txt_404 = oNode_job.selectSingleNode("txt_404").Text
    job_txt_405 = oNode_job.selectSingleNode("txt_405").Text
    job_txt_406 = oNode_job.selectSingleNode("txt_406").Text
    job_txt_407 = oNode_job.selectSingleNode("txt_407").Text
    job_txt_408 = oNode_job.selectSingleNode("txt_408").Text
    job_txt_409 = oNode_job.selectSingleNode("txt_409").Text
    job_txt_410 = oNode_job.selectSingleNode("txt_410").Text

    job_txt_411 = oNode_job.selectSingleNode("txt_411").Text
    job_txt_412 = oNode_job.selectSingleNode("txt_412").Text
    job_txt_413 = oNode_job.selectSingleNode("txt_413").Text
    job_txt_414 = oNode_job.selectSingleNode("txt_414").Text
    job_txt_415 = oNode_job.selectSingleNode("txt_415").Text
    job_txt_416 = oNode_job.selectSingleNode("txt_416").Text
    job_txt_417 = oNode_job.selectSingleNode("txt_417").Text
    job_txt_418 = oNode_job.selectSingleNode("txt_418").Text
    job_txt_419 = oNode_job.selectSingleNode("txt_419").Text
    job_txt_420 = oNode_job.selectSingleNode("txt_420").Text

    job_txt_421 = oNode_job.selectSingleNode("txt_421").Text
    job_txt_422 = oNode_job.selectSingleNode("txt_422").Text
    job_txt_423 = oNode_job.selectSingleNode("txt_423").Text
    job_txt_424 = oNode_job.selectSingleNode("txt_424").Text
    job_txt_425 = oNode_job.selectSingleNode("txt_425").Text
    job_txt_426 = oNode_job.selectSingleNode("txt_426").Text
    job_txt_427 = oNode_job.selectSingleNode("txt_427").Text
    job_txt_428 = oNode_job.selectSingleNode("txt_428").Text
    job_txt_429 = oNode_job.selectSingleNode("txt_429").Text
    job_txt_430 = oNode_job.selectSingleNode("txt_430").Text

    job_txt_431 = oNode_job.selectSingleNode("txt_431").Text
    job_txt_432 = oNode_job.selectSingleNode("txt_432").Text
    job_txt_433 = oNode_job.selectSingleNode("txt_433").Text
    job_txt_434 = oNode_job.selectSingleNode("txt_434").Text
    job_txt_435 = oNode_job.selectSingleNode("txt_435").Text
    job_txt_436 = oNode_job.selectSingleNode("txt_436").Text
    job_txt_437 = oNode_job.selectSingleNode("txt_437").Text
    job_txt_438 = oNode_job.selectSingleNode("txt_438").Text
    job_txt_439 = oNode_job.selectSingleNode("txt_439").Text
    job_txt_440 = oNode_job.selectSingleNode("txt_440").Text

    job_txt_441 = oNode_job.selectSingleNode("txt_441").Text
    job_txt_442 = oNode_job.selectSingleNode("txt_442").Text
    job_txt_443 = oNode_job.selectSingleNode("txt_443").Text
    job_txt_444 = oNode_job.selectSingleNode("txt_444").Text
    job_txt_445 = oNode_job.selectSingleNode("txt_445").Text
    job_txt_446 = oNode_job.selectSingleNode("txt_446").Text
    job_txt_447 = oNode_job.selectSingleNode("txt_447").Text
    job_txt_448 = oNode_job.selectSingleNode("txt_448").Text
    job_txt_449 = oNode_job.selectSingleNode("txt_449").Text
    job_txt_450 = oNode_job.selectSingleNode("txt_450").Text

    job_txt_451 = oNode_job.selectSingleNode("txt_451").Text
    job_txt_452 = oNode_job.selectSingleNode("txt_452").Text
    job_txt_453 = oNode_job.selectSingleNode("txt_453").Text
    job_txt_454 = oNode_job.selectSingleNode("txt_454").Text
    job_txt_455 = oNode_job.selectSingleNode("txt_455").Text
    job_txt_456 = oNode_job.selectSingleNode("txt_456").Text
    job_txt_457 = oNode_job.selectSingleNode("txt_457").Text
    job_txt_458 = oNode_job.selectSingleNode("txt_458").Text
    job_txt_459 = oNode_job.selectSingleNode("txt_459").Text
    job_txt_460 = oNode_job.selectSingleNode("txt_460").Text

    job_txt_461 = oNode_job.selectSingleNode("txt_461").Text
    job_txt_462 = oNode_job.selectSingleNode("txt_462").Text
    job_txt_463 = oNode_job.selectSingleNode("txt_463").Text
    job_txt_464 = oNode_job.selectSingleNode("txt_464").Text
    job_txt_465 = oNode_job.selectSingleNode("txt_465").Text
    job_txt_466 = oNode_job.selectSingleNode("txt_466").Text
    job_txt_467 = oNode_job.selectSingleNode("txt_467").Text
    job_txt_468 = oNode_job.selectSingleNode("txt_468").Text
    job_txt_469 = oNode_job.selectSingleNode("txt_469").Text
    job_txt_470 = oNode_job.selectSingleNode("txt_470").Text

    job_txt_471 = oNode_job.selectSingleNode("txt_471").Text
    job_txt_472 = oNode_job.selectSingleNode("txt_472").Text
    job_txt_473 = oNode_job.selectSingleNode("txt_473").Text
    job_txt_474 = oNode_job.selectSingleNode("txt_474").Text
    job_txt_475 = oNode_job.selectSingleNode("txt_475").Text
    job_txt_476 = oNode_job.selectSingleNode("txt_476").Text
    job_txt_477 = oNode_job.selectSingleNode("txt_477").Text
    job_txt_478 = oNode_job.selectSingleNode("txt_478").Text
    job_txt_479 = oNode_job.selectSingleNode("txt_479").Text
    job_txt_480 = oNode_job.selectSingleNode("txt_480").Text

    job_txt_481 = oNode_job.selectSingleNode("txt_481").Text
    job_txt_482 = oNode_job.selectSingleNode("txt_482").Text
    job_txt_483 = oNode_job.selectSingleNode("txt_483").Text
    job_txt_484 = oNode_job.selectSingleNode("txt_484").Text
    job_txt_485 = oNode_job.selectSingleNode("txt_485").Text
    job_txt_486 = oNode_job.selectSingleNode("txt_486").Text
    job_txt_487 = oNode_job.selectSingleNode("txt_487").Text
    job_txt_488 = oNode_job.selectSingleNode("txt_488").Text
    job_txt_489 = oNode_job.selectSingleNode("txt_489").Text
    job_txt_490 = oNode_job.selectSingleNode("txt_490").Text

    job_txt_491 = oNode_job.selectSingleNode("txt_491").Text
    job_txt_492 = oNode_job.selectSingleNode("txt_492").Text
    job_txt_493 = oNode_job.selectSingleNode("txt_493").Text
    job_txt_494 = oNode_job.selectSingleNode("txt_494").Text
    job_txt_495 = oNode_job.selectSingleNode("txt_495").Text
    job_txt_496 = oNode_job.selectSingleNode("txt_496").Text
    job_txt_497 = oNode_job.selectSingleNode("txt_497").Text
    job_txt_498 = oNode_job.selectSingleNode("txt_498").Text
    job_txt_499 = oNode_job.selectSingleNode("txt_499").Text
    job_txt_500 = oNode_job.selectSingleNode("txt_500").Text

    job_txt_501 = oNode_job.selectSingleNode("txt_501").Text
    job_txt_502 = oNode_job.selectSingleNode("txt_502").Text
    job_txt_503 = oNode_job.selectSingleNode("txt_503").Text
    job_txt_504 = oNode_job.selectSingleNode("txt_504").Text
    job_txt_505 = oNode_job.selectSingleNode("txt_505").Text
    job_txt_506 = oNode_job.selectSingleNode("txt_506").Text
    job_txt_507 = oNode_job.selectSingleNode("txt_507").Text
    job_txt_508 = oNode_job.selectSingleNode("txt_508").Text
    job_txt_509 = oNode_job.selectSingleNode("txt_509").Text
    job_txt_510 = oNode_job.selectSingleNode("txt_510").Text

    job_txt_501 = oNode_job.selectSingleNode("txt_501").Text
    job_txt_502 = oNode_job.selectSingleNode("txt_502").Text
    job_txt_503 = oNode_job.selectSingleNode("txt_503").Text
    job_txt_504 = oNode_job.selectSingleNode("txt_504").Text
    job_txt_505 = oNode_job.selectSingleNode("txt_505").Text
    job_txt_506 = oNode_job.selectSingleNode("txt_506").Text
    job_txt_507 = oNode_job.selectSingleNode("txt_507").Text
    job_txt_508 = oNode_job.selectSingleNode("txt_508").Text
    job_txt_509 = oNode_job.selectSingleNode("txt_509").Text
    job_txt_510 = oNode_job.selectSingleNode("txt_510").Text

    job_txt_511 = oNode_job.selectSingleNode("txt_511").Text
    job_txt_512 = oNode_job.selectSingleNode("txt_512").Text
    job_txt_513 = oNode_job.selectSingleNode("txt_513").Text
    job_txt_514 = oNode_job.selectSingleNode("txt_514").Text
    job_txt_515 = oNode_job.selectSingleNode("txt_515").Text
    job_txt_516 = oNode_job.selectSingleNode("txt_516").Text
    job_txt_517 = oNode_job.selectSingleNode("txt_517").Text
    job_txt_518 = oNode_job.selectSingleNode("txt_518").Text
    job_txt_519 = oNode_job.selectSingleNode("txt_519").Text
    job_txt_520 = oNode_job.selectSingleNode("txt_520").Text

    job_txt_521 = oNode_job.selectSingleNode("txt_521").Text
    job_txt_522 = oNode_job.selectSingleNode("txt_522").Text
    job_txt_523 = oNode_job.selectSingleNode("txt_523").Text
    job_txt_524 = oNode_job.selectSingleNode("txt_524").Text
    job_txt_525 = oNode_job.selectSingleNode("txt_525").Text
    job_txt_526 = oNode_job.selectSingleNode("txt_526").Text
    job_txt_527 = oNode_job.selectSingleNode("txt_527").Text
    job_txt_528 = oNode_job.selectSingleNode("txt_528").Text
    job_txt_529 = oNode_job.selectSingleNode("txt_539").Text
    job_txt_530 = oNode_job.selectSingleNode("txt_540").Text

    job_txt_531 = oNode_job.selectSingleNode("txt_531").Text
    job_txt_532 = oNode_job.selectSingleNode("txt_532").Text
    job_txt_533 = oNode_job.selectSingleNode("txt_533").Text
    job_txt_534 = oNode_job.selectSingleNode("txt_534").Text
    job_txt_535 = oNode_job.selectSingleNode("txt_535").Text
    job_txt_536 = oNode_job.selectSingleNode("txt_536").Text
    job_txt_537 = oNode_job.selectSingleNode("txt_537").Text
    job_txt_538 = oNode_job.selectSingleNode("txt_538").Text
    job_txt_539 = oNode_job.selectSingleNode("txt_539").Text
    job_txt_540 = oNode_job.selectSingleNode("txt_540").Text

    job_txt_541 = oNode_job.selectSingleNode("txt_541").Text
    job_txt_542 = oNode_job.selectSingleNode("txt_542").Text
    job_txt_543 = oNode_job.selectSingleNode("txt_543").Text
    job_txt_544 = oNode_job.selectSingleNode("txt_544").Text
    job_txt_545 = oNode_job.selectSingleNode("txt_545").Text
    job_txt_546 = oNode_job.selectSingleNode("txt_546").Text
    job_txt_547 = oNode_job.selectSingleNode("txt_547").Text
    job_txt_548 = oNode_job.selectSingleNode("txt_548").Text
    job_txt_549 = oNode_job.selectSingleNode("txt_549").Text
    job_txt_550 = oNode_job.selectSingleNode("txt_550").Text

    job_txt_551 = oNode_job.selectSingleNode("txt_551").Text
    job_txt_552 = oNode_job.selectSingleNode("txt_552").Text
    job_txt_553 = oNode_job.selectSingleNode("txt_553").Text
    job_txt_554 = oNode_job.selectSingleNode("txt_554").Text
    job_txt_555 = oNode_job.selectSingleNode("txt_555").Text
    job_txt_556 = oNode_job.selectSingleNode("txt_556").Text
    job_txt_557 = oNode_job.selectSingleNode("txt_557").Text
    job_txt_558 = oNode_job.selectSingleNode("txt_558").Text
    job_txt_559 = oNode_job.selectSingleNode("txt_559").Text
    job_txt_560 = oNode_job.selectSingleNode("txt_560").Text
    job_txt_561 = oNode_job.selectSingleNode("txt_561").Text
    job_txt_562 = oNode_job.selectSingleNode("txt_562").Text
    job_txt_563 = oNode_job.selectSingleNode("txt_563").Text
    job_txt_564 = oNode_job.selectSingleNode("txt_564").Text
    job_txt_565 = oNode_job.selectSingleNode("txt_565").Text
    job_txt_566 = oNode_job.selectSingleNode("txt_566").Text
    job_txt_567 = oNode_job.selectSingleNode("txt_567").Text
    job_txt_568 = oNode_job.selectSingleNode("txt_568").Text
    job_txt_569 = oNode_job.selectSingleNode("txt_569").Text
    job_txt_570 = oNode_job.selectSingleNode("txt_570").Text
    job_txt_571 = oNode_job.selectSingleNode("txt_571").Text
    job_txt_572 = oNode_job.selectSingleNode("txt_572").Text
    job_txt_573 = oNode_job.selectSingleNode("txt_573").Text
    job_txt_574 = oNode_job.selectSingleNode("txt_574").Text
    job_txt_575 = oNode_job.selectSingleNode("txt_575").Text
    job_txt_576 = oNode_job.selectSingleNode("txt_576").Text
    job_txt_577 = oNode_job.selectSingleNode("txt_577").Text
    job_txt_578 = oNode_job.selectSingleNode("txt_578").Text
    job_txt_579 = oNode_job.selectSingleNode("txt_579").Text
    job_txt_580 = oNode_job.selectSingleNode("txt_580").Text
    job_txt_581 = oNode_job.selectSingleNode("txt_581").Text
    job_txt_582 = oNode_job.selectSingleNode("txt_582").Text
    job_txt_583 = oNode_job.selectSingleNode("txt_583").Text
    job_txt_584 = oNode_job.selectSingleNode("txt_584").Text
    job_txt_585 = oNode_job.selectSingleNode("txt_585").Text
    job_txt_586 = oNode_job.selectSingleNode("txt_586").Text
    job_txt_587 = oNode_job.selectSingleNode("txt_587").Text

    job_txt_588 = oNode_job.selectSingleNode("txt_588").Text
    job_txt_589 = oNode_job.selectSingleNode("txt_589").Text
    job_txt_590 = oNode_job.selectSingleNode("txt_590").Text
    job_txt_591 = oNode_job.selectSingleNode("txt_591").Text
    job_txt_592 = oNode_job.selectSingleNode("txt_592").Text
    job_txt_593 = oNode_job.selectSingleNode("txt_593").Text
    job_txt_594 = oNode_job.selectSingleNode("txt_594").Text
    job_txt_595 = oNode_job.selectSingleNode("txt_595").Text
    job_txt_596 = oNode_job.selectSingleNode("txt_596").Text
    job_txt_597 = oNode_job.selectSingleNode("txt_597").Text
    job_txt_598 = oNode_job.selectSingleNode("txt_598").Text
    job_txt_599 = oNode_job.selectSingleNode("txt_599").Text

    job_txt_600 = oNode_job.selectSingleNode("txt_600").Text
    job_txt_601 = oNode_job.selectSingleNode("txt_601").Text
    job_txt_602 = oNode_job.selectSingleNode("txt_602").Text
    job_txt_603 = oNode_job.selectSingleNode("txt_603").Text
    job_txt_604 = oNode_job.selectSingleNode("txt_604").Text
    job_txt_605 = oNode_job.selectSingleNode("txt_605").Text
    job_txt_606 = oNode_job.selectSingleNode("txt_606").Text
    job_txt_607 = oNode_job.selectSingleNode("txt_607").Text
    job_txt_608 = oNode_job.selectSingleNode("txt_608").Text
    job_txt_609 = oNode_job.selectSingleNode("txt_609").Text
    job_txt_610 = oNode_job.selectSingleNode("txt_610").Text
    job_txt_611 = oNode_job.selectSingleNode("txt_611").Text
    job_txt_612 = oNode_job.selectSingleNode("txt_612").Text
    job_txt_613 = oNode_job.selectSingleNode("txt_613").Text
    job_txt_614 = oNode_job.selectSingleNode("txt_614").Text
    job_txt_615 = oNode_job.selectSingleNode("txt_615").Text
    job_txt_616 = oNode_job.selectSingleNode("txt_616").Text
    job_txt_617 = oNode_job.selectSingleNode("txt_617").Text
    job_txt_618 = oNode_job.selectSingleNode("txt_618").Text
    job_txt_619 = oNode_job.selectSingleNode("txt_619").Text
    job_txt_620 = oNode_job.selectSingleNode("txt_620").Text
    job_txt_621 = oNode_job.selectSingleNode("txt_621").Text
    job_txt_622 = oNode_job.selectSingleNode("txt_622").Text
    job_txt_623 = oNode_job.selectSingleNode("txt_623").Text
    job_txt_624 = oNode_job.selectSingleNode("txt_624").Text
    job_txt_625 = oNode_job.selectSingleNode("txt_625").Text

          
    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>