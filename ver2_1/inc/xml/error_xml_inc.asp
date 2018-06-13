
<%




Dim objXMLHTTP_error, objXMLDOM_error, i_error, strHTML_error
Dim Address_error, Latitude_error, Longitude_error
Dim oNode_error, oNodes_error
Dim sXPathQuery_error



Set objXMLDom_error = Server.CreateObject("Microsoft.XMLDOM")
Set objXmlHttp_error = Server.CreateObject("Msxml2.ServerXMLHTTP")
'objXmlHttp_error.open "GET", "http://localhost/Git/timeout_2016/ver2_1/inc/xml/error_sprog.xml", False
'objXmlHttp_error.open "GET", "http://localhost/inc/xml/error_sprog.xml", False
'objXmlHttp_error.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/xml/error_sprog.xml", False
'objXmlHttp_error.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver4_22/inc/xml/error_sprog.xml", False
'objXmlHttp_error.open "GET", "https://timeout.cloud/timeout_xp/wwwroot/ver3_99/inc/xml/error_sprog.xml", False
objXmlHttp_error.open "GET", "https://outzource.dk/timeout_xp/wwwroot/ver2_14/inc/xml/error_sprog.xml", False
objXmlHttp_error.send


Set objXmlDom_error = objXmlHttp_error.responseXML

'Henter indholdet af alle tags med navnet 'titel'
'Set objModuler = objXMLDOM.getElementsByTagName("se")


Set objXmlHttp_error = Nothing





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
sXPathQuery_error = "//sprog/dk"
'Session.LCID = 1030
case 2
sXPathQuery_error = "//sprog/uk"
'Session.LCID = 1033
case 3
sXPathQuery_error = "//sprog/se"
'Session.LCID = 1053
case 4
sXPathQuery_error = "//sprog/no"
'Session.LCID = 2068
case 5
sXPathQuery_error = "//sprog/es"
'Session.LCID = 1034
case 6
sXPathQuery_error = "//sprog/de"
'Session.LCID = 1031
case 7
sXPathQuery_error = "//sprog/fr"
'Session.LCID = 1036
case else
sXPathQuery_error = "//sprog/dk"
'Session.LCID = 1030
end select

'*** ALTID DK ellers er der fejl i alle beløb og valtuaer omregninger hvis der er punktum i tallet.
Session.LCID = 1030


Set oNode_error = objXmlDom_error.documentElement.selectSingleNode(sXPathQuery_error)
Address_error = oNode_error.Text

Set oNodes_error = objXmlDom_error.documentElement.selectNodes(sXPathQuery_error)

    For Each oNode_error in oNodes_error
          
          err_txt_001 = oNode_error.selectSingleNode("txt_001").Text
          err_txt_004 = oNode_error.selectSingleNode("txt_004").Text
          err_txt_028 = oNode_error.selectSingleNode("txt_028").Text
          
          err_txt_074 = oNode_error.selectSingleNode("txt_074").Text
          
          err_txt_090 = oNode_error.selectSingleNode("txt_090").Text
          err_txt_106 = oNode_error.selectSingleNode("txt_106").Text
          err_txt_107 = oNode_error.selectSingleNode("txt_107").Text
          
          err_txt_108A = oNode_error.selectSingleNode("txt_108A").Text
          err_txt_108B = oNode_error.selectSingleNode("txt_108B").Text
          err_txt_108C = oNode_error.selectSingleNode("txt_108C").Text
          
          err_txt_120 = oNode_error.selectSingleNode("txt_120").Text

    err_txt_121 = oNode_error.selectSingleNode("txt_121").Text
    err_txt_122 = oNode_error.selectSingleNode("txt_122").Text
    err_txt_123 = oNode_error.selectSingleNode("txt_123").Text
    err_txt_124 = oNode_error.selectSingleNode("txt_124").Text
    err_txt_125 = oNode_error.selectSingleNode("txt_125").Text
    err_txt_126 = oNode_error.selectSingleNode("txt_126").Text
    err_txt_127 = oNode_error.selectSingleNode("txt_127").Text
    err_txt_128 = oNode_error.selectSingleNode("txt_128").Text
    err_txt_129 = oNode_error.selectSingleNode("txt_129").Text
    err_txt_130 = oNode_error.selectSingleNode("txt_130").Text

    err_txt_131 = oNode_error.selectSingleNode("txt_131").Text
    err_txt_132 = oNode_error.selectSingleNode("txt_132").Text
    err_txt_133 = oNode_error.selectSingleNode("txt_133").Text
    err_txt_134 = oNode_error.selectSingleNode("txt_134").Text
    err_txt_135 = oNode_error.selectSingleNode("txt_135").Text
    err_txt_136 = oNode_error.selectSingleNode("txt_136").Text
    err_txt_137 = oNode_error.selectSingleNode("txt_137").Text
    err_txt_138 = oNode_error.selectSingleNode("txt_138").Text
    err_txt_139 = oNode_error.selectSingleNode("txt_139").Text
    err_txt_140 = oNode_error.selectSingleNode("txt_140").Text

    err_txt_141 = oNode_error.selectSingleNode("txt_141").Text
    err_txt_142 = oNode_error.selectSingleNode("txt_142").Text
    err_txt_143 = oNode_error.selectSingleNode("txt_143").Text
    err_txt_144 = oNode_error.selectSingleNode("txt_144").Text
    err_txt_145 = oNode_error.selectSingleNode("txt_145").Text
    err_txt_146 = oNode_error.selectSingleNode("txt_146").Text
    err_txt_147 = oNode_error.selectSingleNode("txt_147").Text
    err_txt_148 = oNode_error.selectSingleNode("txt_148").Text
    err_txt_149 = oNode_error.selectSingleNode("txt_149").Text
    err_txt_150 = oNode_error.selectSingleNode("txt_150").Text

    err_txt_151 = oNode_error.selectSingleNode("txt_151").Text
    err_txt_152 = oNode_error.selectSingleNode("txt_152").Text
    err_txt_153 = oNode_error.selectSingleNode("txt_153").Text
    err_txt_154 = oNode_error.selectSingleNode("txt_154").Text
    err_txt_155 = oNode_error.selectSingleNode("txt_155").Text
    err_txt_156 = oNode_error.selectSingleNode("txt_156").Text
    err_txt_157 = oNode_error.selectSingleNode("txt_157").Text
    err_txt_158 = oNode_error.selectSingleNode("txt_158").Text
    err_txt_159 = oNode_error.selectSingleNode("txt_159").Text
    err_txt_160 = oNode_error.selectSingleNode("txt_160").Text

    err_txt_161 = oNode_error.selectSingleNode("txt_161").Text
    err_txt_162 = oNode_error.selectSingleNode("txt_162").Text
    err_txt_163 = oNode_error.selectSingleNode("txt_163").Text
    err_txt_164 = oNode_error.selectSingleNode("txt_164").Text
    err_txt_165 = oNode_error.selectSingleNode("txt_165").Text
    err_txt_166 = oNode_error.selectSingleNode("txt_166").Text
    err_txt_167 = oNode_error.selectSingleNode("txt_167").Text
    err_txt_168 = oNode_error.selectSingleNode("txt_168").Text
    err_txt_169 = oNode_error.selectSingleNode("txt_169").Text
    err_txt_170 = oNode_error.selectSingleNode("txt_170").Text

    err_txt_171 = oNode_error.selectSingleNode("txt_171").Text
    err_txt_172 = oNode_error.selectSingleNode("txt_172").Text
    err_txt_173 = oNode_error.selectSingleNode("txt_173").Text
    err_txt_174 = oNode_error.selectSingleNode("txt_174").Text
    err_txt_175 = oNode_error.selectSingleNode("txt_175").Text
    err_txt_176 = oNode_error.selectSingleNode("txt_176").Text
    err_txt_177 = oNode_error.selectSingleNode("txt_177").Text
    err_txt_178 = oNode_error.selectSingleNode("txt_178").Text
    err_txt_179 = oNode_error.selectSingleNode("txt_179").Text
    err_txt_180 = oNode_error.selectSingleNode("txt_180").Text

    err_txt_181 = oNode_error.selectSingleNode("txt_181").Text
    err_txt_182 = oNode_error.selectSingleNode("txt_182").Text
    err_txt_183 = oNode_error.selectSingleNode("txt_183").Text
    err_txt_184 = oNode_error.selectSingleNode("txt_184").Text
    err_txt_185 = oNode_error.selectSingleNode("txt_185").Text
    err_txt_186 = oNode_error.selectSingleNode("txt_186").Text
    err_txt_187 = oNode_error.selectSingleNode("txt_187").Text
    err_txt_188 = oNode_error.selectSingleNode("txt_188").Text
    err_txt_189 = oNode_error.selectSingleNode("txt_189").Text
    err_txt_190 = oNode_error.selectSingleNode("txt_190").Text

    err_txt_191 = oNode_error.selectSingleNode("txt_191").Text
    err_txt_192 = oNode_error.selectSingleNode("txt_192").Text
    err_txt_193 = oNode_error.selectSingleNode("txt_193").Text
    err_txt_194 = oNode_error.selectSingleNode("txt_194").Text
    err_txt_195 = oNode_error.selectSingleNode("txt_195").Text
    err_txt_196 = oNode_error.selectSingleNode("txt_196").Text
    err_txt_197 = oNode_error.selectSingleNode("txt_197").Text
    err_txt_198 = oNode_error.selectSingleNode("txt_198").Text
    err_txt_199 = oNode_error.selectSingleNode("txt_199").Text
    err_txt_200 = oNode_error.selectSingleNode("txt_200").Text

    err_txt_201 = oNode_error.selectSingleNode("txt_201").Text
    err_txt_202 = oNode_error.selectSingleNode("txt_202").Text
    err_txt_203 = oNode_error.selectSingleNode("txt_203").Text
    err_txt_204 = oNode_error.selectSingleNode("txt_204").Text
    err_txt_205 = oNode_error.selectSingleNode("txt_205").Text
    err_txt_206 = oNode_error.selectSingleNode("txt_206").Text
    err_txt_207 = oNode_error.selectSingleNode("txt_207").Text
    err_txt_208 = oNode_error.selectSingleNode("txt_208").Text
    err_txt_209 = oNode_error.selectSingleNode("txt_209").Text
    err_txt_210 = oNode_error.selectSingleNode("txt_210").Text

    err_txt_211 = oNode_error.selectSingleNode("txt_211").Text
    err_txt_212 = oNode_error.selectSingleNode("txt_212").Text
    err_txt_213 = oNode_error.selectSingleNode("txt_213").Text
    err_txt_214 = oNode_error.selectSingleNode("txt_214").Text
    err_txt_215 = oNode_error.selectSingleNode("txt_215").Text
    err_txt_216 = oNode_error.selectSingleNode("txt_216").Text
    err_txt_217 = oNode_error.selectSingleNode("txt_217").Text
    err_txt_218 = oNode_error.selectSingleNode("txt_218").Text
    err_txt_219 = oNode_error.selectSingleNode("txt_219").Text
    err_txt_220 = oNode_error.selectSingleNode("txt_220").Text

    err_txt_221 = oNode_error.selectSingleNode("txt_221").Text
    err_txt_222 = oNode_error.selectSingleNode("txt_222").Text
    err_txt_223 = oNode_error.selectSingleNode("txt_223").Text
    err_txt_224 = oNode_error.selectSingleNode("txt_224").Text
    err_txt_225 = oNode_error.selectSingleNode("txt_225").Text
    err_txt_226 = oNode_error.selectSingleNode("txt_226").Text
    err_txt_227 = oNode_error.selectSingleNode("txt_227").Text
    err_txt_228 = oNode_error.selectSingleNode("txt_228").Text
    err_txt_229 = oNode_error.selectSingleNode("txt_229").Text
    err_txt_230 = oNode_error.selectSingleNode("txt_230").Text

    err_txt_231 = oNode_error.selectSingleNode("txt_231").Text
    err_txt_232 = oNode_error.selectSingleNode("txt_232").Text
    err_txt_233 = oNode_error.selectSingleNode("txt_233").Text
    err_txt_234 = oNode_error.selectSingleNode("txt_234").Text
    err_txt_235 = oNode_error.selectSingleNode("txt_235").Text
    err_txt_236 = oNode_error.selectSingleNode("txt_236").Text
    err_txt_237 = oNode_error.selectSingleNode("txt_237").Text
    err_txt_238 = oNode_error.selectSingleNode("txt_238").Text
    err_txt_239 = oNode_error.selectSingleNode("txt_239").Text
    err_txt_240 = oNode_error.selectSingleNode("txt_240").Text

    err_txt_241 = oNode_error.selectSingleNode("txt_241").Text
    err_txt_242 = oNode_error.selectSingleNode("txt_242").Text
    err_txt_243 = oNode_error.selectSingleNode("txt_243").Text
    err_txt_244 = oNode_error.selectSingleNode("txt_244").Text
    err_txt_245 = oNode_error.selectSingleNode("txt_245").Text
    err_txt_246 = oNode_error.selectSingleNode("txt_246").Text
    err_txt_247 = oNode_error.selectSingleNode("txt_247").Text
    err_txt_248 = oNode_error.selectSingleNode("txt_248").Text
    err_txt_249 = oNode_error.selectSingleNode("txt_249").Text
    err_txt_250 = oNode_error.selectSingleNode("txt_250").Text

    err_txt_251 = oNode_error.selectSingleNode("txt_251").Text
    err_txt_252 = oNode_error.selectSingleNode("txt_252").Text
    err_txt_253 = oNode_error.selectSingleNode("txt_253").Text
    err_txt_254 = oNode_error.selectSingleNode("txt_254").Text
    err_txt_255 = oNode_error.selectSingleNode("txt_255").Text
    err_txt_256 = oNode_error.selectSingleNode("txt_256").Text
    err_txt_257 = oNode_error.selectSingleNode("txt_257").Text
    err_txt_258 = oNode_error.selectSingleNode("txt_258").Text
    err_txt_259 = oNode_error.selectSingleNode("txt_259").Text
    err_txt_260 = oNode_error.selectSingleNode("txt_260").Text

    err_txt_261 = oNode_error.selectSingleNode("txt_261").Text
    err_txt_262 = oNode_error.selectSingleNode("txt_262").Text
    err_txt_263 = oNode_error.selectSingleNode("txt_263").Text
    err_txt_264 = oNode_error.selectSingleNode("txt_264").Text
    err_txt_265 = oNode_error.selectSingleNode("txt_265").Text
    err_txt_266 = oNode_error.selectSingleNode("txt_266").Text
    err_txt_267 = oNode_error.selectSingleNode("txt_267").Text
    err_txt_268 = oNode_error.selectSingleNode("txt_268").Text
    err_txt_269 = oNode_error.selectSingleNode("txt_269").Text
    err_txt_270 = oNode_error.selectSingleNode("txt_270").Text

    err_txt_271 = oNode_error.selectSingleNode("txt_271").Text
    err_txt_272 = oNode_error.selectSingleNode("txt_272").Text
    err_txt_273 = oNode_error.selectSingleNode("txt_273").Text
    err_txt_274 = oNode_error.selectSingleNode("txt_274").Text
    err_txt_275 = oNode_error.selectSingleNode("txt_275").Text
    err_txt_276 = oNode_error.selectSingleNode("txt_276").Text
    err_txt_277 = oNode_error.selectSingleNode("txt_277").Text
    err_txt_278 = oNode_error.selectSingleNode("txt_278").Text
    err_txt_279 = oNode_error.selectSingleNode("txt_279").Text
    err_txt_280 = oNode_error.selectSingleNode("txt_280").Text

    err_txt_281 = oNode_error.selectSingleNode("txt_281").Text
    err_txt_282 = oNode_error.selectSingleNode("txt_282").Text
    err_txt_283 = oNode_error.selectSingleNode("txt_283").Text
    err_txt_284 = oNode_error.selectSingleNode("txt_284").Text
    err_txt_285 = oNode_error.selectSingleNode("txt_285").Text
    err_txt_286 = oNode_error.selectSingleNode("txt_286").Text
    err_txt_287 = oNode_error.selectSingleNode("txt_287").Text
    err_txt_288 = oNode_error.selectSingleNode("txt_288").Text
    err_txt_289 = oNode_error.selectSingleNode("txt_289").Text
    err_txt_290 = oNode_error.selectSingleNode("txt_290").Text

    err_txt_291 = oNode_error.selectSingleNode("txt_291").Text
    err_txt_292 = oNode_error.selectSingleNode("txt_292").Text
    err_txt_293 = oNode_error.selectSingleNode("txt_293").Text
    err_txt_294 = oNode_error.selectSingleNode("txt_294").Text
    err_txt_295 = oNode_error.selectSingleNode("txt_295").Text
    err_txt_296 = oNode_error.selectSingleNode("txt_296").Text
    err_txt_297 = oNode_error.selectSingleNode("txt_297").Text
    err_txt_298 = oNode_error.selectSingleNode("txt_298").Text
    err_txt_299 = oNode_error.selectSingleNode("txt_299").Text
    err_txt_300 = oNode_error.selectSingleNode("txt_300").Text

    err_txt_301 = oNode_error.selectSingleNode("txt_301").Text
    err_txt_302 = oNode_error.selectSingleNode("txt_302").Text
    err_txt_303 = oNode_error.selectSingleNode("txt_303").Text
    err_txt_304 = oNode_error.selectSingleNode("txt_304").Text
    err_txt_305 = oNode_error.selectSingleNode("txt_305").Text
    err_txt_306 = oNode_error.selectSingleNode("txt_306").Text
    err_txt_307 = oNode_error.selectSingleNode("txt_307").Text
    err_txt_308 = oNode_error.selectSingleNode("txt_308").Text
    err_txt_309 = oNode_error.selectSingleNode("txt_309").Text
    err_txt_310 = oNode_error.selectSingleNode("txt_310").Text

    err_txt_311 = oNode_error.selectSingleNode("txt_311").Text
    err_txt_312 = oNode_error.selectSingleNode("txt_312").Text
    err_txt_313 = oNode_error.selectSingleNode("txt_313").Text
    err_txt_314 = oNode_error.selectSingleNode("txt_314").Text
    err_txt_315 = oNode_error.selectSingleNode("txt_315").Text
    err_txt_316 = oNode_error.selectSingleNode("txt_316").Text
    err_txt_317 = oNode_error.selectSingleNode("txt_317").Text
    err_txt_318 = oNode_error.selectSingleNode("txt_318").Text
    err_txt_319 = oNode_error.selectSingleNode("txt_319").Text
    err_txt_320 = oNode_error.selectSingleNode("txt_320").Text

    err_txt_321 = oNode_error.selectSingleNode("txt_321").Text
    err_txt_322 = oNode_error.selectSingleNode("txt_322").Text
    err_txt_323 = oNode_error.selectSingleNode("txt_323").Text
    err_txt_324 = oNode_error.selectSingleNode("txt_324").Text
    err_txt_325 = oNode_error.selectSingleNode("txt_325").Text
    err_txt_326 = oNode_error.selectSingleNode("txt_326").Text
    err_txt_327 = oNode_error.selectSingleNode("txt_327").Text
    err_txt_328 = oNode_error.selectSingleNode("txt_328").Text
    err_txt_329 = oNode_error.selectSingleNode("txt_329").Text
    err_txt_330 = oNode_error.selectSingleNode("txt_330").Text

    err_txt_331 = oNode_error.selectSingleNode("txt_331").Text
    err_txt_332 = oNode_error.selectSingleNode("txt_332").Text
    err_txt_333 = oNode_error.selectSingleNode("txt_333").Text
    err_txt_334 = oNode_error.selectSingleNode("txt_334").Text
    err_txt_335 = oNode_error.selectSingleNode("txt_335").Text
    err_txt_336 = oNode_error.selectSingleNode("txt_336").Text
    err_txt_337 = oNode_error.selectSingleNode("txt_337").Text
    err_txt_338 = oNode_error.selectSingleNode("txt_338").Text
    err_txt_339 = oNode_error.selectSingleNode("txt_339").Text
    err_txt_340 = oNode_error.selectSingleNode("txt_340").Text

    err_txt_341 = oNode_error.selectSingleNode("txt_341").Text
    err_txt_342 = oNode_error.selectSingleNode("txt_342").Text
    err_txt_343 = oNode_error.selectSingleNode("txt_343").Text
    err_txt_344 = oNode_error.selectSingleNode("txt_344").Text
    err_txt_345 = oNode_error.selectSingleNode("txt_345").Text
    err_txt_346 = oNode_error.selectSingleNode("txt_346").Text
    err_txt_347 = oNode_error.selectSingleNode("txt_347").Text
    err_txt_348 = oNode_error.selectSingleNode("txt_348").Text
    err_txt_349 = oNode_error.selectSingleNode("txt_349").Text
    err_txt_350 = oNode_error.selectSingleNode("txt_350").Text

    err_txt_351 = oNode_error.selectSingleNode("txt_351").Text
    err_txt_352 = oNode_error.selectSingleNode("txt_352").Text
    err_txt_353 = oNode_error.selectSingleNode("txt_353").Text
    err_txt_354 = oNode_error.selectSingleNode("txt_354").Text
    err_txt_355 = oNode_error.selectSingleNode("txt_355").Text
    err_txt_356 = oNode_error.selectSingleNode("txt_365").Text
    err_txt_357 = oNode_error.selectSingleNode("txt_357").Text
    err_txt_358 = oNode_error.selectSingleNode("txt_358").Text
    err_txt_359 = oNode_error.selectSingleNode("txt_359").Text
    err_txt_360 = oNode_error.selectSingleNode("txt_360").Text

    err_txt_361 = oNode_error.selectSingleNode("txt_361").Text
    err_txt_362 = oNode_error.selectSingleNode("txt_362").Text
    err_txt_363 = oNode_error.selectSingleNode("txt_363").Text
    err_txt_364 = oNode_error.selectSingleNode("txt_364").Text
    err_txt_365 = oNode_error.selectSingleNode("txt_365").Text
    err_txt_366 = oNode_error.selectSingleNode("txt_366").Text
    err_txt_367 = oNode_error.selectSingleNode("txt_367").Text
    err_txt_368 = oNode_error.selectSingleNode("txt_368").Text
    err_txt_369 = oNode_error.selectSingleNode("txt_369").Text
    err_txt_370 = oNode_error.selectSingleNode("txt_370").Text

    err_txt_371 = oNode_error.selectSingleNode("txt_371").Text
    err_txt_372 = oNode_error.selectSingleNode("txt_372").Text
    err_txt_373 = oNode_error.selectSingleNode("txt_373").Text
    err_txt_374 = oNode_error.selectSingleNode("txt_374").Text
    err_txt_375 = oNode_error.selectSingleNode("txt_375").Text
    err_txt_376 = oNode_error.selectSingleNode("txt_376").Text
    err_txt_377 = oNode_error.selectSingleNode("txt_377").Text
    err_txt_378 = oNode_error.selectSingleNode("txt_378").Text
    err_txt_379 = oNode_error.selectSingleNode("txt_379").Text
    err_txt_380 = oNode_error.selectSingleNode("txt_380").Text

    err_txt_381 = oNode_error.selectSingleNode("txt_381").Text
    err_txt_382 = oNode_error.selectSingleNode("txt_382").Text
    err_txt_383 = oNode_error.selectSingleNode("txt_383").Text
    err_txt_384 = oNode_error.selectSingleNode("txt_384").Text
    err_txt_385 = oNode_error.selectSingleNode("txt_385").Text
    err_txt_386 = oNode_error.selectSingleNode("txt_386").Text
    err_txt_387 = oNode_error.selectSingleNode("txt_387").Text
    err_txt_388 = oNode_error.selectSingleNode("txt_388").Text
    err_txt_389 = oNode_error.selectSingleNode("txt_389").Text
    err_txt_390 = oNode_error.selectSingleNode("txt_390").Text

    err_txt_391 = oNode_error.selectSingleNode("txt_391").Text
    err_txt_392 = oNode_error.selectSingleNode("txt_392").Text
    err_txt_393 = oNode_error.selectSingleNode("txt_393").Text
    err_txt_394 = oNode_error.selectSingleNode("txt_394").Text
    err_txt_395 = oNode_error.selectSingleNode("txt_395").Text
    err_txt_396 = oNode_error.selectSingleNode("txt_396").Text
    err_txt_397 = oNode_error.selectSingleNode("txt_397").Text
    err_txt_398 = oNode_error.selectSingleNode("txt_398").Text
    err_txt_399 = oNode_error.selectSingleNode("txt_399").Text
    err_txt_340 = oNode_error.selectSingleNode("txt_340").Text

    err_txt_401 = oNode_error.selectSingleNode("txt_401").Text
    err_txt_402 = oNode_error.selectSingleNode("txt_402").Text
    err_txt_403 = oNode_error.selectSingleNode("txt_403").Text
    err_txt_404 = oNode_error.selectSingleNode("txt_404").Text
    err_txt_405 = oNode_error.selectSingleNode("txt_405").Text
    err_txt_406 = oNode_error.selectSingleNode("txt_406").Text
    err_txt_407 = oNode_error.selectSingleNode("txt_407").Text
    err_txt_408 = oNode_error.selectSingleNode("txt_408").Text
    err_txt_409 = oNode_error.selectSingleNode("txt_409").Text
    err_txt_410 = oNode_error.selectSingleNode("txt_410").Text

    err_txt_411 = oNode_error.selectSingleNode("txt_411").Text
    err_txt_412 = oNode_error.selectSingleNode("txt_412").Text
    err_txt_413 = oNode_error.selectSingleNode("txt_413").Text
    err_txt_414 = oNode_error.selectSingleNode("txt_414").Text
    err_txt_415 = oNode_error.selectSingleNode("txt_415").Text
    err_txt_416 = oNode_error.selectSingleNode("txt_416").Text
    err_txt_417 = oNode_error.selectSingleNode("txt_417").Text
    err_txt_418 = oNode_error.selectSingleNode("txt_418").Text
    err_txt_419 = oNode_error.selectSingleNode("txt_419").Text
    err_txt_420 = oNode_error.selectSingleNode("txt_420").Text

    err_txt_421 = oNode_error.selectSingleNode("txt_421").Text
    err_txt_422 = oNode_error.selectSingleNode("txt_422").Text
    err_txt_423 = oNode_error.selectSingleNode("txt_423").Text
    err_txt_424 = oNode_error.selectSingleNode("txt_424").Text
    err_txt_425 = oNode_error.selectSingleNode("txt_425").Text
    err_txt_426 = oNode_error.selectSingleNode("txt_426").Text
    err_txt_427 = oNode_error.selectSingleNode("txt_427").Text
    err_txt_428 = oNode_error.selectSingleNode("txt_428").Text
    err_txt_429 = oNode_error.selectSingleNode("txt_429").Text
    err_txt_430 = oNode_error.selectSingleNode("txt_430").Text

    err_txt_431 = oNode_error.selectSingleNode("txt_431").Text
    err_txt_432 = oNode_error.selectSingleNode("txt_432").Text
    err_txt_433 = oNode_error.selectSingleNode("txt_433").Text
    err_txt_434 = oNode_error.selectSingleNode("txt_434").Text
    err_txt_435 = oNode_error.selectSingleNode("txt_435").Text
    err_txt_436 = oNode_error.selectSingleNode("txt_436").Text
    err_txt_437 = oNode_error.selectSingleNode("txt_437").Text
    err_txt_438 = oNode_error.selectSingleNode("txt_438").Text
    err_txt_439 = oNode_error.selectSingleNode("txt_439").Text
    err_txt_440 = oNode_error.selectSingleNode("txt_440").Text

    err_txt_441 = oNode_error.selectSingleNode("txt_441").Text
    err_txt_442 = oNode_error.selectSingleNode("txt_442").Text
    err_txt_443 = oNode_error.selectSingleNode("txt_443").Text
    err_txt_444 = oNode_error.selectSingleNode("txt_444").Text
    err_txt_445 = oNode_error.selectSingleNode("txt_445").Text
    err_txt_446 = oNode_error.selectSingleNode("txt_446").Text
    err_txt_447 = oNode_error.selectSingleNode("txt_447").Text
    err_txt_448 = oNode_error.selectSingleNode("txt_448").Text
    err_txt_449 = oNode_error.selectSingleNode("txt_449").Text
    err_txt_450 = oNode_error.selectSingleNode("txt_450").Text

    err_txt_451 = oNode_error.selectSingleNode("txt_451").Text
    err_txt_452 = oNode_error.selectSingleNode("txt_452").Text
    err_txt_453 = oNode_error.selectSingleNode("txt_453").Text
    err_txt_454 = oNode_error.selectSingleNode("txt_454").Text
    err_txt_455 = oNode_error.selectSingleNode("txt_455").Text
    err_txt_456 = oNode_error.selectSingleNode("txt_456").Text
    err_txt_457 = oNode_error.selectSingleNode("txt_457").Text
    err_txt_458 = oNode_error.selectSingleNode("txt_458").Text
    err_txt_459 = oNode_error.selectSingleNode("txt_459").Text
    err_txt_460 = oNode_error.selectSingleNode("txt_460").Text

    err_txt_461 = oNode_error.selectSingleNode("txt_461").Text
    err_txt_462 = oNode_error.selectSingleNode("txt_462").Text
    err_txt_463 = oNode_error.selectSingleNode("txt_463").Text
    err_txt_464 = oNode_error.selectSingleNode("txt_464").Text
    err_txt_465 = oNode_error.selectSingleNode("txt_465").Text
    err_txt_466 = oNode_error.selectSingleNode("txt_466").Text
    err_txt_467 = oNode_error.selectSingleNode("txt_467").Text
    err_txt_468 = oNode_error.selectSingleNode("txt_468").Text
    err_txt_469 = oNode_error.selectSingleNode("txt_469").Text
    err_txt_470 = oNode_error.selectSingleNode("txt_470").Text

    err_txt_471 = oNode_error.selectSingleNode("txt_471").Text
    err_txt_472 = oNode_error.selectSingleNode("txt_472").Text
    err_txt_473 = oNode_error.selectSingleNode("txt_473").Text
    err_txt_474 = oNode_error.selectSingleNode("txt_474").Text
    err_txt_475 = oNode_error.selectSingleNode("txt_475").Text
    err_txt_476 = oNode_error.selectSingleNode("txt_476").Text
    err_txt_477 = oNode_error.selectSingleNode("txt_477").Text
    err_txt_478 = oNode_error.selectSingleNode("txt_478").Text
    err_txt_479 = oNode_error.selectSingleNode("txt_479").Text
    err_txt_480 = oNode_error.selectSingleNode("txt_480").Text

    err_txt_481 = oNode_error.selectSingleNode("txt_481").Text
    err_txt_482 = oNode_error.selectSingleNode("txt_482").Text
    err_txt_483 = oNode_error.selectSingleNode("txt_483").Text
    err_txt_484 = oNode_error.selectSingleNode("txt_484").Text
    err_txt_485 = oNode_error.selectSingleNode("txt_485").Text
    err_txt_486 = oNode_error.selectSingleNode("txt_486").Text
    err_txt_487 = oNode_error.selectSingleNode("txt_487").Text
    err_txt_488 = oNode_error.selectSingleNode("txt_488").Text
    err_txt_489 = oNode_error.selectSingleNode("txt_489").Text
    err_txt_490 = oNode_error.selectSingleNode("txt_490").Text

    err_txt_491 = oNode_error.selectSingleNode("txt_491").Text
    err_txt_492 = oNode_error.selectSingleNode("txt_492").Text
    err_txt_493 = oNode_error.selectSingleNode("txt_493").Text
    err_txt_494 = oNode_error.selectSingleNode("txt_494").Text
    err_txt_495 = oNode_error.selectSingleNode("txt_495").Text
    err_txt_496 = oNode_error.selectSingleNode("txt_496").Text
    err_txt_497 = oNode_error.selectSingleNode("txt_497").Text
    err_txt_498 = oNode_error.selectSingleNode("txt_498").Text
    err_txt_499 = oNode_error.selectSingleNode("txt_499").Text
    err_txt_500 = oNode_error.selectSingleNode("txt_500").Text

    err_txt_501 = oNode_error.selectSingleNode("txt_501").Text
    err_txt_502 = oNode_error.selectSingleNode("txt_502").Text
    err_txt_503 = oNode_error.selectSingleNode("txt_503").Text
    err_txt_504 = oNode_error.selectSingleNode("txt_504").Text
    err_txt_505 = oNode_error.selectSingleNode("txt_505").Text
    err_txt_506 = oNode_error.selectSingleNode("txt_506").Text
    err_txt_507 = oNode_error.selectSingleNode("txt_507").Text
    err_txt_508 = oNode_error.selectSingleNode("txt_508").Text
    err_txt_509 = oNode_error.selectSingleNode("txt_509").Text
    err_txt_510 = oNode_error.selectSingleNode("txt_510").Text

    err_txt_511 = oNode_error.selectSingleNode("txt_511").Text
    err_txt_512 = oNode_error.selectSingleNode("txt_512").Text
    err_txt_513 = oNode_error.selectSingleNode("txt_513").Text
    err_txt_514 = oNode_error.selectSingleNode("txt_514").Text
    err_txt_515 = oNode_error.selectSingleNode("txt_515").Text
    err_txt_516 = oNode_error.selectSingleNode("txt_516").Text
    err_txt_517 = oNode_error.selectSingleNode("txt_517").Text
    err_txt_518 = oNode_error.selectSingleNode("txt_518").Text
    err_txt_519 = oNode_error.selectSingleNode("txt_519").Text
    err_txt_520 = oNode_error.selectSingleNode("txt_520").Text

    err_txt_521 = oNode_error.selectSingleNode("txt_521").Text
    err_txt_522 = oNode_error.selectSingleNode("txt_522").Text
    err_txt_523 = oNode_error.selectSingleNode("txt_523").Text
    err_txt_524 = oNode_error.selectSingleNode("txt_524").Text
    err_txt_525 = oNode_error.selectSingleNode("txt_525").Text
    err_txt_526 = oNode_error.selectSingleNode("txt_526").Text
    err_txt_527 = oNode_error.selectSingleNode("txt_527").Text
    err_txt_528 = oNode_error.selectSingleNode("txt_528").Text
    err_txt_529 = oNode_error.selectSingleNode("txt_529").Text
    err_txt_530 = oNode_error.selectSingleNode("txt_530").Text

    err_txt_531 = oNode_error.selectSingleNode("txt_531").Text
    err_txt_532 = oNode_error.selectSingleNode("txt_532").Text
    err_txt_533 = oNode_error.selectSingleNode("txt_533").Text
    err_txt_534 = oNode_error.selectSingleNode("txt_534").Text
    err_txt_535 = oNode_error.selectSingleNode("txt_535").Text
    err_txt_536 = oNode_error.selectSingleNode("txt_536").Text
    err_txt_537 = oNode_error.selectSingleNode("txt_537").Text
    err_txt_538 = oNode_error.selectSingleNode("txt_538").Text
    err_txt_539 = oNode_error.selectSingleNode("txt_539").Text
    err_txt_540 = oNode_error.selectSingleNode("txt_540").Text

    err_txt_541 = oNode_error.selectSingleNode("txt_541").Text
    err_txt_542 = oNode_error.selectSingleNode("txt_542").Text
    err_txt_543 = oNode_error.selectSingleNode("txt_543").Text
    err_txt_544 = oNode_error.selectSingleNode("txt_544").Text
    err_txt_545 = oNode_error.selectSingleNode("txt_545").Text
    err_txt_546 = oNode_error.selectSingleNode("txt_546").Text
    err_txt_547 = oNode_error.selectSingleNode("txt_547").Text
    err_txt_548 = oNode_error.selectSingleNode("txt_548").Text
    err_txt_549 = oNode_error.selectSingleNode("txt_549").Text
    err_txt_550 = oNode_error.selectSingleNode("txt_550").Text

    err_txt_551 = oNode_error.selectSingleNode("txt_551").Text
    err_txt_552 = oNode_error.selectSingleNode("txt_552").Text
    err_txt_553 = oNode_error.selectSingleNode("txt_553").Text
    err_txt_554 = oNode_error.selectSingleNode("txt_554").Text
    err_txt_555 = oNode_error.selectSingleNode("txt_555").Text
    err_txt_556 = oNode_error.selectSingleNode("txt_556").Text
    err_txt_557 = oNode_error.selectSingleNode("txt_557").Text

    err_txt_558 = oNode_error.selectSingleNode("txt_558").Text
    err_txt_559 = oNode_error.selectSingleNode("txt_559").Text

    err_txt_560 = oNode_error.selectSingleNode("txt_560").Text
    err_txt_561 = oNode_error.selectSingleNode("txt_561").Text
    err_txt_562 = oNode_error.selectSingleNode("txt_562").Text
    err_txt_563 = oNode_error.selectSingleNode("txt_563").Text
    err_txt_564 = oNode_error.selectSingleNode("txt_564").Text


    next




'Response.Write "tsa_txt_001: " & tsa_txt_001 & "<br>"
'Response.Write "tsa_txt_002: " & tsa_txt_002 & "<br>"


%>









