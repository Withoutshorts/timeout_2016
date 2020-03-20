<%response.buffer = true 
Session.LCID = 1030%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->


<%'** JQUERY START ************************* %>

<%'** JQUERY END ************************* %>

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->

<%'** ER SESSION UDLØBET  ****
    if len(session("user")) = 0 then
	
	errortype = 5
	call showError(errortype)
    response.end
    end if %>


<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<%

    '**** FASTE VAR ****
    media = request("media")
    onTheFly = request("OnTheFly")

    '*** Sætter lokal dato/kr format. Skal indsættes efter kalender.
	Session.LCID = 1030
	
	func = request("func")
	
	if len(trim(request("id"))) <> 0 then
	id = request("id")
	else
	id = 0
	end if
	
	
	
	rdir = request("rdir")


    '**** SLUT FASTE VAR *****
	

    public skptypeTxt
    function kptyp_sb(kptype)

        skptypeTxt = ""
        if cint(kptype) <> 0 then
            
            if cint(kptype) = 1 then
            kptypeTxt = "Faktura adr."
            else
            kptypeTxt = "Lev. adr."
            end if

        end if 
        
        if len(trim(kptypeTxt)) <> 0 then
        Response.write "<br>"& kptypeTxt
        end if

    end function
    

    if media <> "print" ANd rdir <> "fak" AND onTheFly <> "1" then    
        call menu_2014
    end if

    %>
    
    <div id="wrapper">
    <div class="content">
    
    <%
    select case func
    case "slet"

   
        oskrift = "Kontaktperson" 
        slttxta = "Du er ved at <b>SLETTE</b> en kontaktperson. Er dette korrekt?"
        slttxtb = "" 
        slturl = "kontaktpers.asp?func=sletok&id="& id

        call sletcnf_2015(oskrift, slttxta, slttxtb, slturl)

        
	
    case "sletok"

   
    kkundenavn = ""
    kpersnavn = ""

    strSQLk = "SELECT kp.kundeid, kkundenavn, kkundenr, kp.navn AS kpersnavn FROM kontaktpers AS kp "_
    &" LEFT JOIN kunder ON (kid = kp.kundeid) WHERE kp.id = "& id
    
    'response.write strSQLk
    'response.flush

    oRec.open strSQLk, oConn, 3
    if not oRec.EOF then

        kpersnavn = oRec("kpersnavn")
        kkundenavn = oRec("kkundenavn") & " ("& oRec("kkundenr") &")"

    end if
    ORec.close


    strSQL = "DELETE FROM kontaktpers WHERE id = "& id
	oConn.execute(strSQL)

    call insertDelhist("kpers", id, kpersnavn, kkundenavn, session("mid"), session("user"))

	
	response.Redirect "kontaktpers.asp"


    case "dbred", "dbopr"


    function SQLBless2(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, "'", "''")
	SQLBless2 = tmp
	end function
			
			
			if len(trim(request("FM_navn_kpers"))) = 0 then
			%>
			<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
			<%
			useleftdiv = "j2"
			errortype = 138
			call showError(errortype)
			
			Response.end
			end if
			
			
			'******************** Opretter kontaktpersoner ***********************
				
				
					useUpdateDato = year(now)&"/"&month(now)&"/"&day(now)
	
					thiskpersId = id

                    if len(request("kid")) <> 0 then
	                kid = request("kid")
                    else
                    kid = 0
                    end if

					
					if thiskpersId <> 0 then
					'** finder eksiterede pw, så der kan tjekkes om det er ændret og om der skal afsendes email ***'
					            
					             ekspw = ""
					             
					             strSQL = "SELECT password FROM kontaktpers WHERE id = " & thiskpersId
                                 oRec.open strSQL, oConn, 3
                                 if not oRec.EOF then
                                 ekspw = oRec("password")
                                 end if
                                            
                                 oRec.close
					
					end if
					
					
					kundeId = kid
					
					strKpers = SQLBless2(request("FM_navn_kpers"))
					strKperstit = SQLBless2(request("FM_titel_kpers"))
					strKpersdTlf = SQLBless2(request("FM_tlf_kpers"))
					strKpersMobil = SQLBless2(request("FM_mtlf_kpers"))
					strKpersPrivat = SQLBless2(request("FM_ptlf_kpers"))
					strKpersEmail = SQLBless2(request("FM_email_kpers"))
					strKperspw = SQLBless2(request("FM_pw_kpers"))
					strAdr_kpers = SQLBless2(request("FM_adr_kpers"))
					strPostnr_kpers = SQLBless2(request("FM_postnr_kpers"))
					strCity_kpers = SQLBless2(request("FM_city_kpers"))
					strLand_kpers = SQLBless2(request("FM_land_kpers"))
					strAf_kpers = SQLBless2(request("FM_afdeling_kpers"))
					strBeskkpers = SQLBless2(Request("FM_besk_kpers"))
					
                    kpean = SQLBless2(request("FM_pkean"))
                    kpean = replace(kpean, " ", "")

                    kpcvr = SQLBless2(request("FM_pkcvr"))
                    kpcvr = replace(kpcvr, " ", "") 

					kptype = request("FM_kptype")

                    '**** Nulstiller forvalgte faktura adr / Leverings adr *****'
                    if cint(kptype) = 1 then
                    strSQLfvFak = "UPDATE kontaktpers SET kptype = 0 WHERE kptype = 1 AND kundeId = "& kundeId
                    oConn.execute(strSQLfvFak)
                    end if

                    if cint(kptype) = 2 then
                    strSQLfvFak = "UPDATE kontaktpers SET kptype = 0 WHERE kptype = 2 AND kundeId = "& kundeId
                    oConn.execute(strSQLfvFak)
                    end if

                    '************************************************************


					if func = "dbopr" then
					
					strSQLKpers = "INSERT INTO kontaktpers ("_
					&" kundeid, editor, dato, navn, titel, adresse, "_
					&" postnr, town, land, afdeling, email, password, "_
					&" dirtlf, mobiltlf, privattlf, beskrivelse, kpean, kptype, kpcvr) VALUES "_
					&" ("& kundeId &", '"& session("user") &"', '" & useUpdateDato &"', '" & strKpers &"', '"& strKperstit &"', "_
					&" '"& strAdr_kpers &"', '"& strPostnr_kpers &"', '"& strCity_kpers &"', "_
					&" '"& strLand_kpers &"', '"& strAf_kpers &"', "_ 
					&" '"& strKpersEmail &"', '"& strKperspw &"', '"& strKpersdTlf &"', "_
					&" '"& strKpersMobil &"', '"& strKpersPrivat &"', '" & strBeskkpers &"', '"& kpean &"', "& kptype &", '"& kpcvr &"')"
					
					else
					
					strSQLKpers = "UPDATE kontaktpers SET "_
					&" kundeid = "& kundeId &", editor ='"& session("user") &"', dato = '" & useUpdateDato &"', "_
					&" navn = '" & strKpers &"', titel = '"& strKperstit &"', adresse = '"& strAdr_kpers &"', "_
					&" postnr = '"& strPostnr_kpers &"', town = '"& strCity_kpers &"', land = '"& strLand_kpers &"', "_
					&" afdeling = '"& strAf_kpers &"' , email = '"& strKpersEmail &"', password = '"& strKperspw &"', "_
					&" dirtlf = '"& strKpersdTlf &"', mobiltlf = '"& strKpersMobil &"', privattlf = '"& strKpersPrivat &"', "_
					&" beskrivelse = '" & strBeskkpers &"', kpean = '"& kpean &"', kptype = "& kptype &", kpcvr = '"& kpcvr &"'"_
					&" WHERE id = " & thiskpersId 
					
					end if
					
					'response.write strSQLKpers
					
                    'Response.flush
                    oConn.execute(strSQLKpers)
					
					
					 'if lto = "outz" then        
                               
					if len(trim(strKpersEmail)) <> 0 AND len(trim(strKperspw)) <> 0 AND ekspw <> strKperspw then
					 '*** Sender mail, hvis der er givet adgang til kunde område for kpers..**
                               '***** Oprettter Mail object ***
                               if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\kontaktpers.asp" then
                                
                              
           
                                 Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
                                 ' Sætter Charsettet til ISO-8859-1 
                                 Mailer.CharSet = 2
                                 ' Afsenderens navn 
                                 Mailer.FromName = "TimeOut ServiceDesk"
                                 ' Afsenderens e-mail 
                                 Mailer.FromAddress = "timeout_no_reply@outzource.dk"
                                 Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
                                 Mailer.ContentType = "text/html"
                                 
                                 
                                  '** Henter key ***
                                 lkey = ""
                                 strSQL = "SELECT l.key as lkey FROM licens l WHERE id = 1"
                                 oRec.open strSQL, oConn, 3
                                 if not oRec.EOF then
                                 lkey = oRec("lkey")
                                 end if
                                            
                                 oRec.close
                                
                          
                          
                                '** Henter Afsender ***
                                 strSQL = "SELECT mnavn, mnr, init, email FROM medarbejdere WHERE mid = " & session("mid")
                                 oRec.open strSQL, oConn, 3
                                 if not oRec.EOF then
                                            
                                 afsNavn = oRec("mnavn") & "("& oRec("mnr") &") - " & oRec("init")
                                 afsEmail = oRec("email")
                                            
                                 
                                 end if
                                            
                                 oRec.close
                                
                                
                                'response.Write "ok - afsender navn:" & afsNavn
                                'response.end 
                                
                            
                                 
                                 
                                'Mailens emne
                                 Mailer.Subject = "Du er blevet tildelt adgang til vores ServiceDesk"
                                 ' Modtagerens navn og e-mail
                                 Mailer.AddRecipient strKnavn, strKpersEmail
                                 
                          
                                 
                                 ' Selve teksten
                                                      Mailer.BodyText = "Hej  "& strKpers &",<br><BR>" _ 
                                                      &" Du modtager denne mail fordi du er blevet tildelt adgang / har fået opdateret adgang til vores ServiceDesk.<br><br>"_
                                                      & "I vores ServiceDesk kan du oprette Incidents og følge med i status på igangværende Incidents.<BR>" _
                                                      & "Du kan også følge med i timeforbrug på job og aftaler hvortil du er blevet tildelt adgang.<br><br>" _
                                                      & "Du finder vores ServiceDesk her:<br>"_
                                                      & "<a href=""https://outzource.dk/timeout_xp/wwwroot/ver2_1/login_kunder.asp?key="& lkey &"&lto="& lto &""">"_
                                                      & "https://outzource.dk/timeout_xp/wwwroot/ver2_1/login_kunder.asp?key="& lkey &"&lto="& lto &"</a><br><br>"_
                                                      & "Dit logind og password er:<br>"_
                                                      & ""& strKpersEmail &" / " & strKperspw & "<br>"_
                                                      & "<br><br><br>"_
                                                      & "Med venlig hilsen<br>" & afsNavn & ", "& afsEmail & "<br>"_ 
                                                      '& "Incident ansvarlig: " & ansvNavn & ", " & ansvEmail & vbCrLf & vbCrLf _ bcc
                                                      '& "Adressen til TimeOut er: https://outzource.dk/"&lto&"" & vbCrLf & vbCrLf _ 
                                            
                                                      If Mailer.SendMail Then
                                            
                                                      Else
                                                      Response.Write "Fejl...<br>" & Mailer.Response
                                                      End if
                                                      
                             
                            
                             end if     '** C drev: mailer ** 		
							 
							
				        end if

                        if rdir <> "treg" AND rdir <> "fak" then
					        Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
				        end if

                        if onTheFly = "1" then
                            Response.Write("<script language='JavaScript'>window.close();</script>")
                        else
                            response.Redirect "kontaktpers.asp"
                        end if

    case "red", "opr"


                    if len(request("kid")) <> 0 then
	                kid = request("kid")
                    else
                    kid = 0
                    end if
    

    if func = "red" then

        dbfunc = "dbred"

        strSQLKpers = "SELECT id, kundeid, navn, adresse, postnr, town, land, afdeling, email, password, dirtlf, mobiltlf, "_
        &" privattlf, beskrivelse, titel, kpean, kptype, kpcvr FROM kontaktpers WHERE id ="& id
		oRec2.open strSQLKpers, oConn, 3	
		if not oRec2.EOF then				
		
		strKpers = oRec2("navn") 'strKpers1
		strKperstit = oRec2("titel") 'str_titel_kpers1
		strKpersTlf = oRec2("privattlf")
		strKpersmTlf = oRec2("mobiltlf")
		strKpersdTlf = oRec2("dirtlf")
		strKpersEmail = oRec2("email")
		strKperspw = oRec2("password")
		strAdr_kpers = oRec2("adresse")
		strPostnr_kpers = oRec2("postnr")
		strCity_kpers = oRec2("town")
		strLand_kpers = oRec2("land")
		strAf_kpers = oRec2("afdeling")
		strBeskkpers = oRec2("beskrivelse")
		thisKid = oRec2("kundeid")
		kpean = oRec2("kpean") 
        kptype = oRec2("kptype")
        kpcvr = oRec2("kpcvr")

		end if
		oRec2.close


        funcTxt = "Rediger"

    else

        dbfunc = "dbopr"
        thisKid = kid
        funcTxt = "Opret"

    end if

                
        kptypSEL0 = ""
        kptypSEL1 = ""
        kptypSEL2 = ""
        select case kptype 
        case 0
        kptypSEL0 = "SELECTED"
        case 1
        kptypSEL1 = "SELECTED"
        case 2
        kptypSEL2 = "SELECTED"
        end select

    %>


<!--Kontaktperson redigering-->

    <div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u>Kontaktperson - <%=funcTxt %></u>
        </h3>
          <form action="kontaktpers.asp?func=<%=dbfunc%>&id=<%=id %>" method="post">
            <input type="hidden" name="OnTheFly" value="<%=onTheFly %>" />
        <!-- Opdater/Annuller knapper -->
        <div style="margin-top:-15px;margin-bottom:15px;">
          <button type="submit" class="btn btn-success btn-sm pull-right" id="sbm_upd0"><b>Opdatér</b></button>
          
          <div class="clearfix"></div>
        </div>

        <div class="portlet-body">
            <!-- NAVN / SORTERING / ID -->
         <section>
                <div class="well well-white">
                    <div class="row">
                        <div class="col-lg-12">
                            <h4 class="panel-title-well">Stamdata</h4>
                        </div>
                    </div>

                    <div class="row">
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">Navn:&nbsp<span style="color:red;">*</span></div>
                        <div class="col-lg-4"><input type="text" class="form-control input-small" name="FM_navn_kpers" value="<%=strKpers%>" placeholder="Navn"/></div>
                         
                    </div>
                    <div class="row">
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">Titel:</div>
                        <div class="col-lg-4"><input type="text" class="form-control input-small" value="<%=strKperstit%>" name="FM_titel_kpers"/>
                            <span style="color:#999999; font-size:10px;">(kundenr. ell. beskrivelse hvis type = faktura adr.)</span>
                        </div>
                    </div>


                    <div class="row">
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">Kundetilhørsforhold:&nbsp;<span style="color:red;">*</span></div>
                        <div class="col-lg-4">
                            <select class="form-control input-small" id="kid" name="kid">
			                <%
			                strSQL = "SELECT kkundenavn, kkundenr, kid FROM kunder WHERE kid <> 0 ORDER BY kkundenavn"
			                oRec.open strSQL, oConn, 3 
			                while not oRec.EOF 
				                if cint(thisKid) = cint(oRec("kid")) then
				                selThis = "SELECTED"
				                else
				                selThis = ""
				                end if%>
			                <option value="<%=oRec("kid")%>" <%=selThis %>><%=oRec("kkundenavn")%>  (<%=oRec("kkundenr")%>)</option>
			                <%
			                oRec.movenext
			                wend
			                oRec.close 
			                %>
		                   </select>
                        </div>
                    </div>

                     <div class="row">
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">Type:&nbsp;<span style="color:red;">*</span></div>
                         <div class="col-lg-4"><select name="FM_kptype"  class="form-control input-small">
                            <option value="0" <%=kptypSEL0 %>>Kontaktperson</option>
                            <option value="1" <%=kptypSEL1 %>>Faktura adr. (forvalgt)</option>
                            <option value="2" <%=kptypSEL2 %>>Leverings adr. (forvalgt)</option>
                        </select></div>
                    </div>
        </section>
      

        <section>
                <!-- Accordion -->
                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <!-- PersonData -->
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseOne">Kontaktdata</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseOne" class="panel-collapse collapse">
                        <div class="panel-body">

                             <div class="row">
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2">Adresse:</div>
                        <div class="col-lg-6"><input type="text" name="FM_adr_kpers" class="form-control input-small" value="<%=strAdr_kpers%>" style="height:60px;"/></div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Postnr.:</div>
                                <div class="col-lg-2">
                                <input class="form-control input-small" type="text" name="FM_postnr_kpers" value="<%=strPostnr_kpers%>">
                                </div>
                                <div class="col-lg-7">&nbsp</div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">By:</div>
                                <div class="col-lg-4">
                                <input class="form-control input-small" type="text" name="FM_city_kpers" value="<%=strCity_kpers%>"">
                                </div>
                                
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Land:</div>
                                <div class="col-lg-4">
                                <select class="form-control input-small" name="FM_land_kpers">
		                        <%if func = "red" then%>
		                        <option SELECTED><%=strLand_kpers%></option>
		                        <%else%>
		                        <option SELECTED>Danmark</option>
		                        <%end if%>
		                        <!--#include file="../timereg/inc/inc_option_land.asp"-->
		                        </select>
                                </div>
                                <div class="col-lg-6">&nbsp</div>
                            </div>
                            
                              <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Afdeling:</div>
                                <div class="col-lg-4">
                                <input type="text" class="form-control input-small" name="FM_afdeling_kpers" value="<%=strAf_kpers%>">
                                </div>
                                
                            </div>
                            

                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Telefon (direkte):</div>
                                <div class="col-lg-2">
                                <input class="form-control input-small" type="text" name="FM_tlf_kpers" value="<%=strKpersdTlf%>"/>
                                </div>
                                <div class="col-lg-7">&nbsp</div>
                            </div>
                              <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Telefon (privat):</div>
                                <div class="col-lg-2">
                                <input class="form-control input-small" type="text" name="FM_ptlf_kpers" value="<%=strKpersTlf%>"/>
                                </div>
                                <div class="col-lg-7">&nbsp</div>
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Mobil:</div>
                                <div class="col-lg-2">
                                <input class="form-control input-small" type="text" name="FM_mtlf_kpers" value="<%=strKpersmTlf%>"/>
                                </div>
                                <div class="col-lg-7">&nbsp</div>
                            </div>
                           
                          
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Email:
                                      <span style="color:#999999;">(login)</span>
                                </div>
                                <div class="col-lg-4">
                                <input class="form-control input-small" type="text" name="FM_email_kpers" value="<%=strKpersEmail%>" />
                                </div>
                               
                            </div>

                             <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">Pasword:<br />
                                <span style="color:#999999;">Giver adgang til TimeOuts kunde-del.</span>
                                </div>
                                <div class="col-lg-2">
                                <input class="form-control input-small" type="text" name="FM_pw_kpers" value="<%=strKperspw%>" />
                                </div>
                                
                            </div>

                            

                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">EAN:</div>
                                <div class="col-lg-2">
                                <input class="form-control input-small" type="text" name="FM_pkean" value="<%=kpean%>"/>
                                </div>
                                
                            </div>
                            <div class="row">
                                <div class="col-lg-1">&nbsp</div>
                                <div class="col-lg-2">CVR:</div>
                                <div class="col-lg-2">
                                <input class="form-control input-small" type="text" name="FM_pkcvr" value="<%=kpcvr%>"/>
                                </div>
                                
                            </div>
                            

                                <div class="row">
                                    <div class="col-lg-1">&nbsp</div>
                                    <div class="col-lg-10">Beskrivelse:
                                    <textarea class="form-control input-small" name="FM_besk_kpers" style="height:120px;"><%=strBeskkpers%></textarea>
                                    </div>
                                    
                                </div>

                                

                        </div>
                        </div>
                        </div>

                                 <div style="margin-top:25px; margin-bottom:15px;">

                                     <%if func = "red" then %>
                                     <a class="btn btn-primary btn-sm" role="button" href="kontaktpers.asp?func=slet&id=<%=id%>&kid=<%=kid %>"><b>Slet</b></a>&nbsp;
		                             <%end if %>


                                <button type="submit" class="btn btn-success btn-sm pull-right" id="sbm_upd0"><b>Opdatér</b></button>
                                 <div class="clearfix"></div>
                                </div>

                

                   
                 </div>
        </section>
         </form>
       </div> <!-- /.portlet-body -->
      </div> <!-- /.portlet -->

<% 

case else    
%>


    <script src="js/kpers_liste.js" type="text/javascript"></script>


    <!--kontaktperson liste-->

    <%
        
        if len(trim(request("issogsubmitted"))) <> 0 then
	    kid = request("kid")

        response.cookies("kontaktpers_2015")("kid") = kid
        response.cookies("kontaktpers_2015").expires = date + 1

        else

            if request.cookies("kontaktpers_2015")("kid") <> "" then
            kid = request.cookies("kontaktpers_2015")("kid")
            else
            kid = 0
            end if

	    end if
	
	    'intKid = kid

        
        if len(trim(request("issogsubmitted"))) <> 0 then
            thisSogekri = request("FM_soeg")

            response.cookies("kunder_2015")("kperssogekri") = thisSogekri
	        response.cookies("kunder_2015").expires = date + 1
	    
	    else
	        
	        if request.cookies("kunder_2015")("kperssogekri") <> "" then
	        thisSogekri = request.cookies("kunder_2015")("kperssogekri")
	        
	        else
	        thisSogekri = ""
	        
	        end if
	        
	    end if


        if len(trim(request("issogsubmitted"))) <> 0 OR thisSogekri <> "" OR kid <> 0 then
        visKpersListe = 1
        else 
        visKpersListe = 0
        end if
        %>

    <div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u>Kontaktpersoner / filialer</u>
        </h3>
       
         <form action="kontaktpers.asp?func=opr&kid=<%=kid %>" method="post"> 
              
         <section>
                <div class="row">
                    <div class="col-lg-10">&nbsp;</div>
                    <div class="col-lg-2">
                <button class="btn btn-sm btn-success pull-right"><b>Opret ny +</b></button><br />&nbsp;
                </div>
            </div>
        </section>
         </form>
        

           <form action="kontaktpers.asp?issogsubmitted=1" method="POST">
          <section>
                <div class="well">
                         <h4 class="panel-title-well">Søgefilter</h4>
                         <div class="row">
                             <div class="col-lg-3">
                                 Søg på Navn: <br />
			                    <input type="text" name="FM_soeg" id="FM_soeg" class="form-control input-small" placeholder="% Wildcard" value="<%=thisSogekri%>">

                               </div>
                              
                         


                         
                             <div class="col-lg-4">
                               Kunde:<select name="kid" class="form-control input-small" onchange="submit()">
				                    <option value="0">Alle</option>
				                    <%
						                    strSQL = "SELECT kkundenavn, kkundenr, kid FROM kunder WHERE kid <> 0 ORDER BY kkundenavn"
						                    oRec.open strSQL, oConn, 3
						                    while not oRec.EOF
						
							                    if cint(kid) = cint(oRec("kid")) then
							                    isSelected = "SELECTED"
							                    else
							                    isSelected = ""
							                    end if
						
						                    %>
						                    <option value="<%=oRec("kid")%>" <%=isSelected%>><%=oRec("kkundenavn") &" ("& oRec("kkundenr") & ")"%></option>
						                    <%
						                    oRec.movenext
						                    wend
						                    oRec.close
						                    %>
				                    </select>

                               </div>

                               

                               <div class="col-lg-5"><br />
                                    <button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Vis kontaktpersoner >></b></button>
                                   </div>
                                   <!--<input type="submit" class="btn btn-sm btn-secondary pull-right" value=" Søg " /></div>--> 

                         </div><!-- ROW -->

                              


                    </div><!-- Well Well White -->
                    </section>
              </form>




        <div class="portlet-body">
          <table id="tb_kontaktsliste" class="table table-striped table-bordered table-hover ui-datatable" > 
           
            <thead>
              <tr>
                <th style="width: 20%">Kunde</th>
                <th style="width: 30%">Kontaktperson</th>
                <th style="width: 20%">Afdeling</th>
                 <th style="width:10%">Email</th>
                <th style="width: 10%">Mobil</th>
                <th style="width: 10%">Tel. dir.</th>
               
              </tr>
            </thead>
              <tbody>

                  <%
	    '**Søg
        if cint(visKpersListe) <> 0 then

	            if thisSogekri <> "" then
	            strSQLWclaus = " (p.navn LIKE '"& thisSogekri &"%' OR p.email LIKE '%"& thisSogekri &"%' OR p.afdeling LIKE '%"& thisSogekri &"%') "
	            else
	            strSQLWclaus = " p.id <> 0 "
	            end if

        else
                strSQLWclaus = " p.id = 0 "

        end if
	
	    if cint(kid) = 0 then
	    strSQLWclaus = strSQLWclaus
	    'vallkunderChk = "CHECKED"
	    else
	    'vallkunderChk = ""
	    strSQLWclaus = strSQLWclaus & " AND p.kundeid = "& kid &""
	    end if
		
        
		
		strSQLKpers = "SELECT p.id, p.kundeid, p.navn, p.adresse, p.postnr, p.town, "_
		&" p.land, p.afdeling, p.email, p.password, p.dirtlf, p.mobiltlf, p.privattlf, "_
		&" p.beskrivelse, p.titel, k.kkundenavn, kpean, kptype, k.kid FROM kontaktpers p "_
		&" LEFT JOIN kunder k ON (k.kid = p.kundeid) WHERE "& strSQLWclaus &" AND navn <> '0' ORDER BY kkundenavn, p.navn LIMIT 2000"  '"& useThisKpers


		oRec2.open strSQLKpers, oConn, 3	
		t = 0
		while not oRec2.EOF 			
		
		intKpersId = oRec2("id")
		strKpers = oRec2("navn") 'strKpers1
		strKperstit = oRec2("titel") 'str_titel_kpers1
		strKpersTlf = oRec2("privattlf")
		strKpersmTlf = oRec2("mobiltlf")
		strKpersdTlf = oRec2("dirtlf")
		strKpersEmail = oRec2("email")
		strKperspw = oRec2("password")
		strAdr_kpers = oRec2("adresse")
		strPostnr_kpers = oRec2("postnr")
		strCity_kpers = oRec2("town")
		strLand_kpers = oRec2("land")
		strAf_kpers = oRec2("afdeling")
		strBeskkpers = oRec2("beskrivelse")
		strKundenavn = oRec2("kkundenavn")
		
        kpean = oRec2("kpean")
        kptype = oRec2("kptype")

		%> 
              <tr>
                <td><a href="kunder.asp?func=red&id=<%=oRec2("kid")%>"><%=oRec2("kkundenavn") %></a></td>   
                <td><a href="kontaktpers.asp?func=red&id=<%=oRec2("id") %>"><%=oRec2("navn") %></a>

                    <%
                        
                        if cint(kptype) <> 0 then
            
                            if cint(kptype) = 1 then
                            kptypeTxt = "Faktura adr."
                            else
                            kptypeTxt = "Lev. adr."
                            end if


                        Response.write " ("& kptypeTxt &")"

                        end if 
                        
                        
                     %>
                        



                </td>
                <td><%=strAf_kpers %></td>
                  <td><a href=mailto:<%=strKpersEmail %>><%=strKpersEmail %></a></td>
                <td><%=oRec2("mobiltlf") %></td>
                <td><%=oRec2("dirtlf") %></td>
                
              </tr>
            <%oRec2.movenext
                wend
                oRec2.close
                 %>  
            </tbody>
            <tfoot>
               <tr>
                    <th>Kunde</th>
                    <th>Kontaktperson</th>
                    <th>Afdeling</th>
                    <th>Email</th>
                    <th>Mobil</th>
                    <th>Tel dir.</th>
              </tr>
            </tfoot>
          </table>
        </div> <!-- /.portlet-body -->
      </div> <!-- /.portlet -->
        




<%end select  %>

        </div> <!-- /.container -->   
        </div> <!-- /.content -->   
        </div> <!-- /.wrapper -->   




<!--#include file="../inc/regular/footer_inc.asp"-->