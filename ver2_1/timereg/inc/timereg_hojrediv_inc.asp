<%'GIT 20160811 - SK

function hojreDiv()

    '*************************************************************'
    '************ Højre mnubar på timereg. siden *****************'
    '**************** Aktive job, job til afv. +TODO LIST ********'
    '*************************************************************'




      if browstype_client <> "ip" then '**Vis ikke på iPhone
      
      
      '**Vis ikke højremenu ofr disse '***
      if (lto = "intranet - local") OR (lto = "demo") OR (lto = "adra") OR (lto = "wwf") OR (lto = "ngf") OR (lto = "ascendis") OR (lto = "akelius") OR _
      (lto = "hvk_bbb") OR (lto = "mmmi" AND meType = 10) OR (lto = "xdencker") OR (lto = "biofac") OR (lto = "epi_uk") OR (lto = "tec") OR (lto = "esn") OR _
      (lto = "micmatic") OR (lto = "aalund") OR (lto = "nonstop") OR (lto = "oko") OR (lto = "wilke") OR (lto = "cisu") OR (lto = "hidalgo") OR (lto = "bf") OR (lto = "cst") OR (lto = "plan") then

            if lto = "wwf" OR lto = "xintranet - local" then 
            %>
            <div id="favoritlinks" style="position:absolute; left:870px; top:378px; width:212px; visibility:visible; display:; z-index:2;">
            <table cellpadding=0 cellspacing=0 border=0>
                <tr><td>

                    <b>Favorit links</b><br />
                    <a href="joblog_timetotaler.asp?FM_medarb=<%=usemrn %>&vis_restimer=1&nomenu=1" target="_blank" class="vmenu">Se personligt timeforbrug og budget >></a>

                    </td></tr></table>
                </div>
        
        
            <%end if


            '** SKAL VISES FOR ALLE HVOR TIMEBUDGET SIM ER SLÅET TIL
            if lto = "oko" OR lto = "wwf" OR lto = "intranet - local" OR lto = "wilke" OR lto = "hidalgo" OR lto = "bf" then 

            select case lto '** Bør kalde finasår statdato
            case "wwf", "intranet - local"

                if month(tjekdag(1)) >= 7 then
                aarFinansSt = year(tjekdag(1)) & "-7-1"
                aarFinansSl = year(tjekdag(1))+1 & "-6-30"
                else
                aarFinansSt = year(tjekdag(1))-1 & "-7-1"
                aarFinansSl = year(tjekdag(1)) & "-6-30"
                end if

                strDatoRs = " AND ((md >= "& month(aarFinansSt) &" AND aar = "& year(aarFinansSt) &") OR (md <= "& month(aarFinansSl) &" AND aar = "& year(aarFinansSl) &"))"
                
            case else
                 aarFinansSt = year(tjekdag(1)) & "-1-1"
                 aarFinansSl = year(tjekdag(1)) & "-12-31"

                 strDatoRs = " AND ((md >= "& month(aarFinansSt) &" AND aar = "& year(aarFinansSt) &") AND (md <= "& month(aarFinansSl) &" AND aar = "& year(aarFinansSl) &"))"

            end select

                

            %>
            <div id="fctimer" style="position:absolute; left:870px; top:408px; width:212px; visibility:visible; display:; z-index:2;">
            
            <table cellpadding=2 cellspacing=0 border=0 width="100%">
                
                <tr bgcolor="#5C75AA"><td colspan="2" class="alt"><b>Dine job med timebudget/forecast FY </b><br /></td></tr>
                
                    
                    
                    <%strSQLr = "SELECT r.jobid, SUM(r.timer) AS restimer, j.jobnavn, j.jobnr FROM ressourcer_md AS r "_
                    &" LEFT JOIN job AS j ON (j.id = r.jobid) "_
                    &" LEFT JOIN aktiviteter AS a ON (a.id = r.aktid) "_
                    &" WHERE r.medid = "& usemrn & " AND r.aktid <> 0 AND a.navn IS NOT NULL AND jobstatus = 1 "& strDatoRs &" GROUP BY r.jobid ORDER BY jobnavn LIMIT 50" 

                        'response.write strSQLr
                        'response.flush
                        re = 0
                        oRec.open strSQLr, oConn, 3
                        While not oRec.EOF 

                        select case right(re, 1)
                        case 0,2,4,6,8
                        bgcr = "#FFFFFF"
                        case else
                        bgcr = "#EfF3ff"
                        end select
                        %>
                        <tr style="background-color:<%=bgcr%>;">
                            <td class="lille"><%=left(oRec("jobnavn"), 20) & " ("& oRec("jobnr") &")" %></td>
                            <td class="lille" align="right"><%=oRec("restimer") %> t.</td>
                        </tr>
                        <%

                        re = re + 1
                        oRec.movenext
                        wend
                        oRec.close      
                        
                     
                    if re = 0 then   
                    %>
                    <tr bgcolor="#FFFFFF"><td colspan="2">- ingen </td></tr>
                    <%end if %>

                    </table>
                </div>
        
        
            <%end if%>

            <%
      else
      %>



     
       <div id="hmenu" style="position:absolute; left:870px; top:478px; width:212px; visibility:visible; display:; z-index:2;">
   <table cellpadding=0 cellspacing=0 border=0>
        <tr>
            <td align=center id="a_denneuge" style="white-space:nowrap; width:100px; padding:1px; background-color:#EFF3FF; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="#" id="A1" class=rmenu>+ Igangv.</a></td>
             <td align=center id="a_todos" style="white-space:nowrap; width:120px; padding:1px; background-color:#EFF3FF; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="#" id="A2" class=rmenu>+ ToDo's</a></td>
			</td>
            <%if level <= 2 then%>
			<td align=center id="a_tildel" style="white-space:nowrap; width:120px; padding:1px; background-color:#EFF3FF; border-right:1px #d6dff5 solid; border-bottom:0px;">
			<a href="#" id="A3" class=rmenu>+ Tildel job</a></td>
			<%end if%>
        </tr>
    </table>
    </div>

        <div id="div_tildel" style="position:absolute; left:870px; top:493px; width:210px; height:600px; visibility:hidden; display:none; z-index:100; border:1px #cccccc solid; padding:2px; background-color:#ffffff; overflow:auto;">
	     <form action="timereg_akt_2006.asp?func=opdaterjlist" method="post">
         <table cellpadding=2 cellspacing=0 border=0 width=100%>
         
         <tr bgcolor="#5C75AA"><td class=alt style="width:165px;" colspan=2><b>
         Tildel job til aktivjobliste:</b>
         </td>
       
         </tr>

         <tr><td class=lille colspan=2>
        
         <%
         if len(trim(request("tildel_sel_pgrp"))) <> 0 then
         tildel_sel_pgrpSEL = trim(request("tildel_sel_pgrp"))
         else
                if request.cookies("tsa")("td_prg_sel") <> "" then
                tildel_sel_pgrpSEL = request.cookies("tsa")("td_prg_sel")
                else
                tildel_sel_pgrpSEL = 0
                end if
         end if
         
         Response.cookies("tsa")("td_prg_sel") = tildel_sel_pgrpSEL

         strSQLtdgrp = "SELECT id, navn FROM projektgrupper WHERE id <> 0 ORDER BY navn " 
         oRec3.open strSQLtdgrp, oConn, 3
         %>
          Projektgrupper:<br /> <select id="tildel_sel_pgrp" name="tildel_sel_pgrp" style="width:180px; font-size:9px;">
         <%
         while not oRec3.EOF 
            
            if cint(tildel_sel_pgrpSEL) = cint(oRec3("id")) then
            selPrgrp = "SELECTED"
            else
            selPrgrp = ""
            end if
            
            %>
         <option value="<%=oRec3("id") %>" <%=selPrgrp %>><%=oRec3("navn") %></option>
         <%
         oRec3.movenext
         wend
         oRec3.close
         %>
       
         
         </select>

         
         <br /><br />
         Medarbejder(e):<br /> 
         <div id="div_tildeljoblisten_msel"></div>
        
        </td></tr>

         <tr><td colspan=2 valign=top class=lille>
         <div id="div_tildeljoblisten"></div>
         </td></tr>
        
         </table>
         </form>
        </div>

            
          <%select case lto
          case "dencker", "intranet - local", "fe"
          jdu_lmt = 20
          case "epi", "epi_no", "epi_sta", "epi_ab"
          jdu_lmt = 100
          case else
          jdu_lmt = 50
          end select %>        



        <div id="div_denneuge" style="position:absolute; left:870px; top:493px; width:212px; height:600px; visibility:<%=jinf_knap_vzb%>; display:<%=jinf_knap_dsp%>; z-index:100; border:1px #cccccc solid; padding:2px; background-color:#ffffff; overflow:auto;">
	     <table cellpadding=2 cellspacing=0 border=0 width=100%>
         <form method="post" action="#">
         <tr bgcolor="#5C75AA"><td class=alt><b>
         <%=tsa_txt_376 %>:</b>
         </td></tr>
         <tr><td class=lille>
         Viser aktive job (maks <%=jdu_lmt %>) hvor Du er: <br /><input type="checkbox" value="1" id="denneuge_jobans" name="denneuge_jobans" /> Jobansvarlig / jobejer
         <br />
         Vis seneste 2 uger + 
         <select id="denneuge_ignrper" style="font-size:9px; width:50px;">
            <option value="1">1 uge</option>
            <option value="2">2 uge</option>
            <option value="3">3 uge</option>
            <option value="4">4 uge</option>
         </select> frem
         <input type="hidden" value="<%=jdu_lmt %>" id="denneuge_limit"/>
         <!--
         <input type="checkbox" value="1" name="denneuge_ignrper" id="denneuge_ignrper" /> Slutdato uge <%=datepart("ww", tjekdag(1), 2,2) %> <br /> 
         ellers +1 md. og forrige 3 md.
         -->
         <br /><input type="checkbox" value="1" name="denneuge_alfa" id="denneuge_alfa" />  Sortér alfabetisk
         </td></tr>
         <input type="hidden" value="<%=lto%>" name="dennuge_lto" id="dennuge_lto" />
         

         <tr><td>

         <div id="jobidenneuge">
        </div>

         </td></tr>
         </form>
         </table>
        </div>


	    
        <div id="div_todolist" style="position:absolute; left:870px; top:493px; width:212px; height:600px; visibility:hidden; display:none; z-index:100; border:1px #cccccc solid; padding:2px; background-color:#ffffff; overflow:auto;">
	    
       
        <table cellpadding=2 cellspacing=0 border=0 width=100%>
         <tr bgcolor="#5C75AA"><td id="todotd" colspan="2" style="padding:3px; height:20px;" class=alt><b>ToDo liste</b> (seneste 20)</td></tr>
         <tr><td colspan="2"> <a href="webblik_todo.asp?nomenu=1" class=rmenu target="_blank">Se / rediger ToDo listen >></a></td></tr>
         <tr><td class=lille style="padding-left:5px;">ToDo's</td><td class=lille align=right style="padding-right:5px;">Afslut</td></tr>
        <% 
        
        idag_365 = dateadd("d", -365, now)
        idag_365 = year(idag_365) &"/"& month(idag_365) & "/" & day(idag_365)

        strSQLtodo = "SELECT t.navn, t.id, t.parent, t.dato, t.level, t.delt, "_
	    &" t.afsluttet, t.sortorder, trn.medarbid, trn.todoid, t.tododato, t.forvafsl, t.public, t.parent "_
	    &" FROM todo_rel_new trn "_
	    &" LEFT JOIN todo_new t ON (t.id = trn.todoid AND t.afsluttet <> 1)"_
	    &" WHERE ((medarbid = "& useMrn &" AND trn.todoid = t.id) OR (t.public = 1)) AND afsluttet = 0 AND t.dato > '"& idag_365 &"' GROUP BY t.id ORDER BY t.dato DESC, t.id DESC LIMIT 20"

        'AND parent = 0
        'Response.Write strSQLtodo

        oRec.open strSQLtodo, oConn, 3
        t = 0
        while not oRec.EOF

        select case right(t, 1)
        case 0,2,4,6,8
        tBg = "#FFFFFF"
        case else
        tBg = "#EFf3ff"
        end select


        if t = 0 AND cdate(oRec("dato")) = cdate(date()) then
        tbg = "#FFFFe1"
        end if

        if len(trim(oRec("navn"))) > 200 then
        todoTxt = left(oRec("navn"), 200) & ".."
        else
        todoTxt = oRec("navn")
        end if

        if len(todoTxt) > 15 then
        len_todoTxt = len(todoTxt)
        left_todoTxt = "<b>"& left(todoTxt, 15) & "</b>"
        right_todoTxt = right(todoTxt, len_todoTxt- 15)
        
        todoTxt = left_todoTxt & right_todoTxt 

        else

        todoTxt = "<b>"& todoTxt &"</b>"

        end if

        if oRec("public") = 0 then
        privatTxt = ""
        else
        privatTxt = "(off.lig)"
        end if

        if oRec("forvafsl") = 1 then
        todoDatoTxt = "Forv. udført: " & day(oRec("tododato")) & ". "& left(monthname(month(oRec("tododato"))), 3)&" "& year(oRec("tododato"))
        else
        todoDatoTxt = ""
        end if

       %><tr bgcolor="<%=tBg%>">
       <td class=lille valign=top style="border-top:1px #cccccc solid; padding:4px 5px 4px 5px;">
       <%if oRec("parent") <> 0 then 
       
       strSQLpar = "SELECT navn FROM todo_new WHERE id = "& oRec("parent")
       oRec2.open strSQLpar, oConn, 3
       if not oRec2.EOF then%>
       <%=left(oRec2("navn"), 20) %> >> 
       <%end if
       oRec2.close %>


       <%end if %>
       
       <%=todoTxt%> - <span style="color:#999999; font-size:9px;"><%=day(oRec("dato")) & ". "& left(monthname(month(oRec("dato"))), 3)&" "& year(oRec("dato"))%></span>
       <span style="color:darkred; font-size:9px;"><%=todoDatoTxt %></span></td>
       <td align=right valign=top style="border-top:1px #cccccc solid; padding:2px 5px 2px 5px;">
       <%if oRec("public") <> 1 then %>
       <input id="todoid_<%=oRec("id")%>" class="todoid" type="checkbox" value="1" />
       <%else %>
       <span style="color:#5582d2; font-size:9px;"><%=privatTxt %></span>
       <%end if %></td>
           
       </tr>
       
       <%
        t = t + 1
        oRec.movenext
        wend
        oRec.close
        %>
        <tr>
         <td valign=top colspan=2 style="border-top:1px #cccccc solid; padding:4px 5px 4px 5px;">&nbsp;</td>
        </tr>
        </table>

       
        </div>


	<%
    '****************** Slut højreDIV ***********************'

    end if '**vis ikke højre menu for

    end if   'if browstype_client = "ip" then




end function






function notificerEmail(usemrn, EmailNotificerTxt, visning, modtagerid)


        
                    
                                    


                                    'Sender notifikations mail
		                             if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\timereg_akt_2006.asp" then
        

                                        EnigaFooterTxt = ""
                                        regiAfdeling = ""


                                        '*** Henter afsender **
				                        strSQL = "SELECT mnavn, email, mcpr FROM medarbejdere"_
				                        &" WHERE mid = "& usemrn
				                        oRec5.open strSQL, oConn, 3
            				
				                        if not oRec5.EOF then
            				
				                        afsNavn = oRec5("mnavn")
				                        afsEmail = oRec5("email")
                                        afsCpr = oRec5("mcpr")


                                        regiNavn = afsNavn
				                        regiEmail = afsEmail
                                        regiCpr = afsCpr
                                        regiUserId = usemrn

                                                        if lto = "esn" then   
                                                            call afdelinger(regiUserId) 
                                                        end if
            				
				                        end if
				                        oRec5.close





                                    'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            'Mailer.CharSet = 2
		                            'Mailer.FromName = "TimeOut Email Service (" & afsNavn &" "& afsEmail &") "
                                
                                    '** Hvis problemer med mail pga spam filtre, kan der ændrs her så der bliver sendet fra en TO eller anden adresse der er godkendt af spamfilter
		                            'Mailer.FromAddress = "timeout_no_reply@outzource.dk"
                                     
                                    'Mailer.RemoteHost = "webout.smtp.nu" '"195.242.131.254" ' = smtp3.atznet.dk '"webout.smtp.nu" ' '"webmail.abusiness.dk" '"pasmtp.tele.dk"
                                    'Mailer.ContentType = "text/html"



                                    Set myMail=CreateObject("CDO.Message")
                                    myMail.From="timeout_no_reply@outzource.dk" 'TimeOut Email Service 

                                
                                  

                                        '*********************************************************************************************************
                                        select case cint(visning)
                                        case 1
                                        '* Henter alle de teamleder der abonerer på notifikation for denne projektgruppe 
                                        '*********************************************************************************************************
                                        myMailTostr = ""                            

                                                    call medariprogrpFn(usemrn)
                
                                                    minePrgArr = split(medariprogrpTxt, ",#") 

                                                    for p = 0 to UBOUND(minePrgArr)

                                                    minePrgArr(p) = replace(minePrgArr(p), "#", "")
                                                    minePrgArr(p) = replace(minePrgArr(p), " ", "")
            
                                                        if minePrgArr(p) <> 10 then
                                                        call fnotificer(minePrgArr(p))

                    
                                                                erNotificerArr = split(erNotificer, ",#")

                                                                for n = 0 TO UBOUND(erNotificerArr)

                                                                erNotificerArr(n) = replace(erNotificerArr(n), "#", "")
                                                                erNotificerArr(n) = replace(erNotificerArr(n), " ", "")

                                                                if len(trim(erNotificerArr(n))) <> 0 AND erNotificerArr(n) <> 0 then

                                                                '*** Henter modtager **
				                                                strSQL = "SELECT mnavn, email FROM medarbejdere"_
				                                                &" WHERE mid = "& erNotificerArr(n)
				                                                oRec5.open strSQL, oConn, 3
            				
				                                                if not oRec5.EOF then
            				
				                                                modNavn = oRec5("mnavn")
				                                                modEmail = oRec5("email")
            				
				                                                end if
				                                                oRec5.close

		            				                              
                                                                   myMailTostr = myMailTostr & modNavn &"<"& modEmail &">;"
                                                     
            
                                                               end if
                                                
                                                                next


                                                        end if
                                                    next


                                                        myMail.To = myMailTostr '""& modNavn &"<"& modEmail &">"
                                                        modEmail = myMailTostr 'For at være sikker på mail bliver sendt

                                                        'if lto = "esn" then
                                                        '            myMail.Cc= "Eniga Drift - TimeOut<timeoutfravaer@eniga.dk>"
                                                        'end if


            			                                'Mailer.Subject = "Til Teamleder for: "& afsNavn  
                                                        myMail.Subject= "Til Teamleder for: "& afsNavn  
                                    
		                                                strBody = "<br><br>"
                                                        strBody = strBody & "Følgende medarbejder har registreret en ferie eller fraværs aktivitet som Du skal have besked om:<br><br> "
                                                        strBody = strBody & EmailNotificerTxt
						           
                                                        strBody = strBody & "<br><br><br><br>Du modtager denne mail da Du står som Notifikations modtager for ovenstående medarbejder."

                                                    
                                                                'if lto = "esn" then
                                                                '     
                                                                '    call EnigaFooter
                                                                '    strBody = strBody & EnigaFooterTxt
                                                                '                       
                                                                'end if			

                                                        strBody = strBody &"<br><br><br><br>Med venlig hilsen<br><i>" 
            
                                                        if cstr(trim((afsNavn))) <> cstr(trim((session("user")))) then
		                                                strBody = strBody & session("user") & " på vegne af "& afsNavn &"</i>"
                                                        else
                                                        strBody = strBody & afsNavn &"</i>"
                                                        end if

                                                         strBody = strBody & "<br><br>Denne mail ar afsendt af TimeOut notifikations service<br><br>&nbsp;"
                            




                                        case 2 '2 Leder har indtastet sygdom for medarbejder



                                                       '*** Henter modtager **
                                                    
				                                        strSQL = "SELECT mnavn, email, mcpr FROM medarbejdere"_
				                                        &" WHERE mid = "& modtagerid
				                                    
                                                        oRec5.open strSQL, oConn, 3
                                                        if not oRec5.EOF then
            				
				                                        modNavn = oRec5("mnavn")
				                                        modEmail = oRec5("email")
                                                        modCpr = oRec5("mcpr")
                    
                                                        regiNavn = modNavn
				                                        regiEmail = modEmail
                                                        regiCpr = modCpr
                                                        regiUserId = modtagerid


                                                                    if lto = "esn" then   
                                                                        call afdelinger(regiUserId) 
                                                                    end if
            				
				                                        end if
				                                        oRec5.close

		            				                    'Mailer.AddRecipient "" & modNavn & "", "" & modEmail & ""
                                                        myMail.To= ""& modNavn &"<"& modEmail &">"
                
                                                        'if lto = "esn" then
                                                              'myMail.Cc= "Eniga Drift - TimeOut<timeoutfravaer@eniga.dk>"
                                                        'end if



                                                'Mailer.Subject = "Din leder har indtastet fravær på dine vegne"
                                                myMail.Subject= "Din leder har indtastet fravær på dine vegne"
                                    
		                                        strBody = "<br><br>"
                                                strBody = strBody & "Din leder "& session("user") &" har indtastet følgende fravær for dig:<br><br> "
                                                strBody = strBody & EmailNotificerTxt
		        
                                                                
                                                        'if lto = "esn" then
                                                         
                                                        '     call EnigaFooter
                                                        '    strBody = strBody & EnigaFooterTxt
                                                                           
                                                        'end if				           


                                                strBody = strBody &"<br><br><br><br>Med venlig hilsen<br><i>" 
                                                strBody = strBody & session("user") & "</i>"
                                                strBody = strBody & "<br><br>Denne mail ar afsendt af TimeOut notifikations service<br><br>&nbsp;"
                                  



                                  case 3 '3 Eniga mail



                                                '*** Henter modtager **
                                                    
				                                strSQL = "SELECT mnavn, email, mcpr FROM medarbejdere"_
				                                &" WHERE mid = "& modtagerid
				                                    
                                                oRec5.open strSQL, oConn, 3
                                                if not oRec5.EOF then
            				
				                                modNavn = oRec5("mnavn")
				                                modEmail = oRec5("email")
                                                modCpr = oRec5("mcpr")
                    
                                                regiNavn = modNavn
				                                regiEmail = modEmail
                                                regiCpr = modCpr
                                                regiUserId = modtagerid


                                                          
            				
				                                end if
				                                oRec5.close

		            				          
                                                myMail.To= "Eniga Drift - TimeOut<timeoutfravaer@eniga.dk>"
                                               

                                                if lto = "esn" then
                                                myMail.Cc= "Eniga Drift DORTE TEST - <DFN@esnord.dk>"
                                                myMail.Bcc = "SK TEST ENIGA<sk@outzource.dk>"
                                                end if



                                                'Mailer.Subject = "Din leder har indtastet fravær på dine vegne"
                                                myMail.Subject= "Eniga Løn - Der er indtastet fravær i Timeout (ESNord)"
                                    
		                                        strBody = "<br><br>"
                                                strBody = strBody & "Hej Eniga løn,<br>Der er indtastet følgende fravær i TimeOut:<br><br> "
                                                strBody = strBody & EmailNotificerTxt3
		        
                                                                
                                                        if lto = "esn" then
                                                         
                                                             call EnigaFooter
                                                            strBody = strBody & EnigaFooterTxt
                                                                           
                                                        end if				           


                                                strBody = strBody &"<br><br><br><br>Med venlig hilsen<br><i>" 
                                                strBody = strBody & session("user") & "</i>"
                                                strBody = strBody & "<br><br>Denne mail ar afsendt af TimeOut notifikations service<br><br>&nbsp;"
                                  




                                    end select 'visning


                                 
		                            'Mailer.BodyText = strBody
                                    myMail.HTMLBody= "<html><head></head><body>" & strBody & "</body>"

                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                                    'Name or IP of remote SMTP server
                                    
                                                if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                                   smtpServer = "webout.smtp.nu"
                                                else
                                                   smtpServer = "formrelay.rackhosting.com" 
                                                end if
                    
                                                myMail.Configuration.Fields.Item _
                                                ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                                    'Server port
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                                    myMail.Configuration.Fields.Update

                                    if len(trim(modEmail)) <> 0 then
                                    myMail.Send
                                    end if

                                    set myMail=nothing

            		
		                            'Mailer.sendmail()
		                            'Set Mailer = Nothing

                              
                                    end if


end function

public EnigaFooterTxt, regiNavn, regiCpr
sub Enigafooter


        EnigaFooterTxt = EnigaFooterTxt &"<br><br><br>Partner: ESNord"
        EnigaFooterTxt = EnigaFooterTxt &"<br>Navn: "& regiNavn
        EnigaFooterTxt = EnigaFooterTxt &"<br>Cpr. nr.: "& regiCpr  
        EnigaFooterTxt = EnigaFooterTxt &"<br>Afdeling: "&  regiAfdeling 
        EnigaFooterTxt = EnigaFooterTxt &"<br>Årsag: Frav&aelig;r/Sygdom"

end sub

public regiAfdeling
function afdelinger(modtagerid)


      
                    strSQLafd = "SELECT pr.projektgruppeid, p.navn AS afdeling FROM progrupperelationer AS pr "_
                    &" LEFT JOIN projektgrupper AS p ON (p.id = pr.projektgruppeid) WHERE pr.medarbejderid = " & modtagerid & " AND p.id <> 10"
                                                                        
                    af = 0
                    oRec6.open strSQLafd, oConn, 3
                    while not oRec6.EOF 
                    if af = 0 then
                    regiAfdeling = regiAfdeling & oRec6("afdeling")
                    else
                    regiAfdeling = regiAfdeling &", "& oRec6("afdeling")
                    end if
                    af = af + 1
                    oRec6.movenext
                    wend
				    oRec6.close

             


end function



%>





