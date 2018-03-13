<%
    'session("lto") = "xxx"
    'session("lto") = "intranet - local"
    session("lto") = "9K2017-1124-TO179"
    'lto = "intranet - local"
    lto = "cflow"
%>


<!--#include file="../inc/connection/conn_db_inc.asp"-->

<%



 '**** Søgekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_gemjobc"

                jq_jobidc = request("jq_jobidc")
                 'response.Cookies("monitor_job")(session("mid")) = jq_jobidc

                strSQLdel = "DELETE FROM login_historik_aktivejob_rel WHERE lha_mid = "& session("mid") & " AND lha_jobid <> 0"
                oConn.execute(strSQLdel)

                strSQLins = "INSERT INTO login_historik_aktivejob_rel (lha_mid, lha_jobid)  VALUES ("& session("mid") &", "& jq_jobidc &")" 
                'response.Write strSQLins
                'response.end
    
                oConn.execute(strSQLins)


        case "FN_gemaktc"

                jq_aktidc = request("jq_aktidc")

                'response.Cookies("monitor_akt")(session("mid")) = jq_aktidc
                strSQLdel = "DELETE FROM login_historik_aktivejob_rel WHERE lha_mid = "& session("mid") & " AND jq_aktidc <> 0"
                oConn.execute(strSQLdel)

                strSQLins = "INSERT INTO login_historik_aktivejob_rel (lha_mid, lha_aktid) VALUES ("& session("mid") &", "& jq_aktidc &")"
                oConn.execute(strSQLins)


        case "FN_sogjobogkunde"


                

                call selectJobogKunde_jq()


        
          case "FN_sogakt"

               
                call selectAkt_jq()

        end select
        response.end
        end if    
    
%>





<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->





       
        
        <%
           

            'if len(session("user")) = 0 then

	        'errortype = 5
	        'call showError(errortype)
               'response.End
	        'end if
            
            'lto = "cflow"

            func = request("func")
            medid = request("medid")
            medarb_navn = request("medarb_navn")
            stempelurindstilling = request("stempelurindstilling")

            if len(trim(request("skiftjob"))) <> 0 then
            skiftjob = request("skiftjob")
            else
            skiftjob = 0
            end if

            select case func
            case "scan" 
            
                RFID = request("RFID_field")


                '******* Tjekker medarbejderen ********'
                if len(trim(RFID)) <> 0 then
                    strSQL = "SELECT Mid, Mnavn FROM medarbejdere WHERE (medarbejder_rfid = "& RFID & " OR mnr = "& RFID & ") AND (mansat = 1 OR mansat = 3)"
                    'response.Write strSQL
                    oRec.open strSQL, oConn, 3
                    if not oRec.EOF then
                    medarb_navn = oRec("Mnavn")
                    medid = oRec("Mid")  
                    session("login") = Trim(oRec("Mid"))
	                session("mid") = Trim(oRec("Mid"))
                    session("user") = oRec("Mnavn")
                    'session("rettigheder") = oRec("rettigheder")
                    else
                    response.Redirect "monitor.asp?func=startside&redType=3"

                    end if
                    oRec.close

                    '******* Tjekker logindtid ********'
                    'strSQL = "SELECT id, logud, login FROM login_historik WHERE mid = "& medid &" AND stempelurindstilling = 1 ORDER BY id DESC"
                    'oRec.open strSQL, oConn, 3
                    'if not oRec.EOF then
                    'logud = oRec("logud")
                    'login = oRec("login")
                    'id = oRec("id")
                    'end if
                    'oRec.close
                    'response.Write logud & id
                    
                    fo_oprettetnytlogin = 0

                    strUsrId = medid
                    intStempelur = 1 'request("stempelurindstilling")
                    io = 1 '= Afslutter hvis åbent mere end 20 timer
                    tid = now

                    call logindStatus(strUsrId, intStempelur, io, tid)

                    'if medid = 1 then
                    'Response.Write "medid: "& medid &" login: "& login &" logud: "& logud &" = "& isNull(logud) & " fo_afsluttetlogin: "& fo_afsluttetlogin &" AND fo_oprettetnytlogin: "& cint(fo_oprettetnytlogin) & " fo_autoafsluttet: "& fo_autoafsluttet & " skiftjob: " & skiftjob
                    'response.end
                    'end if


                    'if isNull(logud) = true then 'AND cint(fo_oprettetnytlogin) = 0 AND cint(fo_autoafsluttet) = 0 then
                    'fo_oprettetnytlogin = "1"

                    if skiftjob = "0" then

                            if fo_oprettetnytlogin = "1" then
                            response.Redirect "monitor.asp?func=startside&redType=1"
                            else
                            response.Redirect "monitor.asp?func=registrer" 
                            end if 

                    else

                          
                            response.Redirect "monitor.asp?func=registrer" 
                            
                    end if


                else
                response.Redirect "monitor.asp?func=startside"
                end if

            



            case "xxlogin" 

                'medid = session("mid") 'request("medid")
                'medarb_navn = request("medarb_navn")

                 'loginTid = Year(now) &"-"& Month(now) &"-"& Day(now) & " " & Time
                
                    strUsrId = medid
                    intStempelur = 1 'request("stempelurindstilling")
                    io = 1
                    tid = now

                    'call logindStatus(strUsrId, intStempelur, io, tid)
                    
                    'Response.write "fo_logud" & fo_logud
                    'response.end

                'strSQLlogin = "Insert Into login_historik SET dato = '"& Year(now) & "-" & Month(now) & "-" & Day(now) &"', "_
                '&" login = '"& loginTid &"', login_first = '"& loginTid &"', mid ="& medid &", stempelurindstilling = "& stempelurindstilling &"" 
                'response.Write strSQLlogin
                'oConn.execute(strSQLlogin)
           
            
                'response.Redirect "monitor.asp?func=startside&redType=1" '&medarb_navn="&medarb_navn&"&medid="&medid

         


            case "xlogud"
                    
                                '** Dencker
                                'call smileyAfslutSettings()  

                                'strUsrId = session("mid")

                                'if cint(SmiWeekOrMonth) = 2 then 'videre til ugeafslutning og ugeseddel
                                'Response.redirect "../timereg/stempelur.asp?func=afslut&medarbSel="& strUsrId &"&showonlyone=1&hidemenu=1&id=0&rdir=sesaba" '../to_2015/ugeseddel_2011.asp?nomenu=1
                                'else
                                'Response.redirect "../sesaba.asp?logudDone=1"
                                'Response.end
                                'end if            

                'response.Redirect "../sesaba.asp"




                'medid = session("mid") 'request("medid")
                'medarb_navn = session("user") 'request("medarb_navn")
                
                ' logudtid = Year(now) &"-"& Month(now) &"-"& Day(Now) &" "& Time

                'strSQLId = "SELECT id, logud FROM login_historik WHERE mid = "& medid &" ORDER BY id DESC"
                'oRec.open strSQLId, oConn, 3
                'if not oRec.EOF then
                'lastid = oRec("id")
                'logud = oRec("logud")
                'end if
                'oRec.close
                
               
                'if isNull(logud) = true then 

                '    strSQL = "Update login_historik SET logud = '"& logudtid &"' WHERE id = "& lastid
                '    oConn.execute(strSQL)
                '    response.Redirect "monitor.asp?func=startside&redtype=2&medarb_navn="&medarb_navn
            
                'end if
                

            case "startside", "xlogin_valg", "registrer"

           

        %>

        <div class="container" style="width:100%; height:100%">
            <div class="portlet">
                <div class="portlet-body">
                    <script src="js/monitor_jav6.js" type="text/javascript"></script>
                    <%
                    if func = "startside" then
                                       

                        redType = request("redType")


                        %>
                            <div class="row">
                                <div class="col-lg-12" style="text-align:center"><img src="img/cflow_logo.png" /></div>
                            </div>
                        
                    
                            
                            <br />
                            <form action="monitor.asp?func=scan" id="scan" method="post">
                                <div style="text-align:center">
                                    <table style="display:inline-table">
                                        <tr>
                                            <td><input type="number" class="form-control input-small" id="RFID_field" name="RFID_field" placeholder="Scan kort" autofocus /></td>
                                        </tr>
                                    </table>
                                </div>
                                <div class="row">
                              
                                <div class="col-lg-12" style="text-align:right;">
                                    <br /><br /><br />
                                    <input type="submit" value="Logind >> " class="btn btn-default" /></div>
                            </div>
                                
                                
                            </form>
                           
                            
                          
                            <%
                            if redType = 1 then
                                'medarb_navn = request("medarb_navn")

                                if len(trim(session("mid"))) <> 0 then
                                medidloggetInd = session("mid")
                                else
                                medidloggetInd = 0
                                end if

                                call meStamdata(medidloggetInd)
                            %>
                                <div id="text_besked" class="row">        
                                    <h3 class="col-lg-12" style="text-align:center">Velkommen <%=meNavn %><br /><span style="font-size:14px;">Du er logget ind <%=formatdatetime(now, 1) & " " & formatdatetime(now, 3)  %></span>


                                        <%
                                            jobidC = "-1"
                                            aktidC = "-1"
                                            'end if

                                            '** IKKE PÅ COOKIE DA DET SKAL GÆLDE ALLE TERMINALER
                                            strSQLsel = "SELECT lha_jobid FROM login_historik_aktivejob_rel WHERE lha_mid = "& medidloggetInd & " AND lha_jobid <> 0"

                                            'response.write strSQLsel
                                            'response.flush

                                            oRec.open strSQLsel, oConn, 3
                                            if not oRec.EOF then

                                                jobidC = oRec("lha_jobid")

                                            end if
                                            oRec.close
                

                                             '** IKKE PÅ COOKIE DA DET SKAL GÆLDE ALLE TERMINALER
                                            strSQLsel = "SELECT lha_aktid FROM login_historik_aktivejob_rel WHERE lha_mid = "& medidloggetInd & " AND lha_aktid <> 0"
                                            oRec.open strSQLsel, oConn, 3
                                            if not oRec.EOF then

                                                aktidC = oRec("lha_aktid")

                                            end if
                                            oRec.close


                                        jobTxtloggetindpaa = "Du er ikke logget ind på et projekt"
                                        if jobidC <> "-1" AND jobidC <> "" then
                                            
                                            strSQLjob = "SELECT jobnavn, jobnr FROM job WHERE id = " & jobidC
                                            oRec.open strSQLjob, oConn, 3
                                            if not oRec.EOF then

                                            jobTxtloggetindpaa = oRec("jobnr") & " " & oRec("jobnavn")

                                            end if
                                            oRec.close

                                        end if

                                        %>

                                        <br /><br />Projekt: <%=jobTxtloggetindpaa %>

                                    </h3>
                                </div>
                            <%
                            end if
                            %>
                            <%
                            if redType = 2 then
                                'medarb_navn = request("medarb_navn")
                            %>
                                <div id="text_besked" class="row">        
                                    <h3 class="col-lg-12" style="text-align:center">Tak for i dag <%=medarb_navn %></h3>
                                </div>
                            <%
                            end if
                            %>
                             <%
                            if redType = 3 then
                                'medarb_navn = request("medarb_navn")
                            %>
                                <div id="text_besked" class="row">        
                                    <h3 class="col-lg-12" style="text-align:center">Bruger findes ikke</h3>
                                </div>
                            <%
                            end if
                            %>

                        
                    <%
                        end if





                        if func = "login_valg" then

                        medarb_navn = session("user") 'request("medarb_navn")
                        medid = session("mid") 'request("medid")
                    %>
                        <div class="portlet-title">
                          <table style="width:100%;">
                              <tr>
                                  <td><h3 style="color:black"><%=medarb_navn %> <!--(<%=medid %>)--></h3></td>
                                  <td style="text-align:right"><img style="width:15%; margin-bottom:2px; margin-left:55px" src="img/cflow_logo.png" /></td>
                              </tr>
                          </table>
                        </div>

                        <br /><br /><br /><br />
                        <div style="text-align:center">
                            <table style="display:inline-table; margin-right:200px">
                                <tr>
                                    <td><a href="monitor.asp?func=login&medarb_navn=<%=medarb_navn %>&medid=<%=medid %>&stempelurindstilling=1" class="btn btn-success btn-lg"><b>Login</b></a></td>
                                </tr>
                            </table>

                            <table style="display:inline-table">
                                <tr>
                                    <td><a href="monitor.asp?func=login&medarb_navn=<%=medarb_navn %>&medid=<%=medid %>&stempelurindstilling=2" class="btn btn-success btn-lg"><b>Overtid</b></a></td>
                                </tr>
                            </table>
                        </div>
                    
                    <%end if



                    if func = "registrer" then

                        %>
                        <meta http-equiv="refresh" content="120;url=monitor.asp?func=startside" />
                        <script src="js/ugeseddel_2011_jav4.js" type="text/javascript"></script>
                        <!--#include file="inc/touchmonitor_inc.asp"-->

                        <%
                        
                        varTjDatoUS_man = year(now)&"/"&month(now)&"/"&day(now)

                        call mobil_week_reg_dd_fn()

                        medarb_navn = session("user") 'request("medarb_navn")
                        medid = session("mid") 'request("medid")

                        call meStamdata(medid)

                    %>
                          <table style="width:100%;">
                              <tr>
                                  <td valign="top"><h3 style="color:black;"><%=medarb_navn%> (<%=meNr %>)</h3>
                            
                            
                                       <%
                                    '******* Tjekker logindtid ********'
                                    strSQL = "SELECT id, login FROM login_historik WHERE mid = "& medid &" AND stempelurindstilling > 0 ORDER BY id DESC"
                                    oRec.open strSQL, oConn, 3
                                    if not oRec.EOF then
                                    login = oRec("login")
                            
                                    end if
                                    oRec.close
                            
                              
                                  %><br />
                                  <span style="font-size:12px;">Logget på kl. <%=login %></span>

                                  <% 
                                    overtidtot = 0
                                    strSQL6 = "SELECT SUM(timer) As overtidtimer, tdato FROM timer WHERE tmnr = "& medid &" AND tfaktim = 30 GROUP BY tmnr" 'AND tdato BETWEEN '"& sqlDatoStartATD &"' AND '"& sqlDatoEnd &"'
                                    oRec.open strSQL6, oConn, 3
                                    if not oRec.EOF then

                                    overtidtot = overtidtot + oRec("overtidtimer")*1
        
                                    end if 
                                    oRec.close 
                                      


                                  %><br /><span style="font-size:12px;">Timebank/Flexsaldo: - <!--<%=formatnumber(overtidtot, 2) %>--></span>


                                  </td>
                                  <td style="text-align:right" valign="top"><img style="width:15%; margin-bottom:2px; margin-left:55px" src="img/cflow_logo.png" /></td>
                              </tr>
                          </table>
                        
                        
                   

                        <div class="row">
                        <div class="col-lg-3">&nbsp;
                            <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
                            <br /><br /><br /><br /><br /><br /><br /><br />
                              <a href="monitor.asp?func=startside" class="btn btn-lg btn-default"><< <b>Tilbage / Ny bruger</b></a>   

                        </div>
                        <div class="col-lg-3"> 
                            
                            <%
                            touchMonitor = 1
                            call keyfigures(touchMonitor, login)
                              %>
                               
                        </div>
                        <div class="col-lg-3">

                                        <table class="table" style="width:300px;">

                                            <tr>
                                                <th colspan="3">Historik (seneste 14)</th>
                                            </tr>
                                        

                            <%
                                    strSQL = "SELECT id, login, logud FROM login_historik WHERE mid = "& medid &" AND stempelurindstilling > 0 AND logud IS NOT NULL ORDER BY id DESC LIMIT 14"
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF 
                                    login = oRec("login")
                                    logud = oRec("logud")
                            
                                    %>
                                            <tr>
                                                <td><%=left(weekdayname(weekday(login)), 3) &". "& formatdatetime(login, 2)%></td>
                                                <td><%=formatdatetime(login, 3)%></td>
                                                <td>- <%=formatdatetime(logud, 3)%></td>
                                            </tr>

                                    <%

                                    oRec.movenext
                                    wend
                                    oRec.close    
                                
                            %></table>


                        </div>
                         </div><!--row -->

                   
                    
                    <%end if %>


                </div>
            </div>
        </div>
<%end select %>

<!--#include file="../inc/regular/footer_inc.asp"-->