<%
    'session("lto") = "xxx"
    'session("lto") = "intranet - local"
    session("lto") = "9K2017-1124-TO179"
    'lto = "intranet - local"
    lto = "cflow"
%>


<!--#include file="../inc/connection/conn_db_inc.asp"-->
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
                    io = 1 '= Afslutter hvis �bent mere end 20 timer
                    tid = now

                    call logindStatus(strUsrId, intStempelur, io, tid)

                    'if medid = 1 OR medid = 118 then
                    'Response.Write "medid: "& medid &" login: "& login &" logud: "& logud &" = "& isNull(logud) & " fo_afsluttetlogin: "& fo_afsluttetlogin &" AND fo_oprettetnytlogin: "& cint(fo_oprettetnytlogin) & " fo_autoafsluttet: "& fo_autoafsluttet
                    'response.end
                    'end if


                    'if isNull(logud) = true then 'AND cint(fo_oprettetnytlogin) = 0 AND cint(fo_autoafsluttet) = 0 then
                    'fo_oprettetnytlogin = "1"
                    if fo_oprettetnytlogin = "1" then
                    'if cint(fo_logud) <> 3 then
                    response.Redirect "monitor.asp?func=startside&redType=1"

                    'Response.write "HER: monitor.asp?func=startside&redType=1"
                    'response.end

                    else
                    'response.Redirect "monitor.asp?func=login_valg&medarb_navn="&medarb_navn&"&medid="&medid
                    'response.Redirect  "monitor.asp?func=login&medarb_navn="&medarb_navn&"&medid="&medid&"&stempelurindstilling=1"
                    'Response.write "HER 2:ugeseddel_2011.asp"
                    'response.end
                    'response.Redirect "monitor.asp?func=startside&redType=1"
                    response.Redirect "ugeseddel_2011.asp?lto="&lto&"medarb_navn="&medarb_navn&"&medid="&medid&"&usemrn="&medid&"&touchMonitor=1&varTjDatoUS_man="&year(now)&"/"&month(now)&"/"&day(now)
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
                

            case "startside", "xlogin_valg"

           

        %>

        <div class="container" style="width:100%; height:100%">
            <div class="portlet">
                <div class="portlet-body">
                    <script src="js/monitor_jav7.js" type="text/javascript"></script>
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
                                            <td><input type="text" class="form-control input-small" id="RFID_field" name="RFID_field" placeholder="Scan kort" autofocus /></td>
                                        </tr>
                                        <tr>
                                            <td style="padding-top:20px;"><a class="btn btn-default submit">Logind</a></td>
                                        </tr>
                                    </table>
                                </div>

                              <!--  <div class="row">
                                <div class="col-lg-12" style="text-align:right;">
                                    <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
                                    <input type="submit" value="Logind >> " class="btn btn-default" /></div>
                            </div> -->
                                
                                
                            </form>
                            
                            <br /><br />
                            <%
                            if redType = 1 then
                                'medarb_navn = request("medarb_navn")
                            %>
                                <div id="text_besked" class="row">        
                                    <h3 class="col-lg-12" style="text-align:center">Velkommen <%=medarb_navn %><br /><span style="font-size:11px;">Du er logget ind <%=formatdatetime(now, 1) & " " & formatdatetime(now, 3)  %></span></h3>
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
                    
                    <%end if %>
                </div>
            </div>
        </div>
<%end select %>

<!--#include file="../inc/regular/footer_inc.asp"-->