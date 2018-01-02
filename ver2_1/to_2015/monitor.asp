<%
    session("lto") = "9K2017-1124-TO179"
%>


<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
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


            if func = "scan" then        
            
                RFID = request("RFID_field")


                '******* Tjekker medarbejderen ********'
                if RFID <> 0 then
                    strSQL = "SELECT Mid, Mnavn FROM medarbejdere WHERE medarbejder_rfid = "& RFID
                    response.Write strSQL
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
                    strSQL = "SELECT id, logud FROM login_historik WHERE mid = "& medid &" ORDER BY id DESC"
                    oRec.open strSQL, oConn, 3
                    if not oRec.EOF then
                    logud = oRec("logud")
                    id = oRec("id")
                    end if
                    oRec.close
                    response.Write logud & id
                    if isNull(logud) = true then
                    response.Redirect "ugeseddel_2011.asp?medarb_navn="&medarb_navn&"&medid="&medid&"&touchMonitor=1&varTjDatoUS_man="&year(now)&"/"&month(now)&"/"&day(now)
                    else
                    response.Redirect "monitor.asp?func=login_valg&medarb_navn="&medarb_navn&"&medid="&medid
                    end if 
                else
                response.Redirect "monitor.asp?func=startside"
                end if

            end if



            if func = "login" then

                'medid = request("medid")
                'medarb_navn = request("medarb_navn")

                loginTid = Year(now) &"-"& Month(now) &"-"& Day(now) & " " & Time
                strSQLlogin = "Insert Into login_historik SET dato = '"& Year(now) & "-" & Month(now) & "-" & Day(now) &"', login = '"& loginTid &"', mid ="& medid
                'response.Write strSQLlogin
                oConn.execute(strSQLlogin)
           
            
                response.Redirect "monitor.asp?func=startside&medarb_navn="&medarb_navn&"&medid="&medid&"&redType=1"

            end if



            if func = "logud" then

                'medid = request("medid")
                'medarb_navn = request("medarb_navn")
                logudtid = Year(now) &"-"& Month(now) &"-"& Day(Now) &" "& Time

                strSQLId = "SELECT id, logud FROM login_historik WHERE mid = "& medid &" ORDER BY id DESC"
                oRec.open strSQLId, oConn, 3
                if not oRec.EOF then
                lastid = oRec("id")
                logud = oRec("logud")
                end if
                oRec.close
                response.Write "HERHER" & lastid

                if isNull(logud) = true then 

                    strSQL = "Update login_historik SET logud = '"& logudtid &"', stempelurindstilling = 1 WHERE id = "& lastid
                    oConn.execute(strSQL)
                    response.Redirect "monitor.asp?func=startside&redtype=2&medarb_navn="&medarb_navn
            
                end if
                

            end if

        %>

        <div class="container" style="width:100%; height:100%">
            <div class="portlet">
                <div class="portlet-body">
                    <script src="js/monitor_jav3.js" type="text/javascript"></script>
                    <%
                    select case func
                                       


                        case "startside"

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
                                            <td><input type="password" class="form-control input-small" id="RFID_field" name="RFID_field" placeholder="Scan kort" autofocus /></td>
                                        </tr>
                                    </table>
                                </div>
                            </form>
                            
                            <br /><br />
                            <%
                            if redType = 1 then
                                'medarb_navn = request("medarb_navn")
                            %>
                                <div id="text_besked" class="row">        
                                    <h3 class="col-lg-12" style="text-align:center">Velkommen <%=medarb_navn %></h3>
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
                        case "login_valg"

                        'medarb_navn = request("medarb_navn")
                        'medid = request("medid")
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
                                    <td><a href="monitor.asp?func=login&medarb_navn=<%=medarb_navn %>&medid=<%=medid %>" class="btn btn-success btn-lg"><b>Login</b></a></td>
                                </tr>
                            </table>

                            <table style="display:inline-table">
                                <tr>
                                    <td><a class="btn btn-success btn-lg"><b>Overtid</b></a></td>
                                </tr>
                            </table>
                        </div>
                    
                    <%end select %>
                </div>
            </div>
        </div>




    

<!--#include file="../inc/regular/footer_inc.asp"-->