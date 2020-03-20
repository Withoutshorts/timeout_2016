<%


    'if len(trim(request("fromlogin"))) <> 0 then

    ' session("lto") = request("key")
    ' lto = request("lto")

    'else

    'session("lto") = "xxx"
    'session("lto") = "intranet - local"
    '        if len(trim(session("lto"))) = "" then

   '         if session("lto") = "9K2018-2506-TO181" then
    '
    '        session("lto") = "9K2018-2506-TO181"
    '        lto = "miele"
    '
    '          else
            
            ltokey = request("ltokey")

            if ltokey <> "" AND len(trim(ltokey)) <> 0 then
                session("lto") = ltokey
            else
                
                if session("lto") = "" then
                    response.Write monitor_txt_001
                    response.End
                end if

            end if

            

            'lto = "intranet - local"
            
    '        else
    '        session("lto") = session("lto")
    '        lto = lto
    '        end if
    '        end if

    'end if

   
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
                'strSQLdel = "DELETE FROM login_historik_aktivejob_rel WHERE lha_mid = "& session("mid") & " AND jq_aktidc <> 0"
                'oConn.execute(strSQLdel)

                strSQLins = "UPDATE login_historik_aktivejob_rel SET lha_aktid = "& jq_aktidc &" WHERE lha_mid = "& session("mid")
                oConn.execute(strSQLins)


        case "FN_sogjobogkunde"


                

                call selectJobogKunde_jq()


        
        case "FN_sogakt"

               
                call selectAkt_jq()


        case "FindGuestCompany"
                guestval = request("guestval")
                stropt = ""

                strSQL = "SELECT firma FROM medarbejdere WHERE firma LIKE '%"& guestval &"%' AND mansat = 4 GROUP BY firma"
                oRec.open strSQL, oConn, 3
                i = 0
                while not oRec.EOF
                    stropt = stropt & "<option value='"&oRec("firma")&"'>"& oRec("firma") &"</option>"   
                i = i + 1
                oRec.movenext
                wend
                oRec.close

                call jq_format(stropt)
                stropt = jq_formatTxt

                response.Write stropt

        case "FindGuestEmployee"

                guestval = request("guestval")
                guestCompany = request("guestCompany")
                stropt = ""

                strSQL = "SELECT mid, mnavn, mtlf FROM medarbejdere WHERE mnavn LIKE '%"& guestval &"%' AND firma LIKE '%"& guestCompany &"%' AND mansat = 4"
                oRec.open strSQL, oConn, 3
                i = 0
                while not oRec.EOF
                    stropt = stropt & "<option value='"&oRec("mnavn")&"' data-tlf='"&oRec("mtlf")&"'>"& oRec("mnavn") &"</option>"   
                i = i + 1
                oRec.movenext
                wend
                oRec.close

                call jq_format(stropt)
                stropt = jq_formatTxt

                response.Write stropt

           case "FindEmployee" 
                inputName = request("besogInput")
                        
                strOpt = ""
                strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE mnavn LIKE '%"& inputName &"%' AND mansat = 1"
                oRec.open strSQL, oConn, 3
                while not oRec.EOF
                    strOpt = strOpt & "<option value='"&oRec("mid")&"' data-name='"&oRec("mnavn")&"'>"& oRec("mnavn") &"</option>"  
                oRec.movenext
                wend
                oRec.close

                call jq_format(strOpt)
                strOpt = jq_formatTxt

                response.Write strOpt


            case "GetVisitsTel"

                medid = request("medid")
                strtlf = ""
                strSQL = "SELECT mtlf FROM medarbejdere WHERE mid = "& medid
                oRec.open strSQL, oConn, 3
                if not oRec.EOF then
                    strtlf = oRec("mtlf")
                end if
                oRec.close

                strtlf = replace(strtlf, "+", "")
                strtlf = replace(strtlf, " ", "")

                response.Write strtlf
    

        end select
        response.end
        end if    
    
%>





<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->





       
        
        <%          
                   ' response.End

            'if len(session("user")) = 0 then

	        'errortype = 5
	        'call showError(errortype)
               'response.End
	        'end if
            
            'lto = "cflow"

            func = request("func")

            if len(trim(request("vislogud"))) <> 0 then
                vislogud = request("vislogud")
            else
                vislogud = 0
            end if

            if func <> "registrer" AND vislogud <> 1 then
                session("login") = ""
	            session("mid") = ""
                session("user") = ""
            end if

            select case lto
                case "outz", "cool", "cflow"
                    useGuestLogin = 1
                case else
                    useGuestLogin = 0
            end select

            select case lto
                case "cflow"
                    useJobAktPicker = 1
                case else
                    useJobAktPicker = 0
            end select

            medid = request("medid")
            medarb_navn = request("medarb_navn")
            stempelurindstilling = request("stempelurindstilling")

            if len(trim(request("skiftjob"))) <> 0 then
            skiftjob = request("skiftjob")
            else
            skiftjob = 0
            end if

            select case func
                case "meeting"
                    medid = request("medid")
                    fradato = year(now) &"-"& month(now) &"-"& day(now)
                    fratidspunkt = hour(now) &":"& minute(now)
                    heledagen = 0
                    tildato = year(now) &"-"& month(now) &"-"& day(now)
                
                    fromDateTim = fradato &" "& fratidspunkt        

                    strSQL = "INSERT INTO udeafhuset (medid, fradato, fratidspunkt, tildato, heledagen, fra) VALUES ("& medid &", '"& fradato &"', '"& fratidspunkt &"', '"& tildato &"', "& heledagen &", '"& fromDateTim &"')"
                    'response.Write strSQL
                    oConn.execute(strSQL)
                    
                    
                    'response.End
                    response.Redirect "monitor.asp?func=startside&redType=5" 
                    

                case "guestlogud"

                    guestid = request("guestid")
           
                    call meStamdata(guestid)

                    %>
                    <div class="container">
                        <div class="portlet">
                            <div class="portlet-body">
                                <div class="row" style="text-align:center;">
                                    <h3 class="col-lg-12"><%=monitor_txt_002 & " " %><%=menavn & " " %><%=monitor_txt_003 %></h3>
                                </div>

                                <br />

                                <div class="row">
                                    <div class="col-lg-12" style="text-align:center;">
                                        <a href="../sesaba.asp?frommonitor=1&guestlogud=1&medid=<%=guestid %>" class="btn btn-lg btn-danger"><%=monitor_txt_004 %></a>
                                        <a href="monitor.asp?func=startside" class="btn btn-lg btn-default"><%=monitor_txt_005 %></a>
                                    </div>
                                </div>

                            </div>
                        </div>
                    </div>
                    <%

            case "guestlogin"

                guestEmployee = request("guestEmployee")
                guestCompany = request("guestCompany")
                guestTlf = replace(request("guestTlf"), "'", "''")
                visit = request("visit")

                if len(trim(guestCompany)) <> 0 AND len(trim(guestEmployee)) <> 0 then
            
                    ' Tjekker om gæstemedarbejder findes i forvejen
                    guestFindes = 0
                    strSQLFindesGuest = "SELECT mid, mnavn, mtlf FROM medarbejdere WHERE mnavn = '"& guestEmployee &"' AND firma = '"& guestCompany &"' AND mansat = 4"
                    oRec.open strSQLFindesGuest, oConn, 3
                    if not oRec.EOF then
                        session("login") = Trim(oRec("Mid"))
	                    session("mid") = Trim(oRec("Mid"))
                        session("user") = oRec("Mnavn")

                        response.Cookies("2015")("lastlogin") = oRec("Mid")

                        '* Opdater telefon nummer til hvis det ikke blev indtastet første gang eller det er blevet ændret
                        if len(trim(guestTlf)) <> 0 then
                            oConn.execute("UPDATE medarbejdere SET mtlf = '"& guestTlf &"' WHERE mid = "& oRec("mid"))
                        end if

                        guestFindes = 1

                        'Sletter gæstens frokost ordning fra sidste gang
                        strDEL = "DELETE FROM progrupperelationer WHERE projektgruppeid = 12 AND medarbejderid = "& Trim(oRec("Mid"))
                        oConn.execute(strDEL)


                    end if
                    oRec.close

                    'Hvis gæstemedarbejder ikke findes, oprettes den
                    if guestFindes = 0 then
                        ansatdato = year(now) &"-"& month(now) &"-"& day(now)
                        sqlOpr = "INSERT INTO medarbejdere SET mnavn = '"&guestEmployee&"', firma = '"& guestCompany &"', mansat = 4, ansatdato = '"& ansatdato &"', mtlf = '"& guestTlf &"', brugergruppe = 4"
                        oConn.execute(sqlOpr)
                    
                            
                        ' Henter den nyoprettede gæst
                        oRec.open strSQLFindesGuest, oConn, 3
                        if not oRec.EOF then
                            session("login") = Trim(oRec("Mid"))
	                        session("mid") = Trim(oRec("Mid"))
                            session("user") = oRec("Mnavn")

                            response.Cookies("2015")("lastlogin") = oRec("Mid")

                            guestFindes = 1
                        end if
                        oRec.close

                    end if


                    fo_oprettetnytlogin = 0
                    strUsrId = session("mid")
                    intStempelur = 1 'request("stempelurindstilling")
                    io = 1 '= Afslutter hvis åbent mere end 20 timer
                    tid = now
                    call logindStatus(strUsrId, intStempelur, io, tid)

                    if fo_oprettetnytlogin = 0 then
                        vislogud = 1
                        'message = 0

                        'response.Redirect "monitor.asp?func=startside&vislogud="&vislogud 
                    else
                        'response.Redirect "monitor.asp?func=startside&redType=1" 
                    end if
                    
                    response.Redirect "monitor.asp?func=startside&redType=1" 

                end if

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
                    
                    response.Cookies("2015")("lastlogin") = oRec("Mid")
                    
                    'session("rettigheder") = oRec("rettigheder")
                    else
                    response.Redirect "monitor.asp?func=startside&redType=3"

                    end if
                    oRec.close


                    '**** Tjekker om der er nogle ude af huset registeringer der skal afsluttes
                    strSQL = "SELECT id FROM udeafhuset WHERE medid =" & session("mid") & " AND isnull(til) ORDER BY id DESC"
                    oRec.open strSQL, oConn, 3
                    if not oRec.EOF then

                        sqlNow = year(now) &"-"& month(now) &"-"& day(now)
                        sqlNowTime = hour(now) &":"& minute(now)
                        sqlTilDateTime = sqlNow &" "& sqlNowTime
                        strConn = "UPDATE udeafhuset SET tildato = '"& sqlNow &"', tiltidspunkt = '"& sqlNowTime &"', til = '"& sqlTilDateTime &"' WHERE id = "& oRec("id")
                        oConn.execute(strConn)

                        response.Redirect "monitor.asp?func=startside&redType=4&medid="& session("mid")
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


                    if cint(useJobAktPicker) = 1 then
                        if skiftjob = "0" then

                                if fo_oprettetnytlogin = "1" then
                                response.Redirect "monitor.asp?func=startside&redType=1"
                                else
                                response.Redirect "monitor.asp?func=registrer" 
                                end if 

                        else

                          
                                response.Redirect "monitor.asp?func=registrer&skiftjob=1" 
                            
                        end if
                    else

                        if fo_oprettetnytlogin = 0 then
                            vislogud = 1
                            'message = 0

                            response.Redirect "monitor.asp?func=startside&vislogud="&vislogud 
                        else
                            response.Redirect "monitor.asp?func=startside&redType=1" 
                        end if
                    
                        

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
        <input type="hidden" value="<%=lto %>" id="lto" />
        <div class="container" style="width:100%; height:100%">
            <div class="portlet">
                <div class="portlet-body">
                    <script src="js/monitor_jav15.js" type="text/javascript"></script>
                    <%
                    if func = "startside" then
                                       

                        redType = request("redType")

                        select case lto
                        case "cflow"
                         %>
                            <div class="row">
                                <div class="col-lg-12" style="text-align:center"><img src="img/cflow_ogo_red_300ppt.jpg" width="300" /></div>
                            </div>
                     <%
                        case "miele"
                     %>
                            <div class="row">
                                <div class="col-lg-12" style="text-align:center"><img src="img/Miele.JPG" /></div>
                            </div>



                    <%case "outz", "cool" %>
                            <div class="row">
                                <div class="col-lg-12" style="text-align:center"><img src="img/coolsorption_logo.png" /></div>
                            </div>
                    <% 
                        end select
                       
                     %>   
                    
                            <input type="hidden" id="vislogud" value="<%=vislogud %>" />
                            <br />
                            <form action="monitor.asp?func=scan" id="scan" method="post">
                                <div style="text-align:center">
                                    <table style="display:inline-table">
                                        <tr>
                                            <td><input type="number" class="form-control input-small" id="RFID_field" name="RFID_field" placeholder="<%=monitor_txt_039 %>" autofocus autocomplete="off" /></td>
                                        </tr>
                                    </table>
                                </div>
                                <div class="row">
                              
                                <div class="col-lg-12" style="text-align:right;">
                                    <br /><br /><br />
                                    <input type="submit" value="<%=monitor_txt_006 %> >> " class="btn btn-default" /></div>
                            </div>
                                
                                
                            </form>

                            <style>
                                /* The Modal (background) */
                                .modal {
                                    display: none; /* Hidden by default */
                                    position: fixed; /* Stay in place */
                                    z-index: 1; /* Sit on top */
                                    padding-top: 100px; /* Location of the box */
                                    left: 0;
                                    top: 0;
                                    width: 100%; /* Full width */
                                    height: 100%; /* Full height */
                                    overflow: auto; /* Enable scroll if needed */
                                    background-color: rgb(0,0,0); /* Fallback color */
                                    background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
                                }

                                /* Modal Content */
                                .modal-content {
                                    background-color: #fefefe;
                                    margin: auto;
                                    padding: 20px;
                                    border: 1px solid #888;
                                    width: 500px;
                                 /*   height: 350px; */
                                }

                                .picmodal:hover,
                                .picmodal:focus {
                                text-decoration: none;
                                cursor: pointer;
                                }
                            </style>

                            <%if cint(useGuestLogin) = 1 then %>
                           
                                <input type="hidden" id="errorMessege1" value="<%=monitor_txt_040 %>" />
                                <input type="hidden" id="errorMessege2" value="<%=monitor_txt_041 %>" />
                                <input type="hidden" id="errorMessege3" value="<%=monitor_txt_042 %>" />
                                <input type="hidden" id="errorMessege4" value="<%=monitor_txt_043 %>" />
                                <input type="hidden" id="errorMessege5" value="<%=monitor_txt_044 %>" />
                                <input type="hidden" id="errorMessege6" value="<%=monitor_txt_045 %>" />

                                <div class="row">
                                    <div class="col-lg-12" style="text-align:center;"><a id="gustlogin" class="btn btn-sm btn-default"><%=monitor_txt_007 %></a></div>
                                </div>

                                <div id="guestmodal" class="modal">
                                    <div class="modal-content" style="min-height:350px;">
                                        <a href="monitor.asp?func=startside" style="color:darkgrey;"><span style="padding-left:95%; font-size:150%;" class="fa fa-times"></span></a>
                                        <form action="monitor.asp?func=guestlogin" id="guestlogin" method="post">
                                            <div class="row">
                                                <h5 class="col-lg-6"><%=monitor_txt_008 %> <span style="color:darkred;">*</span></h5>  
                                                <h5 class="col-lg-6"><%=monitor_txt_009 %> <span style="color:darkred;">*</span></h5>
                                            </div>

                                            <div class="row">
                                                <h5 class="col-lg-6"><input autocomplete="off" id="guestfirma" name="guestCompany" type="text" class="form-control guestsignField" />
                                                    <select id="guestfirmalist" class="form-control" style="display:none;" multiple>
                                                        <option value="0"><%=monitor_txt_010 %></option>
                                                    </select>
                                                </h5>
                                    

                                                <h5 class="col-lg-6"><input autocomplete="off" id="guestEmployee" name="guestEmployee" type="text" class="form-control guestsignField" />
                                                    <select id="guestEmployeelist" class="form-control" style="display:none;" multiple>
                                                        <option value="0"><%=monitor_txt_010 %></option>
                                                    </select>
                                                </h5> 
                                            </div>

                                            <div class="row">
                                                <h5 class="col-lg-6"><%=monitor_txt_011 %> <span style="color:darkred;">*</span></h5>
                                                <h5 class="col-lg-6"><%=monitor_txt_012 %> <span style="color:darkred;">*</span></h5>
                                            </div>
                                            <div class="row">
                                                <h5 class="col-lg-6">
                                                    +
                                                    &nbsp
                                                    <input type="text" class="form-control guestsignField" id="guestTlf" name="guestTlf" style="width:92%; display:inline-block;" /></h5>
                                                <h5 class="col-lg-6"><input autocomplete="off" type="text" class="form-control guestsignField" id="besogInput" name="besogEmployee" />
                                                    <select id="besogslist" class="form-control" style="display:none;" multiple>
                                                        <option value="0"><%=monitor_txt_010 %></option>
                                                    </select>
                                                </h5>
                                                
                                                <input type="hidden" id="visit" name="visit" />
                                            </div>

                                            <div class="row">
                                                <h5 class="col-lg-12" style="text-align:center;"><%=monitor_txt_013 & " " %> <%=lto & " " %> <%=monitor_txt_014 %>
                                                    <br /> <input type="checkbox" id="infoApr" CHECKED class="guestsignField" />
                                                </h5>
                                            </div>

                                            <br /><br />  
                                            <div class="row">
                                                <div class="col-lg-12" style="text-align:center;">
                                                    <span id="guestErrorMessege" style="color:darkred;"></span>
                                                </div>
                                            </div>

                                            <div class="row">
                                                <div class="col-lg-12" style="text-align:center;">
                                                    <a id="guestsub" style="width:150px;" class="btn btn-lg btn-default" ><%=monitor_txt_015 %></a>
                                                </div>
                                            </div>                                           

                                        </form>

                                        
                                      <!--  <div class="row">
                                            <div class="col-lg-12" style="text-align:center;">
                                                <a id="closeguestmodal" class="btn btn-sm btn-danger"><%=monitor_txt_016 %></a>
                                            </div>
                                        </div> -->

                                    </div>
                                </div>


                                <!-- Liste over gæster der er logget ind, så de kan logges ud igen -->
                                <div class="row">
                                    <div class="col-lg-12" style="text-align:center;"><a id="gustlist" class="btn btn-sm btn-default"><%=monitor_txt_017 %></a></div>
                                </div>
                                <div id="guestlistmodal" class="modal">
                                    <div class="modal-content">
                                        <span style="padding-left:95%; font-size:150%; cursor:pointer; color:darkgray;" id="closeguestlist" class="fa fa-times"></span>

                                        <table class="table">
                                            <thead>
                                                <tr>
                                                    <th><%=monitor_txt_018 %></th>
                                                    <th><%=monitor_txt_019 %></th>
                                                    <th><%=monitor_txt_020 %></th>
                                                    <th><%=monitor_txt_021 %></th>
                                                </tr>
                                            </thead>

                                            <tbody>
                                                <%
                                                nowSQL = year(now) &"-"& month(now) &"-"& day(now)
                                                strSQL = "SELECT lh.mid as mid, m.mnavn as mnavn, m.firma as firmanavn, lh.login as login FROM login_historik lh LEFT JOIN medarbejdere m ON (m.mid = lh.mid) WHERE lh.dato ='"&nowSQL&"' AND m.mansat = 4 AND IsNull(logud) = true"
                                                'response.Write strSQL
                                                oRec.open strSQL, oConn, 3
                                                while not oRec.EOF
                                                i = 0
                                                %>
                                                <tr>
                                                    <td><%=oRec("mnavn") %></td>
                                                    <td><%=oRec("firmanavn") %></td>
                                                    <td><%=oRec("login") %></td>
                                                    <td><a href="monitor.asp?func=guestlogud&guestid=<%=oRec("mid") %>" class="btn btn-sm btn-danger"><%=monitor_txt_021 %></a></td>
                                                </tr>
                                                <%
                                                i = i + 1
                                                oRec.movenext
                                                wend
                                                oRec.close
                                                %>
                                                <%if i = 0 then %>
                                                    <tr>
                                                        <td colspan="4" style="text-align:center;"><%=monitor_txt_022 %></td>
                                                    </tr>
                                                <%end if %>
                                            </tbody>

                                        </table>

                                       <!-- <div class="row" style="padding-top:35%;">
                                            <div class="col-lg-12" style="text-align:center;">
                                                <a id="closeguestlist" class="btn btn-sm btn-danger"><%=monitor_txt_016 %></a>
                                            </div>
                                        </div> -->
                                    </div>
                                </div>






                                <br /><br />
                            <%end if %>


                            


                            <%
                            if len(trim(request.Cookies("2015")("lastlogin"))) <> 0 then
                            medidloggetInd = request.Cookies("2015")("lastlogin")
                            else
                            medidloggetInd = 0
                            end if

                            call meStamdata(medidloggetInd)
                            %>

                            <div id="logudmodal" class="modal">
                                <div class="modal-content">
                                    <a href="monitor.asp?func=startside" style="color:darkgrey;"><span style="padding-left:95%; font-size:150%;" class="fa fa-times"></span></a>
                                    <div class="row">
                                        <div class="col-lg-12" style="text-align:center;"><h3><%=monitor_txt_023 & " " %><%=meNavn %></h3></div>
                                    </div>
                                    <div class="row">
                                        <div class="col-lg-12" style="text-align:center;"><h4><%=monitor_txt_024 %></h4></div>
                                    </div>


                                    <br /><br /><br /><br /><br /><br />
                                    <div class="row">
                                        <div class="col-lg-12" style="text-align:center;">                                       
                                            <table style="width:100%">
                                                <tr>
                                                    <td style="width:50%; text-align:center;">
                                                        <form action="../sesaba.asp?frommonitor=1&medid=<%=request.Cookies("2015")("lastlogin") %>" method="post">
                                                            <input type="submit" style="width:150px;" class="btn btn-lg btn-danger" value="<%=monitor_txt_025 %>">
                                                        </form>
                                                    </td>                                                   
                                                    <td style="width:50%; text-align:center;">
                                                        <form action="monitor.asp?func=meeting&medid=<%=request.Cookies("2015")("lastlogin") %>" method="post">
                                                            <input type="submit" style="width:150px;" class="btn btn-lg btn-warning" value="<%=monitor_txt_026 %>">
                                                        </form>
                                                    </td>                                                  
                                                </tr>
                                            </table>

                                        </div>
                                    </div>

                                </div>
                            </div>







                          
                            <%
                            if redType = 1 then
                                'medarb_navn = request("medarb_navn")

                                if len(trim(request.Cookies("2015")("lastlogin"))) <> 0 then
                                medidloggetInd = request.Cookies("2015")("lastlogin")
                                else
                                medidloggetInd = 0
                                end if

                                call meStamdata(medidloggetInd)

                                '** Er medarbejder med i Aumatisjon Eller Enginnering
                                call medariprogrpFn(medidloggetInd)
                                
                            %>
                                <div id="text_besked" class="row">        
                                    <h3 class="col-lg-12" style="text-align:center"><%=monitor_txt_027 & " " %><%=meNavn %><br /><span style="font-size:14px;"><%=monitor_txt_028 & " " %><%=formatdatetime(now, 1) & " " & formatdatetime(now, 3)  %></span>


                                            <%
                                            if cint(useJobAktPicker) = 1 then
                                                'AND instr(medariprogrpTxt, "#3#") = 0
                                                if instr(medariprogrpTxt, "#14#") = 0 AND instr(medariprogrpTxt, "#16#") = 0 AND instr(medariprogrpTxt, "#19#") = 0 then

                                                jobidC = "-1"
                                                aktidC = "-1"
                                                'end if

                                                '** IKKE PÅ COOKIE DA DET SKAL GÆLDE ALLE TERMINALER
                                                strSQLsel = "SELECT lha_jobid FROM login_historik_aktivejob_rel WHERE lha_mid = "& medidloggetInd & " AND lha_jobid <> 0 ORDER BY lha_id DESC"

                                                'response.write strSQLsel
                                                'response.flush

                                                oRec.open strSQLsel, oConn, 3
                                                if not oRec.EOF then

                                                    jobidC = oRec("lha_jobid")

                                                end if
                                                oRec.close
                

                                                 '** IKKE PÅ COOKIE DA DET SKAL GÆLDE ALLE TERMINALER
                                                strSQLsel = "SELECT lha_aktid FROM login_historik_aktivejob_rel WHERE lha_mid = "& medidloggetInd & " AND lha_aktid <> 0 ORDER BY lha_id DESC"
                                                oRec.open strSQLsel, oConn, 3
                                                if not oRec.EOF then

                                                    aktidC = oRec("lha_aktid")

                                                end if
                                                oRec.close


                                            jobTxtloggetindpaa = monitor_txt_046
                                            if jobidC <> "-1" AND jobidC <> "" then
                                            
                                                strSQLjob = "SELECT jobnavn, jobnr FROM job WHERE id = " & jobidC
                                                oRec.open strSQLjob, oConn, 3
                                                if not oRec.EOF then

                                                jobTxtloggetindpaa = oRec("jobnr") & " " & oRec("jobnavn")

                                                end if
                                                oRec.close

                                            end if

                                            %>

                                            <br /><br /><%=monitor_txt_029 %>: <%=jobTxtloggetindpaa %>

                                            <%end if %>
                                        <%end if %>
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
                                    <h3 class="col-lg-12" style="text-align:center"><%=monitor_txt_030 %> <%=medarb_navn %></h3>
                                </div>
                            <%
                            end if
                            %>
                             <%
                            if redType = 3 then
                                'medarb_navn = request("medarb_navn")
                            %>
                                <div id="text_besked" class="row">        
                                    <h3 class="col-lg-12" style="text-align:center"><%=monitor_txt_031 %></h3>
                                </div>
                            <%
                            end if
                            %>

                            <%if redType = 4 then %>
                                <%call meStamdata(request("medid")) %>
                                <div id="text_besked" class="row">        
                                    <h3 class="col-lg-12" style="text-align:center"><%=monitor_txt_032 & " " %><%=menavn %></h3>
                                </div>
                            <%end if %>

                            <%if redType = 5 then %>
                                <%call meStamdata(request("medid")) %>
                                <div id="text_besked" class="row">        
                                    <h3 class="col-lg-12" style="text-align:center"><%=monitor_txt_033 & " " %><%=menavn %></h3>
                                </div>
                            <%end if %>
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
                                    <td><a href="monitor.asp?func=login&medarb_navn=<%=medarb_navn %>&medid=<%=medid %>&stempelurindstilling=1" class="btn btn-success btn-lg"><b><%=monitor_txt_006 %></b></a></td>
                                </tr>
                            </table>

                            <table style="display:inline-table">
                                <tr>
                                    <td><a href="monitor.asp?func=login&medarb_navn=<%=medarb_navn %>&medid=<%=medid %>&stempelurindstilling=2" class="btn btn-success btn-lg"><b><%=monitor_txt_034 %></b></a></td>
                                </tr>
                            </table>
                        </div>
                    
                    <%end if



                    if func = "registrer" then

                        %>
                        <meta http-equiv="refresh" content="120;url=monitor.asp?func=startside" />
                        <script src="js/ugeseddel_2011_jav6.js" type="text/javascript"></script>
                        <!--#include file="inc/touchmonitor_inc.asp"-->

                        <%
                        
                        varTjDatoUS_man = year(now)&"/"&month(now)&"/"&day(now)

                        call mobil_week_reg_dd_fn()

                        medarb_navn = session("user") 'request("medarb_navn")
                        medid = session("mid") 'request("medid")

                        call meStamdata(medid)

                        call browsertype()

                       

                    '******* Tjekker logindtid ********'
                    strSQL = "SELECT id, login FROM login_historik WHERE mid = "& medid &" AND stempelurindstilling > 0 ORDER BY id DESC"
                    oRec.open strSQL, oConn, 3
                    if not oRec.EOF then
                    login = oRec("login")
                            
                    end if
                    oRec.close


                    if browstype_client <> "ip" then
                    %>
                          <table style="width:100%;">
                              <tr>
                                  <td valign="top"><h3 style="color:black;"><%=medarb_navn%> (<%=meNr %>)</h3>
                                  <br />
                                  <span style="font-size:12px;"><%=monitor_txt_035 & " " %> <%=login %></span>

                                  <% 
                                    overtidtot = 0
                                    strSQL6 = "SELECT SUM(timer) As overtidtimer, tdato FROM timer WHERE tmnr = "& medid &" AND tfaktim = 30 GROUP BY tmnr" 'AND tdato BETWEEN '"& sqlDatoStartATD &"' AND '"& sqlDatoEnd &"'
                                    oRec.open strSQL6, oConn, 3
                                    if not oRec.EOF then

                                    overtidtot = overtidtot + oRec("overtidtimer")*1
        
                                    end if 
                                    oRec.close 
                                      


                                  %><br /><span style="font-size:12px;"><%=monitor_txt_036 %>: - <!--<%=formatnumber(overtidtot, 2) %>--></span>


                                  </td>
                                  <td style="text-align:right" valign="top"><img style="width:15%; margin-bottom:2px; margin-left:55px" src="img/cflow_logo.png" /></td>
                              </tr>
                          </table>
                        
                        <% 
                        end if
                         
	                    %><div class="row"><%
                                
                          if browstype_client <> "ip" then
                            
                            %>
                            <div class="col-lg-3">&nbsp;
                            <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
                            <br /><br /><br /><br /><br /><br /><br /><br />
                              <a href="monitor.asp?func=startside" class="btn btn-lg btn-default"><< <b><%=monitor_txt_037 %></b></a>   

                        </div>
                        <%end if %>

                        <div class="col-lg-5"> 
                            
                            <%
                            touchMonitor = 1
                            call keyfigures(touchMonitor, login)
                              %>
                               
                        </div>

                        <%if browstype_client <> "ip" then %>
                        <div class="col-lg-3">

                                        <table class="table" style="width:400px;">

                                            <tr>
                                                <th colspan="4"><%=monitor_txt_038 %></th>
                                            </tr>
                                        

                            <%

                                

                                    l = 0
                                    strSQL = "SELECT id, login, logud, minutter FROM login_historik WHERE mid = "& medid &" AND stempelurindstilling > 0 AND logud IS NOT NULL ORDER BY login DESC LIMIT 20"
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF 
                                    login = oRec("login")
                                    logud = oRec("logud")

                                    
                                    if lastWeek <> datepart("ww", oRec("login"), 2,2) AND l > 0 then


                                            timerThisLastWeek = formatnumber(lastWeekTotal/60, 2)
                                            minutterThisDBLastWeek = lastWeekTotal

                                           
                                            if cdbl(timerThisLastWeek) < 1 then
                                            timerThisLastWeek = 0
                                            end if

                                            timerThis_kommaLastWeek = instr(timerThisLastWeek, ",")
                                            if timerThis_kommaLastWeek <> 0 then
                                            timerThisLastWeek = left(timerThisLastWeek, timerThis_kommaLastWeek - 1)
                                            end if

                                            timerThis_kommaLastWeek = instr(timerThisLastWeek, ".")
                                            if timerThis_kommaLastWeek <> 0 then
                                            timerThisLastWeek = left(timerThisLastWeek, timerThis_kommaLastWeek - 1)
                                            end if

                                            minutterThisLastWeek = 0
                                            minutterThisLastWeek = formatnumber(minutterThisDBLastWeek/60, 2)
                                            minutterThisLastWeek = right(minutterThisLastWeek, 2)
                                            minutterThisLastWeek = formatnumber((minutterThisLastWeek*60)/100, 0)
                                          

                                            if minutterThisLastWeek < 10 then
                                            minutterThisLastWeek = "0"& minutterThisLastWeek
                                            end if

                                           %>

                                                    <tr><td colspan="4" style="text-align:right;"><span style="font-size:9px;">Week <%=lastWeek %>:</span><b>
                                                    <%= timerThisLastWeek &":"& minutterThisLastWeek %>
                                                    </b></td></tr>

                                    <%
                                        lastWeekTotal = 0

                                    end if

                            
                                    if formatdatetime(login, 2) = formatdatetime(now, 2) then
                                        trBgcol = "lemonchiffon"
                                    else
                                        trBgcol = ""
                                    end if

                                    %>
                                            <tr style="background-color:<%=trBgcol%>;">
                                                <td><%=left(weekdayname(weekday(login)), 3) &". "& formatdatetime(login, 2)%></td>
                                                <td><%=formatdatetime(login, 3)%></td>
                                                <td>- <%=formatdatetime(logud, 3)%></td>
                                                <td style="text-align:right;">

                                                    <%  
                                                        
                                                        
                                            timerThis = formatnumber(oRec("minutter")/60, 2)
                                            minutterThisDB = oRec("minutter")

                                           
                                            if cdbl(timerThis) < 1 then
                                            timerThis = 0
                                            end if

                                            timerThis_komma = instr(timerThis, ",")
                                            if timerThis_komma <> 0 then
                                            timerThis = left(timerThis, timerThis_komma - 1)
                                            end if

                                            timerThis_komma = instr(timerThis, ".")
                                            if timerThis_komma <> 0 then
                                            timerThis = left(timerThis, timerThis_komma - 1)
                                            end if

                                            minutterThis = 0
                                            minutterThis = formatnumber(minutterThisDB/60, 2)
                                            minutterThis = right(minutterThis, 2)
                                            minutterThis = formatnumber((minutterThis*60)/100, 0)
                                          

                                            if minutterThis < 10 then
                                            minutterThis = "0"& minutterThis
                                            end if

                                           %>


                                                    <%= timerThis &":"& minutterThis %>

                                                </td>
                                              
                                            </tr>

                                    <%

                                    lastWeek = datepart("ww", oRec("login"), 2,2)
                                    lastWeekTotal = lastWeekTotal + oRec("minutter") 
                                    l = l + 1

                                    oRec.movenext
                                    wend
                                    oRec.close    
                                


                                        
                                          

                                           %>

                                                    <tr><td colspan="4" style="text-align:right;"><span style="font-size:9px;">...</b></td></tr>

                                    <%
                                        lastWeekTotal = 0



                            %></table>


                        </div>
                        <%end if %>

                         </div><!--row -->

                   
                    
                    <%end if %>


                    <% 
                    'Skriver til logfil    
                    if (func = "startside" AND redType = "1") OR func = "registrer" then


                         '********************* Skriver til logfil ********************************************************'
		            if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_14\to_2015\monitor.asp" then
                
                    
                    'response.write "d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\logfile_timeout_"&lto&".txt"
                    'response.end    

                            toVer = "ver2_14"

				            Set objFSO = server.createobject("Scripting.FileSystemObject")
				            select case lto

                            case "cflow"
				          
					            Set objF = objFSO.GetFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\logfile_timeout_"&lto&".txt")
					            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\logfile_timeout_"&lto&".txt", 8)
            				
				            
            				
				            objF.writeLine(session("user") &chr(009)&chr(009)&chr(009)& date &chr(009)& time &chr(009)&request.servervariables("REMOTE_ADDR")&chr(009)&" - Terminal")
				            objF.close	

                            end select


		            end if
		            '*******************************************************************************************'
                                       

                       
                    end if
                    %>


                </div>
            </div>
        </div>
<%end select %>

<%'=lto %>
<!--#include file="../inc/regular/footer_inc.asp"-->