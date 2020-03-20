
<%

    
	if instr(request.servervariables("HTTP_HOST"), "localhost") <> 0 then



		session("lto") = "intranet - local"
		lto = "intranet - local"
		'lto = "cflow"
		ltokey = request("ltokey")

    
	else


        thisIp = request.servervariables("REMOTE_ADDR")
        '1 gang
        'response.Cookies("timeout_2015")("ltokey") = "9K2017-1124-TO179"
        'response.Cookies("timeout_2015").expires = date + 2000
        select case left(thisIp, 7)
        case "213.161", "213.161", "89.8.64" 'clow
        ltokey = "9K2017-1124-TO179"
        session("lto") = ltokey

        case "222.222.222.222" 'Dencker

        case "111.111.111.111" 'Cool

        case else 'ltokey OR Cookie - Hvis Monitor ikke genkender IP nummer
        
        %>
        <%=request.servervariables("REMOTE_ADDR")%><br />
        <%

        if len(trim(request("ltokey"))) <> 0 then
        ltokey = request("ltokey")
        session("lto") = ltokey        
        response.Cookies("timeout_2015")("ltokey") = ltokey
        response.Cookies("timeout_2015").expires = date + 2000
        else
            
            if request.Cookies("timeout_2015")("ltokey") <> "" then
            ltokey = request.Cookies("timeout_2015")("ltokey")
            session("lto") = ltokey
            else

                if len(trim(session("lto"))) <> 0 then 'Inden for 12 timer
                session("lto") = session("lto")
                else
                response.Write "<br><br><center>An error has occurred<br><br>" & monitor_txt_001 & "<br><br>a. no cookie</center>"
                response.End
                end if

            end if

        end if

        end select
    


    'session("lto") = session("lto") 
    end if

  thisfile = "infoscreen.asp"   
    
%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->



<%
if Request.Form("AjaxUpdateField") = "true" then
    Select Case Request.Form("control")
        case "UploadHoursToAkt"
            aktid = request("aktid")
            medid = request("medid")
                   
            sqlNow = year(now) &"-"& month(now) &"-"& day(now) 
            call normtimerPer(medid, sqlNow, 0, 0)
            antal = ntimper '7

            timerkom = ""
            koregnr = ""
            destination = ""
            usebopal = 0

            call indlasTimerTfaktimAktid(lto, medid, antal, 0, aktid, 2, 0, 0, timerkom, koregnr, destination, usebopal)

        case "AddGuestToLunch"
            medid = request("medid")
            strConn = "INSERT INTO progrupperelationer SET projektgruppeid = 12, medarbejderid = "&medid&", teamleder = 0, notificer = 0"
            oConn.execute(strConn)
            
        case "GetEmployeeTel"
            sqlDTD = year(now) &"-"& month(now) &"-"& day(now)
            strTlfs = ""
            strSQL = "SELECT lh.mid, m.mnavn as mnavn, init, mtlf, lh.login as login, mansat, firma FROM login_historik lh LEFT JOIN medarbejdere m on (m.mid = lh.mid) WHERE lh.dato = '"& sqlDTD &"' AND isnull(logud) ORDER BY id DESC"
            oRec.open strSQL, oConn, 3  
            while not oRec.EOF
                if oRec("mtlf") <> "" AND trim(len(oRec("mtlf"))) > 8 AND instr(1, oRec("mtlf"), ".") = 0 then
                    if len(strTlfs) <> 0 then
                        strTlfs = strTlfs & "," & oRec("mtlf")
                    else
                        strTlfs = oRec("mtlf")
                    end if
                end if
            oRec.movenext
            wend
            oRec.close

            strTlfs = replace(strTlfs, " ", "")
            strTlfs = replace(strTlfs, "+", "")
            response.Write strTlfs

        case "Mail"
            
            strEmployeeMails = ""
            sqlDTD = year(now) &"-"& month(now) &"-"& day(now)
            ekspTxt = ""
            strSQL = "SELECT lh.mid as mid, m.mnavn as mnavn, init, mtlf, lh.login as login, mansat, firma, email FROM login_historik lh " 
            strSQL = strSQL & "LEFT JOIN medarbejdere m on (m.mid = lh.mid) WHERE lh.dato = '"& sqlDTD &"' AND isnull(logud) ORDER BY mnavn, id DESC"
            oRec.open strSQL, oConn, 3
            while not oRec.EOF
            
                if cint(oRec("mansat")) <> 4 then 'Ikke gust
                    medarbTxt = oRec("mnavn") &" "& oRec("init")
                else
                    medarbTxt = oRec("mnavn") &" - "& oRec("firma") & " (Gæst) "
                end if

                ekspTxt = ekspTxt & medarbTxt &";" & oRec("mtlf") &";"& oRec("login") &";"& vbCrLf

                'if cint(oRec("mid")) = 1 then
                    if InStr(1, oRec("email"), "@") > 0 AND InStr(1, oRec("email"), ".") > 0 then
                        strEmployeeMails = strEmployeeMails & "<"& oRec("email") &">;"
                    end if
                'end if                

            oRec.movenext
            wend
            oRec.close

            call TimeOutVersion()
                             
            filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	        filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)

            Set objFSO = server.createobject("Scripting.FileSystemObject")
				           
	    	if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\medarb.asp" then
				Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\timeoutcheckin_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
				Set objNewFile = nothing
				Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\timeoutcheckin_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
			else
				Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\timeoutcheckin_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
				Set objNewFile = nothing
				Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\timeoutcheckin_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
			end if

            file = "timeoutcheckin_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
            
            strOskrifter = infoscreen_txt_002&";"&infoscreen_txt_003&";"&infoscreen_txt_004&";"
            objF.WriteLine(strOskrifter & chr(013))
		    objF.WriteLine(ekspTxt)
		    objF.close

            Set myMail=CreateObject("CDO.Message")
            myMail.Subject="TimeOut - Employees checked in"
            myMail.From="timeout_no_reply@outzource.dk"

            select case lto
                case "cool"
                    strEmail = "<osn@outzource.dk>;<TA@coolsorption.com>;<LH@coolsorption.com>;<TW@coolsorption.com>"
                case "dencker"
                    strEmail = "<support@outzource.dk>;<sf@dencker.net>;"
                case else
                    strEmail = "<support@outzource.dk>;"
            end select
            'strmnavn = "Oliver Storm"
            myMail.To= strEmail 'strmnavn & "<"& strEmail &">"
            
            myMail.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data/"& file

            strmailtekst = infoscreen_txt_005 & " " & strmnavn & vbCrLf
            strmailtekst = strmailtekst & infoscreen_txt_051

            strmailtekst = strmailtekst & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
            strmailtekst = strmailtekst & "With best regards " & vbCrLf & "Timeout - mail service"
            

            myMail.TextBody = strmailtekst

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

            if len(trim(strEmail)) <> 0 then
            myMail.Send
            end if
            set myMail=nothing
                        
            'response.Write "Der er blevet sendt en mail med medarbejdere logget ind, til osn@outzource.dk"
            select case lto
                case "cool"
                    response.Write infoscreen_txt_041 & ", TA@coolsorption.com, LH@coolsorption.com, TW@coolsorption.com"
                case "dencker"
                    response.Write infoscreen_txt_041 & " sf@dencker.net"
                case else
                    response.Write infoscreen_txt_041 & " support@outzource.dk"
            end select



            if lto = "cool" then 'Sender anden mail til medarbejdere om at de skal forlade byggningen 

                if strEmployeeMails <> "" then
                    Set myMail=CreateObject("CDO.Message")
                    myMail.Subject="TimeOut - Security alert"
                    myMail.From="timeout_no_reply@outzource.dk"
                
                    myMail.To= strEmployeeMails 'strmnavn & "<"& strEmail &">"

                    'strmailtekst = infoscreen_txt_005 & " " & strmnavn & vbCrLf
                    'strmailtekst = strmailtekst & infoscreen_txt_051
                    strmailtekst = "An alarm has been triggered. Please leave the building"
                    strmailtekst = strmailtekst & "<br><br><br><br><br>"
                    strmailtekst = strmailtekst & "With best regards <br>"
                    strmailtekst = strmailtekst & "<b>Timeout - mail service</b>"

                    myMail.HTMLBody = "<html><head></head><body>"& strmailtekst &"</body></html>"

                
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

                    if len(trim(strEmail)) <> 0 then
                    myMail.Send
                    end if
                    set myMail=nothing
                end if
    
            end if

        case "DatoPopReg"

            medid = request("medid")
            regtype = request("regkey")
            startdato = request("startdato")
            startdatoSQL = year(startdato) &"-"& month(startdato) &"-"& day(startdato)
            slutdato = request("slutdato")
            slutdatoSQL = year(slutdato) &"-"& month(slutdato) &"-"& day(slutdato)
            'response.Write "ST " & startdatoSQL & " SL " & slutdatoSQL

            'if cdate(startdatoSQL) > cdate(slutdatoSQL) then
                'response.Write "Error"
                'response.End
            'end if

           ' udeafhusettype - 0 = ubestemt, 1 = forretningsrejse, 2 = ferie   
            
            
            call normtimerPer(medid, startdatoSQL, 0, 0)           
            antal = ntimper '7
                        
            select case regtype
                case "ferie"
                    select case lto
                        case "cool"
                            aktid = 13
                        case "dencker"
                            aktid = 32660
                        case "outz"
                            aktid = 938
                        case else
                            aktid = 0
                    end select

                    udeafhusettype = 2
                case "forrejse"
                    select case lto
                        case "cool"
                            aktid = 33
                        case "outz"
                            aktid = 5577
                    end select

                    udeafhusettype = 1
            end select

                
            startdatoFullSQL = startdatoSQL & " 00:00:00"
            slutdatoFullSQL = slutdatoSQL & " 23:59:59"

            strSQL = "INSERT INTO udeafhuset (medid, fradato, fratidspunkt, tildato, tiltidspunkt, heledagen, fra, til, udeafhusettype) VALUES ("& medid &", '"& startdatoSQL &"', '00:00:00', '"& slutdatoSQL &"', '23:59:59', 1, '"& startdatoFullSQL &"', '"& slutdatoFullSQL &"', "& udeafhusettype &")"
            oConn.execute(strSQL)

            'Finder sidste udeafhuset id så det kan bruges til extsystid på timer tablenen
            lastUid = 0
            strSQL = "SELECT id FROM udeafhuset ORDER BY id DESC" 
            oRec3.open strSQL, oConn, 3
            if not oRec3.EOF then
                lastUid = oRec3("id")
            end if
            oRec3.close
            'response.Write lastUid
            'response.end
            ' Lægger timer ind for i dag, så man kan se det med det samme

            timerkom = ""
            koregnr = ""
            destination = ""
            usebopal = 0
            call indlasTimerTfaktimAktid(lto, medid, antal, 0, aktid, 0, startdatoSQL, lastUid , timerkom, koregnr, destination, usebopal)



    end select
response.End
end if
%>






<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/cls_checkin.asp"-->




    <%if request("print") = "1" then %>

      
        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u><%=infoscreen_txt_001 %></u></h3>
                <div class="portlet-body"> 
                    <table class="table dataTable table-striped table-bordered table-hover ui-datatable">
                        <thead>
                            <tr>
                                <th><%=infoscreen_txt_002 %></th>
                                <th><%=infoscreen_txt_003 %></th>
                                <th><%=infoscreen_txt_004 %></th>
                            </tr>
                        </thead>
                        <tbody>                            
                            <% 
                            sqlDTD = year(now) &"-"& month(now) &"-"& day(now)
                            strSQL = "SELECT lh.mid, m.mnavn as mnavn, init, mtlf, lh.login as login, mansat, firma FROM login_historik lh LEFT JOIN medarbejdere m on (m.mid = lh.mid) WHERE lh.dato = '"& sqlDTD &"' AND isnull(logud) ORDER BY id DESC"
                            'response.Write strSQL
                            oRec.open strSQL, oConn, 3
                            while not oRec.EOF

                                if cint(oRec("mansat")) <> 4 then 'Ikke gust
                                    medarbTxt = oRec("mnavn") &" "& oRec("init")
                                else
                                    medarbTxt = oRec("mnavn") &" - "& oRec("firma") & " (Gæst) "
                                end if

                            %>
                            <tr>
                                <td><%=medarbTxt %></td>
                                <td><%=oRec("mtlf") %></td>
                                <td><%=oRec("login") %></td>
                            </tr>
                            <%
                            oRec.movenext
                            wend
                            oRec.close
                            %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

    <%
        Response.Write("<script language=""JavaScript"">window.print();</script>")
    response.End
    end if 
    %>









    <%
    func = request("func")

    select case func
        
        case "scan"
            key = request("keyfield")

            if len(trim(key)) <> 0 then

                medid = 0
                call CheckKey(key, 3)  

                response.Redirect "infoscreen.asp?func=addlunch&medid="& medid
                
            end if

        case "EmployeeNotFound"
            %>
            <div class="container">
                <div class="portlet">
                    <div class="portlet-body">
                        <div class="row">
                            <div class="col-lg-12" style="text-align:center"><h4><%=infoscreen_txt_055 %></h4></div>
                        </div>

                        <br /><br />
                        <div class="row">
                            <div class="col-lg-12" style="text-align:center;">
                                <a href="infoscreen.asp?frokostvisning=1" class="btn btn-lg btn-default"><%=infoscreen_txt_044 %></a>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <%
            response.End


        case "addlunch"

            medid = request("medid")
            call meStamdata(medid)

            'Tjekeker om medarbejderen er med i forkost projektgruppe 
            mediFrokostordning = 0
            strSQL = "SELECT medarbejderid FROM progrupperelationer WHERE medarbejderid = "& medid & " and projektgruppeid = 11 OR projektgruppeid = 12"
            oRec.open strSQL, oConn, 3
            if not oRec.EOF then
                mediFrokostordning = 1
            end if
            oRec.close

            if mediFrokostordning = 1 then
                harspist = 0
                sqlNow = year(now) &"-"& month(now) &"-"& day(now)
                strSQL = "SELECT timer FROM timer WHERE tmnr = "& medid & " AND taktivitetid = 12 AND tdato = '"& sqlNow &"'"
                oRec.open strSQL, oConn, 3
                if not oRec.EOF then
                    harspist = 1
                end if
                oRec.close
            end if

            if mediFrokostordning = 1 AND harspist = 0 AND cint(memansat) <> 4 then
                response.Redirect "infoscreen.asp?func=addlunchOK&medid="& medid
            end if
            %>
                <div class="container">
                    <div class="portlet">
                        <div class="portlet-body">
                            <div class="col-lg-12" style="text-align:center"><h3><%=infoscreen_txt_005 & " " %><%=menavn %></h3></div>
                            <%if mediFrokostordning = 1 then %>
                                <%if harspist = 0 then %>
                                <div class="col-lg-12" style="text-align:center"><h4><%=infoscreen_txt_006 %></h4></div>
                                <%else %>
                                <div class="col-lg-12" style="text-align:center"><h4><%=infoscreen_txt_042 %></h4></div>
                                <%end if %>
                            <%else %>
                            <div class="col-lg-12" style="text-align:center"><h4><%=infoscreen_txt_043 %></h4></div>
                            <%end if %>
                            <br /><br /><br /><br /><br /><br />
                            <div class="row">
                                <div class="col-lg-12" style="text-align:center;">  
                                    <%if mediFrokostordning = 1 then %>
                                    <table style="width:100%">
                                        <tr>
                                            <%if harspist = 0 then %>
                                            <td style="width:50%; text-align:center;"><a href="infoscreen.asp?func=addlunchOK&medid=<%=medid %>" class="btn btn-lg btn-success"><%=infoscreen_txt_007 %></a></td>
                                            <td style="width:50%; text-align:center;"><a href="infoscreen.asp?frokostvisning=1" class="btn btn-lg btn-default"><%=infoscreen_txt_008 %></a></td>
                                            <%else %>
                                            <td style="width:100%; text-align:center;"><a href="infoscreen.asp?frokostvisning=1" class="btn btn-lg btn-default"><%=infoscreen_txt_044 %></a></td>
                                            <%end if %>
                                        </tr>
                                    </table>
                                    <%else %>
                                    <table style="width:100%">
                                        <tr>
                                            <td style="width:100%; text-align:center;"><a href="infoscreen.asp?frokostvisning=1" class="btn btn-lg btn-default"><%=infoscreen_txt_044 %></a></td>
                                        </tr>
                                    </table>
                                    <%end if %>
                                </div>
                            </div>

                        </div>
                    </div>
                </div>
            <%            
            response.end

        case "addlunchOK"
            'Laver en frokost reg i time tabellen
            medid = request("medid")
            aktid = 12 'Id for frokost aktivitet
            antal = 1
            timerkom = ""
            koregnr = ""
            destination = ""
            usebopal = 0
            call indlasTimerTfaktimAktid(lto, medid, antal, 0, aktid, 2, 0, 0, timerkom, koregnr, destination, usebopal)

            response.Redirect "infoscreen.asp?frokostvisning=1"

           


    end select


 sqlDTD = year(now) &"-"& month(now) &"-"& day(now)


    select case lto
        case "cflow", "dencker"
            brugprojektgrupper = 1
        case else
            brugprojektgrupper = 0
    end select

    select case lto
        case "outz", "cool", "dencker", "hidalgo", "kongeaa"
            visaltidmedarbejder = 1 ' Dette betyder at medabrjederen skal vises selom han er logget ind.
        case else
            visaltidmedarbejder = 0 ' Her det kun ikke-loggede-ind-medarbejdere der vises
    end select

    if len(trim(request("frokostvisning"))) <> 0 then
        frokostvisning = request("frokostvisning")
    else
        frokostvisning = 0
    end if

    if cint(frokostvisning) = 1 then
        loggetIndValue = 1
        visaltidmedarbejder = 0

        frokostSQL = "OR tfaktim = 10"
        'frokostGROUP = ", tfaktim = 10"
        frokostGROUP = "" 
    else
        loggetIndValue = 0
        visaltidmedarbejder = visaltidmedarbejder

        frokostSQL = ""
        frokostGROUP = ""
    end if

    medarbvisning = 0

    if cint(brugprojektgrupper) = 1 then '* Det er muligt man skal kunne sætte dette uden brug af projektgrupper, men lige nu er det kun cflow der bruger det og de bruger brugprojektgrupper
        if len(trim(request("FM_visning"))) <> 0 AND frokostvisning = 0 then
            medarbvisning = request("FM_visning")
            MBVSEL0 = ""
            MBVSEL1 = ""
            MBVSEL2 = ""
            select case cint(medarbvisning)
                case 0
                    MBVSEL0 = "SELECTED"
                    visaltidmedarbejder = 1
                case 1
                    loggetIndValue = 1
                    visaltidmedarbejder = 0
                    MBVSEL1 = "SELECTED"
                case 2
                    loggetIndValue = 0
                    visaltidmedarbejder = 0
                    MBVSEL2 = "SELECTED"
            end select
        else
            if frokostvisning = 0 then
                loggetIndValue = 0
                visaltidmedarbejder = 0
                MBVSEL2 = "SELECTED"
            end if
        end if
    end if


    if cint(brugprojektgrupper) = 1 then
        if len(trim(request("FM_prg"))) <> 0 then
            FM_prg = request("FM_prg")
            response.Cookies("2015")("infoscreen_FM_prg") = FM_prg
        else
            if request.Cookies("2015")("infoscreen_FM_prg") <> "" then
                FM_prg = request.Cookies("2015")("infoscreen_FM_prg")
            else
                FM_prg = 4
            end if
        end if

        if cint(FM_prg) <> 0 then
            prgFilter = "AND projektgruppeid = "& FM_prg
        else
            prgFilter = ""
        end if

    end if


    select case lto
        case "dencker", "hidalgo", "kongeaa"
                useGuest = 0
        case else
                useGuest = 1
    end select
%>


    <style>
        .sorter:hover,
        .sorter:focus {
        text-decoration: none;
        cursor: pointer;
        }
    </style>


    

   <!-- <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script> -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/jquery-contextmenu/2.7.1/jquery.contextMenu.min.css">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-contextmenu/2.7.1/jquery.contextMenu.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-contextmenu/2.7.1/jquery.ui.position.js"></script>

    <script src="https://cdn.jsdelivr.net/npm/js-cookie@2/src/js.cookie.min.js"></script>

    <script src="js/infoscreen_jav10.js" type="text/javascript"></script>
    <script src="js/newsscroll_jav.js" type="text/javascript"></script>

    
    <input type="hidden" id="rightclick_txt_forrejse" value="<%=infoscreen_txt_036 %>" />
    <input type="hidden" id="rightclick_txt_arbejdhjemme" value="<%=infoscreen_txt_037 %>" />
    <input type="hidden" id="rightclick_txt_ferie" value="<%=infoscreen_txt_020 %>" />
    <input type="hidden" id="rightclick_txt_syg" value="<%=infoscreen_txt_016 %>" />
    <input type="hidden" id="rightclick_txt_guestfrokost" value="<%=infoscreen_txt_035 %>" />

    <input type="hidden" id="lto" value="<%=lto %>" />
    <input type="hidden" id="frokostvisning" value="<%=frokostvisning %>" />

    <%
        if frokostvisning = 1 then
            colsize = "12"
        else
            colsize = "8"

            if len(request.Cookies("infoscreenSize")) <> 0 then
                if request.Cookies("infoscreenSize") = "1" then
                    bigscreen = 1
                else
                    bigscreen = 0
                end if
            else
                bigscreen = 0
            end if

        end if

        if bigscreen = 1 then
            antalrows = 4
            nameWidth = "325px"
            statusWidth = "250px"
            lettersAllowed = 32
        else
            if frokostvisning = 1 then
            antalrows = 5
            else
            antalrows = 2
            end if

            nameWidth = "215px"
            statusWidth = "150px"
            lettersAllowed = 15
        end if

    %>


    <%
    if frokostvisning = 0 AND lto = "cool" AND bigscreen = 1 then
        conSize = "95%"
    else 
        conSize = "100%"
    end if
     %>


    <div class="container" style="width:<%=conSize%>; height:100%">
        <div class="portlet">
            <div class="portlet-body">

                <%if lto = "cool" then %>
                    <div class="row">
                        <div class="col-lg-12" style="text-align:center"><img src="img/coolsorption_logo.png" /></div>
                    </div>
                <%end if %>

                <%
                    dateNow = day(now) &"-"& monthName(month(now)) &"-"& year(now)
                %>

                <div class="row">
                     
                        <%if cint(frokostvisning) = 1 then %>
                           <h4 class="col-lg-4"> <form action="infoscreen.asp?func=scan" method="post" id="scan">
                            - <input type="text" class="form-control input-small" id="keyfield" name="keyfield" style="width:100px; display:inline-block;" autofocus autocomplete="off" />
                            </form> </h4>
                        <%else %>
                        <h2 class="col-lg-4" style="color:gray"><%=dateNow %> &nbsp <span id="live_clock"></span></h2>
                        <%end if %>              
                    
                        <%if lto = "cflow" OR lto = "cool" OR lto = "dencker" OR lto = "outz" then %>
                            <div class="col-lg-8">
                                <div class="pull-right">
                                    <%if lto = "cflow" then %>
                                    <a href="infoscreen.asp?print=1" target="_blank" class="btn btn-sm" style="background-color:red; color:white;"><b><%=infoscreen_txt_045 %></b></a>
                                    <%end if %>
                                    <span class="btn btn-sm" style="background-color:red; color:white;" id="alarm"><b><%=infoscreen_txt_009 %></b></span>
                                </div>
                            </div>
                        <%end if %>

                </div>
                
                <div class="row">
                    <%if cint(brugprojektgrupper) = 1 then %>
                        <form action="infoscreen.asp?" method="post">
                            <div class="col-lg-6"><h3><%=infoscreen_txt_014 %> -
                             <select class="form-control" name="FM_prg" style="width:39%; display:inline-block" onchange="submit();">
                                    <option value="0"><%=infoscreen_txt_010 %></option>
                                    <%
                                        strSQL = "SELECT navn, id FROM projektgrupper WHERE id <> 10"
                                        oRec.open strSQL, oConn, 3
                                        while not oRec.EOF

                                            if cint(FM_prg) = cint(oRec("id")) then
                                                prgSEL = "SELECTED"
                                            else 
                                                prgSEL = ""
                                            end if

                                            response.Write "<option value='"&oRec("id")&"' "&prgSEL&" >"&oRec("navn")&"</option>"
                                        oRec.movenext
                                        wend
                                        oRec.close
                                    %>
                              </select>

                            <select class="form-control" style="display:inline-block; width:20%;" name="FM_visning" onchange="submit();">
                                <option value="0" <%=MBVSEL0 %>>Vis alle</option>
                                <option value="1" <%=MBVSEL1 %>>Vis mødte</option>
                                <option value="2" <%=MBVSEL2 %>>Vis ikke-mødte</option>
                            </select>
                            </h3></div>

                        </form>
                     <%else %>
                        <%if lto <> "cool" then %><div class="col-lg-6"><h3><%=infoscreen_txt_014 %></h3></div><%else %> <div class="col-lg-6"></div> <%end if %>
                    <%end if%>

                    <%if frokostvisning <> 1 then %>
                    <div class="col-lg-2"></div>
                    <div class="col-lg-4"><h3><%=infoscreen_txt_015 %></h3></div>
                    <%end if %>
                  </div> 

                <br />

                
                <div class="row">
                <div class="col-lg-<%=colsize %>">
                <table style="width:100%" id="myTable">
                    <tbody>

                       <!-- <tr>
                            <td><span style="color:dimgray;" id="sort_name" class="sorter fa fa-sort-down"></span></td>
                            <td style="padding-left:15px;"><span id="sort_color" style="color:dimgray;" class="sorter fa fa-sort-down"></span></td>
                        </tr> -->

                        <%

                        'Henter antal medarbejder for at der ikke skal komme en fejl op hver gang der kommer nye medarbejdere og får arrayets kapacitet ikke bliver større end nødvendigt
                        antalM = 100
                        strSQL = "SELECT count(mid) antalMedarbejdere FROM medarbejdere WHERE mansat = 1"
                        oRec.open strSQL, oConn, 3
                        if not oRec.EOF then
                            antalM = oRec("antalMedarbejdere") + 2
                        end if
                        oRec.close

                        dim timerColor, timerReg, timerTxt, sortNumber, erUde, medarbLoggetInd, medarbedjerid, medarbejderNavn
                        redim timerColor(antalM), timerReg(antalM), timerTxt(antalM), sortNumber(antalM), erUde(antalM), medarbLoggetInd(antalM), medarbedjerid(antalM), medarbejderNavn(antalM)

                        medarbejere = 0
                        antal = 0
                        select case lto
                            case "dencker"
                                if cint(brugprojektgrupper) = 1 then
                                    strSQL = "SELECT mid, mnavn FROM medarbejdere LEFT JOIN progrupperelationer ON (medarbejderid = mid) WHERE mansat = 1 AND mid <> 21 "& prgFilter &" GROUP BY mid ORDER BY mnavn"
                                else
                                    strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE mansat = 1 AND mid <> 21 AND mid <> 112 ORDER BY mnavn"
                                end if
                            case "cool"
                                if cint(frokostvisning) = 1 then
                                    strSQL = "SELECT mid, mnavn FROM medarbejdere LEFT JOIN progrupperelationer ON (medarbejderid = mid) WHERE mansat = 1 AND mid <> 1 AND projektgruppeid = 11 GROUP BY mid ORDER BY mnavn"
                                else
                                    strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE mansat = 1 AND mid <> 1 ORDER BY mnavn"
                                end if
                            case else
                                if cint(brugprojektgrupper) = 1 then
                                    strSQL = "SELECT mid, mnavn FROM medarbejdere LEFT JOIN progrupperelationer ON (medarbejderid = mid) WHERE mansat = 1 AND mid <> 1 "& prgFilter &" GROUP BY mid ORDER BY mnavn"
                                else
                                    strSQL = "SELECT mid, mnavn FROM medarbejdere WHERE mansat = 1 AND mid <> 1 ORDER BY mnavn"
                                end if                               
                        end select
                        m = 0
                        'response.Write strSQL
                        oRec.open strSQL, oConn, 3
                        while not oRec.EOF
                            'response.Write "<br> medarb " & oRec("mnavn")
                            medarbedjerid(m) = oRec("mid")
                            medarbejderNavn(m) = oRec("mnavn")
                        %>
                        <!--
                            syg = 20 = red
                            barnsyg = 21 = red

                            skole = 91 = green (info)
                            kursus = 91 = green (info)

                            Barsel = 22 = blue
                            ferieafholdt = 14 = blue
                            feriefridagebrugt = 13 = blue
                            afspadering = 31 = blue

                            gul ude af huset. 
                        -->
                            
                        <%
                           
                            timerColor(m) = "limegreen"
                            timerTxt(m) = infoscreen_txt_050 '"Mødt"'
                            timerReg(m) = 0
                            antalviste = 0
                            strSQL2 = "SELECT timer, tfaktim from timer WHERE tmnr = "& oRec("mid") &" AND tdato = '"& sqlDTD &"' AND (tfaktim = 20 OR tfaktim = 21 OR tfaktim = 22 OR tfaktim = 14 OR tfaktim = 13 OR tfaktim = 31 OR tfaktim = 91 OR tfaktim = 90 "& frokostSQL &") "& frokostGROUP &" ORDER BY tid"
                            '"GROUP BY tfaktim = 20, tfaktim = 21, tfaktim = 22, tfaktim = 14, tfaktim = 13, tfaktim = 31, tfaktim = 91, tfaktim = 90" Fjerern for at f[ den nyeste f'rst
                            'response.Write strSQL2
                            oRec2.open strSQL2, oConn, 3
                            antalfrokost = 0
                            while not oRec2.EOF
 
                            timerReg(m) = oRec2("timer")


                            select case oRec2("tfaktim")

                                case 20
                                timerColor(m) = "#d9534f"
                                timerTxt(m) = infoscreen_txt_016 '"Syg"
                                sortNumber(m) = "1"
                                case 21
                                timerColor(m) = "#d9534f"
                                timerTxt(m) = infoscreen_txt_017 '"Barn syg"'
                                sortNumber(m) = "2"

                                case 90
                                timerColor(m) = "#dee055"
                                timerTxt(m) = infoscreen_txt_046 '"Arbejder hjemmefra"'
                                sortNumber(m) = "3"

                                case 91
                                    select case lto
                                        case "dencker"
                                            timerColor(m) = "#5cb85c"
                                            timerTxt(m) = infoscreen_txt_018 '"Kursus / skole"'
                                            sortNumber(m) = "3"
                                        case "cool", "outz", "hidalgo", "kongeaa"

                                            'Tjekker hvornår medarbejderen kommer hjem igen
                                            sqlNow = year(now) &"-"& month(now) &"-"& day(now)
                                            backDate = sqlNow
                                            strSQL3 = "SELECT fradato, tildato FROM udeafhuset WHERE medid =" & medarbedjerid(m) & " AND udeafhusettype = 1 AND ((fradato < '"& sqlNow &"' AND tildato > '"& sqlNow &"') OR fradato = '"&sqlNow&"')"
                                            
                                            oRec3.open strSQL3, oConn, 3
                                            if not oRec3.EOF then
                                                if oRec3("tildato") = oRec3("fradato") then
                                                    backDate = dateadd("d", 1, oRec3("tildato")) 
                                                else 
                                                    backDate = year(oRec3("tildato")) &"-"& month(oRec3("tildato")) &"-"& day(oRec3("tildato"))
                                                    backDate = dateadd("d", 0, backDate) 'For at få i rigtig format
                                                end if

                                                'response.Write "herherh23" & backDate
                                            end if
                                            oRec3.close

                                            timerColor(m) = "#ad84ce"
                                            timerTxt(m) = infoscreen_txt_047 &" - <span style='font-size:9px;'>"& backDate &"</span>" '"Forretningsrejse"'
                                            sortNumber(m) = "3"
                                    end select
                                

                                case 22
                                timerColor(m) = "#5bc0de"
                                timerTxt(m) = infoscreen_txt_019 '"Barsel"'
                                sortNumber(m) = "6"
                                case 14
                                    select case lto
                                    case "cool", "dencker", "outz", "hidalgo", "kongeaa"
                                    
                                        'Tjekker hvornår medarbejderen kommer hjem igen
                                        sqlNow = year(now) &"-"& month(now) &"-"& day(now)
                                        backDate = sqlNow
                                        strSQL3 = "SELECT fradato, tildato FROM udeafhuset WHERE medid =" & medarbedjerid(m) & " AND udeafhusettype = 2 AND ((fradato < '"& sqlNow &"' AND tildato > '"& sqlNow &"') OR fradato = '"&sqlNow&"')"
                                        
                                        oRec3.open strSQL3, oConn, 3
                                        if not oRec3.EOF then
                                            if oRec3("tildato") = oRec3("fradato") then
                                                backDate = dateadd("d", 1, oRec3("tildato")) 
                                            else 
                                                backDate = year(oRec3("tildato")) &"-"& month(oRec3("tildato")) &"-"& day(oRec3("tildato"))
                                                backDate = dateadd("d", 0, backDate) 'For at få i rigtig format
                                            end if

                                            'response.Write "herherh23" & backDate
                                        end if
                                        oRec3.close

                                        if backDate <> sqlNow then
                                            backdateHTML = " - <span style='font-size:9px;'>"& backDate &"</span>"
                                        else
                                            backdateHTML = ""
                                        end if

                                        timerColor(m) = "#5bc0de"
                                        timerTxt(m) = infoscreen_txt_020 & backdateHTML '"Ferie"'
                                        sortNumber(m) = "4"
                                    case else
                                        timerColor(m) = "#5bc0de"
                                        timerTxt(m) = infoscreen_txt_020 '"Ferie"'
                                        sortNumber(m) = "4"
                                    end select
                                case 13
                                timerColor(m) = "#5bc0de"
                                timerTxt(m) = infoscreen_txt_021 'Feriefri'
                                sortNumber(m) = "5"
                                case 31
                                timerColor(m) = "#5bc0de"
                                timerTxt(m) = infoscreen_txt_022 '"Adspads."
                                sortNumber(m) = "7"
                                
                            end select

                            if cint(frokostvisning) = 1 then
                                if oRec2("tfaktim") = 10 then
                                    antalfrokost = antalfrokost + 1
                                end if
                            end if

                            antal = antal + 1

                            oRec2.movenext
                            wend
                            oRec2.close

                            'Hvis der ikke blev fundet nogen timer har medarbejderen ikke spist endnu
                            if cint(frokostvisning) = 1 then
                                if antalfrokost > 0 then
                                    timerColor(m) = "limegreen"
                                    timerTxt(m) = infoscreen_txt_023 '"Ikke spist"
                                else
                                    timerColor(m) = "grey"
                                    timerTxt(m) = infoscreen_txt_024 '"spist"'
                                end if
                            end if

                            
                        
                        %>
                        
                        <%
                            
                            erUde(m) = 0
                            nowTime = hour(now) &":"& minute(now)

                            'response.Write "Klokken lige nu er " & nowTime
                            dateTimeNow = sqlDTD &" "& nowTime

                            strSQLudeafhuset = "SELECT id, til FROM udeafhuset WHERE medid = "& oRec("mid") &" AND ((fra <= '"& dateTimeNow &"' AND (til >= '"& dateTimeNow &"') OR isnull(til)) OR (fradato <= '"& sqlDTD &"' AND tildato >= '"& sqlDTD &"' AND heledagen = 1)) AND udeafhusettype = 0"
                            tilbageigen = ""
                            oRec3.open strSQLudeafhuset, oConn, 3
                            if not oRec3.EOF then
                                erUde(m) = 1
                                tilbageigen = oRec3("til")
                            end if
                            oRec3.close
                           'response.Write "<br>areude " & erUde

                            if erUde(m) = 1 then
                                timerTxt(m) = infoscreen_txt_025 &" - <span style='font-size:9px;'>"& tilbageigen &"</span>" '"Ude af huset"
                                timerColor(m) = "#f0ad4e"
                                sortNumber(m) = "8"
                            end if



                            medarbLoggetInd(m) = 0
                            strSQL3 = "SELECT mid FROM login_historik WHERE mid = "& oRec("mid") &" AND dato = '"& sqlDTD &"' AND isnull(logud) ORDER BY id DESC"
                            'response.write strSQL3
                            oRec3.open strSQL3, oConn, 3
                            if not oRec3.EOF then
                            medarbejere = medarbejere + 1
                            medarbLoggetInd(m) = 1
                            end if
                            oRec3.close

                            if medarbLoggetInd(m) = 0 AND timerReg(m) = 0 AND erUde(m) = 0 then
                                timerTxt(m) = infoscreen_txt_026 '"ikke mødt"
                                timerColor(m) = "grey"
                                sortNumber(m) = "9"
                            end if

                            if frokostvisning <> 1 AND medarbLoggetInd(m) = 0 AND erUde(m) = 0 then
                                if timerReg(m) = 0 then 'Tjekker om medarbejderen har været logget ind men er logget ud igen - kun for at om medarbejderern er taget hjem
                                    strSQLlog = "SELECT mid FROM login_historik WHERE mid = "& oRec("mid") &" AND dato = '"& sqlDTD &"' AND isnull(logud) = false ORDER BY id DESC"
                                    oRec3.open strSQLlog, oConn, 3
                                    if not oRec3.EOF then
                                        timerTxt(m) = infoscreen_txt_034 '"taget hjem"
                                        timerColor(m) = "grey"
                                        sortNumber(m) = "8"
                                    end if
                                    oRec3.close
                                end if
                            end if


                            if cint(frokostvisning) = 1 then
                                if antalfrokost = 0 AND timerReg(m) <> 0 then
                                    medarbLoggetInd(m) = 0
                                    timerReg(m) = 0
                                end if
                            end if
                        %>
 
                        <%'if medarbLoggetInd(m) <> 1 OR timerReg(m) <> 0 OR erUde(m) <> 0 Then %>
                       <!-- <tr>
                            <td style="text-align:center"><span style="color:<%=timerColor(m) %>; font-size:25px;">O</span></td>
                            <td><%=oRec("mnavn") %></td>
                        </tr> -->

                    <!--  <tr>
                            <td><input type="hidden" value="<%=oRec("mnavn") %>" />
                                <div class="alert" style="height:100%; background-color:<%=timerColor(m) %>;">
                                    <strong style="color:white"><%=oRec("mnavn") %></strong>
                                </div>
                            </td>

                            <td style="padding-left:15px">
                                <input type="hidden" value="<%=sortNumber(m) %>" />
                                <div class="alert" style="height:100%; background-color:grey"">
                                    <strong style="color:white"><%=timerTxt(m) %></strong>
                                </div>
                            </td>
                       </tr> -->

                        <%'end if %>
                                                            
                        <%
                        'response.Write "<br> antal " & antal & " medarbejere " & medarbejere
                        if antal <> 0 OR medarbejere <> 0 then
                            antalviste = antalviste + 1
                        end if


                        m = m + 1
                        oRec.movenext
                        wend
                        oRec.close
                        %>

                        

                        <%if antal = 0 AND medarbejere = 0 then%>

                           <%
                            if cint(brugprojektgrupper) = 1 then
                               slutText = infoscreen_txt_027
                            else
                               slutText = infoscreen_txt_028
                            end if
                            %>

                          <!-- <tr>
                                <td>
                                    <div class="alert" style="height:100%; background-color:grey"">
                                        <strong style="color:white"><%=slutText %></strong>
                                    </div>
                                </td>
                            </tr> -->
                        <%end if %>

                        <%
                            'Finder de medbarejdere der skal vises
                            dim vis_medarbedjerid, vis_medarbnavn, vis_medarbtimerColor, vis_medarbsortNumber, vis_medarbtimerTxt
                            redim vis_medarbedjerid(antalM), vis_medarbnavn(antalM), vis_medarbtimerColor(antalM), vis_medarbsortNumber(antalM), vis_medarbtimerTxt(antalM)
                            vm = 0
                            for v = 0 to UBOUND(medarbedjerid)
                                
                                   ' response.Write "<br> medarbejder der MASKE bliver vist " & medarbejdernavn(v)

                                if len(medarbedjerid(v)) <> 0 then
                                
                                    if (medarbLoggetInd(v) = loggetIndValue) OR (cint(visaltidmedarbejder) = 1) OR (timerReg(v) <> 0 AND (medarbvisning = 0 OR medarbvisning = 2)) OR erUde(v) <> 0 Then

                                        vis_medarbedjerid(vm) = medarbedjerid(v)
                                        vis_medarbnavn(vm) = medarbejdernavn(v)
                                        vis_medarbtimerColor(vm) = timerColor(v)   
                                        vis_medarbsortNumber(vm) = sortNumber(i)
                                        vis_medarbtimerTxt(vm) = timerTxt(v)

                                        'response.Write "<br> medarbejder der bliver vist " & vis_medarbnavn(vm)
                                        vm = vm + 1
                                    end if
                                end if
                            next

                            'response.Write "antal til visning " & vm & " - "
                            
                            'response.Write round((vm/ antalrows) + 0.5) & "   vm " & vm & " ar " & antalrows
                            columnsOnEach = round((vm/ antalrows) + 0.6) 'round((vm + 0.6) / antalrows) '3 fordi de bliver delt op i 3 rækker + 0.6 så den altid runder op

                            'response.Write "herer " & columnsOnEach

                            %><tr><%

                            
                            select case lto
                                case "cool", "outz"
                                    contextmenu1class = "context-menu-one context-menu"
                                case "dencker"
                                    contextmenu1class = "context-menu-dencker context-menu"
                                case else
                                    contextmenu1class = ""                              
                            end select

                            columnReached = 0
                            columnsToAdd = columnsOnEach
                            lastmid = -1
                            i = 0
                            for i = 0 TO antalrows - 1
                                
                                'response.Write "<br> medarbejder "& vis_medarbnavn(i) & " - " & vis_medarbnavn(i+1)
                                %>
                                    
                                        <%'if vis_medarbnavn(i) <> "" then %>
                                        <!--<td>
                                            <div class="alert context-menu-one" id="<%=vis_medarbedjerid(i) %>" style="height:100%; background-color:<%=vis_medarbtimerColor(i) %>;">
                                                <strong style="color:white"><%=vis_medarbnavn(i) %></strong>
                                            </div>
                                        </td>
                                            <td style="padding-left:15px">
                                            <input type="hidden" value="<%=vis_medarbsortNumber(i) %>" />
                                            <div class="alert" style="height:100%; background-color:grey"">
                                                <strong style="color:white"><%=vis_medarbtimerTxt(i) %></strong>
                                            </div>
                                        </td> -->
                                        <td>
                                            <%for v = columnReached TO (columnsToAdd -1) %>
                                                                                              
                                              <!--  <div class="alert context-menu-one" id="<%=vis_medarbedjerid(v) %>" style="height:100%; background-color:<%=vis_medarbtimerColor(v) %>;">
                                                    <strong style="color:white"><%=vis_medarbnavn(v) %></strong>
                                                </div>

                                                <input type="hidden" value="<%=vis_medarbsortNumber(v) %>" />
                                                <div class="alert" style="height:100%; background-color:grey;">
                                                    <strong style="color:white"><%=vis_medarbtimerTxt(v) %></strong>
                                                </div> -->
                                                
                                                <%'if vis_medarbnavn(v) <> "" AND cdbl(vis_medarbedjerid(v)) <> cdbl(lastmid) then %>

                                                <%if len(vis_medarbnavn(v)) <> 0 then  %>
                                                    <div class="alert <%=contextmenu1class %>" id="<%=vis_medarbedjerid(v) %>" style="background-color:<%=vis_medarbtimerColor(v) %>; display:inline-block; width:<%=nameWidth%>;"><strong style="color:white;"><%=Left(vis_medarbnavn(v),lettersAllowed) %></strong></div>
                                                    <%if cint(frokostvisning) <> 1 then %>
                                                    <div class="alert" style="background-color:<%=vis_medarbtimerColor(v) %>; display:inline-block; width:<%=statusWidth%>;"><strong style="color:white;"><%=vis_medarbtimerTxt(v) %></strong></div>                                             
                                                    <%end if %>
                                                <%else %>
                                                    <div class="alert <%=contextmenu1class %>" style="background-color:grey; display:inline-block; width:300px; visibility:hidden;"><strong style="color:white;">None</strong></div>
                                                    <div class="alert" style="background-color:gray; display:inline-block; width:150px; visibility:hidden;"><strong style="color:white;">None</strong></div>
                                                <%end if %>
                                                <br />

                                                <%'end if %>
                                                <%lastmid = vis_medarbedjerid(v) %>
                                            <%next %>
                                            <%
                                                columnReached = columnsToAdd
                                                columnsToAdd = columnReached + columnsOnEach %>
                                        </td>

                                        <%'end if %>
                                    
                                <%                                 
                            next

                        %>
                        </tr>



                    </tbody>
                </table>
                </div>
                    <%if frokostvisning <> 1 then %>
                    <%
                        if lto = "dencker" then
                            newsStyle = "overflow-y:scroll;"
                        else
                            newsStyle = ""
                        end if
                    %>
                    <div class="col-lg-1"></div>
                    <div class="col-lg-4" id="newssection" style="max-height:1000px; <%=newsStyle%>">
                        <table>
                            <%
                                strSQL = "SELECT overskrift, brodtext, datofra, datotil, vigtig, filnavn, editor FROM info_screen WHERE ('"& sqlDTD &"' >= datofra AND '"& sqlDTD &"' <= datotil) AND type = 2 ORDER BY vigtig DESC, datofra DESC, id DESC"
                                'response.Write strSQL
                                oRec.open strSQL, oConn, 3
                                x = 0
                                while not oRec.EOF 
                            %>


                            <tr>
                                <td>
                                    <%if oRec("vigtig") = 1 then  %>
                                    <h3 style="color:#f44842;"><%=oRec("overskrift") %> <br /> <span style="font-size:12px; color:#9e9e9e;"><%=infoscreen_txt_029 %> <%=oRec("datofra") %></span></h3>
                                    <%else %>
                                    <h3><%=oRec("overskrift") %> <br /> <span style="font-size:12px; color:#9e9e9e;"><%=infoscreen_txt_052 &" "& oRec("datofra") &" "& infoscreen_txt_053 &" "& oRec("editor") %></span></h3>
                                    <%end if %>
                                </td>                                
                            </tr>
                            <tr>
                                <td>
                                    <%if oRec("brodtext") <> "" then  %>
                                    <strong><%=oRec("brodtext") %></strong>
                                    <%else %>
                                    <strong ><%=infoscreen_txt_030 %></strong>
                                    <%end if %>

                                    <%if len(oRec("filnavn")) <> 0 then %>
                                        <br /> <img src="https://timeout.cloud/timeout_xp/wwwroot/ver2_14/inc/upload/<%=lto%>/<%=oRec("filnavn") %>" alt='' border='0' style="max-width:500px;">
                                    <%end if %>

                                </td>
                            </tr>
                            <tr>
                                <td style="height:75px;">&nbsp</td>
                            </tr>


                            <%
                                x = x + 1
                                oRec.movenext
                                wend
                                oRec.close
                            %>


                            <%if x = 0 then %>

                            <tr>
                                <td><h3><%=infoscreen_txt_031 %></h3></td>
                            </tr>

                            <%end if %>

                        </table>
                    </div>
                    <%end if %>
                </div>

                <%if useGuest = 1 then %>
                <br /><br />

                <div class="row">
                    <div class="col-lg-6"><h3><%=infoscreen_txt_032 %></h3></div>
                </div>
                <br />
                <div class="row">
                    <div class="col-lg-6">
                        <table style="width:100%;">
                            <tbody>
                                <%
                                sqlNow = year(now) &"-"& month(now) &"-"& day(now)
                                dim guestEmployeeId, guestEmployeeName, guestEmployeeColor, guestEmployeeText, guestCompany
                                redim guestEmployeeId(10), guestEmployeeName(10), guestEmployeeColor(10), guestEmployeeText(10), guestCompany(10)
                                sqlNow = year(now) &"-"& month(now) &"-"& day(now) 

                                if cint(frokostvisning) = 1 then
                                    prjoin = "LEFT JOIN progrupperelationer pr ON (pr.medarbejderid = m.mid)"
                                    prwhere = "AND pr.projektgruppeid = 12"
                                    contextmenuclass = ""
                                else
                                    prjoin = ""
                                    prwhere = ""
                                    if lto = "cool" then
                                    contextmenuclass = "context-menu-two context-menu"
                                    else
                                    contextmenuclass = ""
                                    end if
                                end if

                                strSQLGuests = "SELECT lh.mid as mid, m.mnavn as mnavn, firma FROM login_historik lh LEFT JOIN medarbejdere m ON (m.mid = lh.mid) "& prjoin &" WHERE lh.dato = '"& sqlNow &"' AND m.mansat = 4 "& prwhere &" GROUP BY lh.mid"
                                'response.Write strSQLGuests
                                oRec.open strSQLGuests, oConn, 3
                                m = 0
                                while not oRec.EOF
                                
                                    guestEmployeeId(m) = oRec("mid")
                                    guestEmployeeName(m) = oRec("mnavn")
                                    guestCompany(m) = oRec("firma")
                                    guestEmployeeColor(m) = "grey"
                                    guestEmployeeText(m) = infoscreen_txt_024                                 

                                    if cint(frokostvisning) = 1 then
                                        'Tjekker om gæsten har spist
                                        strSQLLunch = "SELECT timer FROM timer WHERE tmnr =" & guestEmployeeId(m) & " AND tdato ='"& sqlNow &"' AND tfaktim = 10"
                                        oRec2.open strSQLLunch, oConn, 3
                                        if not oRec2.EOF then
                                            guestEmployeeColor(m) = "limegreen"
                                            guestEmployeeText(m) = infoscreen_txt_023 'Har spist
                                        end if
                                        oRec2.close
                                    else
                                        
                                        loggetud = ""
                                        strSQLlogud = "SELECT logud FROM login_historik WHERE mid =" & guestEmployeeId(m) & " ORDER BY login DESC"
                                        oRec2.open strSQLlogud, oConn, 3
                                        if not oRec2.EOF then
                                            loggetud = oRec2("logud")                                         
                                        end if
                                        oRec2.close

                                        if loggetud = "" OR isnull(loggetud) then
                                            guestEmployeeColor(m) = "limegreen"
                                            guestEmployeeText(m) = infoscreen_txt_033
                                        else
                                            guestEmployeeColor(m) = "grey"
                                            guestEmployeeText(m) = infoscreen_txt_034
                                        end if

                                    end if

                                m = m + 1
                                oRec.movenext
                                wend
                                oRec.close


                                i = 0
                                for i = 0 TO UBOUND(guestEmployeeId)
                                    %><tr><%

                                        if guestEmployeeName(i) <> "" then

                                            %>
                                            <td>
                                               <%if cint(frokostvisning) = 1 then %> <a href="infoscreen.asp?func=addlunch&medid=<%=guestEmployeeId(i) %>"> <%end if %>                                            
                                                    <div class="alert <%=contextmenuclass %>" id="<%=guestEmployeeId(i) %>" style="height:100%; background-color:<%=guestEmployeeColor(i) %>;">
                                                        <strong style="color:white"><%=guestEmployeeName(i) %> - <%=guestCompany(i) %></strong>
                                                    </div>
                                               <%if cint(frokostvisning) = 1 then %> </a> <%end if %> 
                                            </td>
                                             <td style="padding-left:15px">
                                                <div class="alert" style="height:100%; background-color:grey"">
                                                    <strong style="color:white"><%=guestEmployeeText(i) %></strong>
                                                </div>
                                            </td>
                                            <%

                                        end if

                                        if i < 10 then
                                         if guestEmployeeName(i+1) <> "" then

                                            %>
                                            <td style="padding-left:30px">
                                               <%if cint(frokostvisning) = 1 then %> <a href="infoscreen.asp?func=addlunch&medid=<%=guestEmployeeId(i+1) %>"> <%end if %>
                                                    <div class="alert <%=contextmenuclass %>" id="<%=guestEmployeeId(i+1) %>" style="height:100%; background-color:<%=guestEmployeeColor(i+1) %>;">
                                                        <strong style="color:white"><%=guestEmployeeName(i+1) %> - <%=guestCompany(i+1) %></strong>
                                                    </div>
                                                <%if cint(frokostvisning) = 1 then %></a> <%end if %>
                                            </td>
                                            <td style="padding-left:15px">
                                                <div class="alert" style="height:100%; background-color:grey"">
                                                    <strong style="color:white"><%=guestEmployeeText(i+1) %></strong>
                                                </div>
                                            </td>
                                            <%

                                        end if
                                       end if




                                    %></tr><%
                                    i = i + 1
                                next
                                %>
                            </tbody>
                        </table>
                    </div>
                </div>
                <%end if 'Use guest %>

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
                        height: 250px;
                    }

                    .ui-datepicker { left:250px; } 
 
                </style>


                <!-- Dato pop up til ude af huset regsitrering - ferie og forretningsrejse -->
                <div id="datopop" class="modal">
                    <!-- Modal content -->
                    <div class="modal-content">

                            <input type="hidden" id="datoPopMedid" />
                            <input type="hidden" id="datoPopKey" />

                            <input type="hidden" id="errormesseege_1" value="<%=infoscreen_txt_048 %>" />
                            <input type="hidden" id="errormesseege_2" value="<%=infoscreen_txt_049 %>" />

                            <div class="row">
                               <!-- <div class="col-lg-6"><strong>Rejse dato</strong></div> -->
                                <div class="col-lg-3"></div>
                                <div class="col-lg-6"><strong><%=infoscreen_txt_038 %></strong></div>
                            </div>
                            <%dateNow = day(now) &"-"& month(now) &"-"& year(now) %>
                            <div class="row">
                               <!-- <div class="col-lg-6">
                                    <div class='input-group date'>
                                        <input type="text" class="form-control" id="datoPopSt" value="<%=dateNow %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon">
                                            <span class="fa fa-calendar"></span>
                                        </span>
                                    </div>
                                </div> -->
                                <div class="col-lg-3">
                                    <input type="hidden" class="form-control" id="datoPopSt" value="<%=dateNow %>" />
                                </div>
                                <div class="col-lg-6">
                                    <div class='input-group date'>
                                        <input type="text" class="form-control" id="datoPopSl" value="<%=dateNow %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon">
                                            <span class="fa fa-calendar"></span>
                                        </span>
                                    </div>
                                </div>
                            </div>

                            <br /><br /><br />
                            
                            <div class="row">
                                <div class="collg-12" style="text-align:center;">
                                    <button id="datoPopUpGodkend" class="btn btn-success" style="width:125px;" ><b><%=infoscreen_txt_039 %></b></button>
                                </div>
                            </div>

                            <div class="row">
                                <div class="collg-12" style="text-align:center;">
                                    <button id="datoPopUpAnnulller" class="btn btn-default" style="width:125px;"><b><%=infoscreen_txt_040 %></b></button>
                                </div>
                            </div>

                        </div>
                    </div>
                </div>



            </div>
        </div>




<!--
</div>
</div> -->


<!--#include file="../inc/regular/footer_inc.asp"-->