
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->


<div class="wrapper">
<div class="content">


<%
   
    public accessDenied, pwmNavn
    function CanEmployeeChangePassword(medarbejderid)

         findesLog = 0
         accessDenied = 0
         strSQL = "SELECT medid, stampdate, stamptime FROM forgot_password WHERE medid ="& medarbejderid &" AND password_updated = 0 ORDER BY stampdate DESC, stamptime DESC"
         'response.Write "strSQL" & strSQL
         oRec.open strSQL, oConn, 3
         if not oRec.EOF then
            findesLog = 1
            stampdate = year(oRec("stampdate")) &"-"& month(oRec("stampdate")) &"-"& day(oRec("stampdate"))
            stamptime = hour(oRec("stamptime")) &":"& minute(oRec("stamptime")) &":"& second(oRec("stamptime"))
            
            stampmovement = stampdate & " " & stamptime
         end if
         oRec.close

        call meStamdata(medarbejderid)
        pwmNavn = meNavn 
        

        if cint(findesLog) = 1 then

            nowmovement = year(now) &"-"& month(now) &"-"& day(now)
            nowmovement = nowmovement & " " & hour(now) &":"& minute(now) &":"& second(now)

            logdiff = datediff("n", stampmovement, nowmovement, 1)
            logdiffHours = logdiff/60

            if logdiffHours > 24 then
                accessDenied = 1
            end if

        else
            accessDenied = 1
        end if

    end function


    func = request("func")

    select case func

        case "changepassword"
     
            medid = request("medid")
            lto = request("lto")

            'Tjekker om Medarbeder må skifte password
            call CanEmployeeChangePassword(medid)
            if cint(accessDenied) = 0 then
                newpasssword = request("FM_newpassword")
                newpasssword_repeat = request("FM_newpassword_repeat")                

                if newpasssword = "" then
                    errortype = 204
			        call showError(errortype)
                    response.end
                end if

                if newpasssword <> newpasssword_repeat then
                    errortype = 203
			        call showError(errortype)
                    response.end
                end if
            
                strSQL = "UPDATE medarbejdere SET pw = MD5('"& newpasssword &"') WHERE mid = "& medid
                'response.Write strSQL
                oConn.execute(strSQL)

                strSQL = "UPDATE forgot_password SET password_updated = 1"
                'response.Write strSQL
                oConn.execute(strSQL)

                response.Redirect "newpassword.asp?func=password_updated&lto="&lto
            else 'accessDenied
                response.Redirect "https://timeout.cloud/"&lto
            end if

    
    case "newpassword"

        
        Private Function Decrypt(ByVal encryptedstring)
	        Dim x, i, tmp
	        encryptedstring = StrReverse( encryptedstring )
	        For i = 1 To Len( encryptedstring )
		        x = Mid( encryptedstring, i, 1 )
		        tmp = tmp & Chr( Asc( x ) - 1 )
	        Next
	        Decrypt = tmp
        End Function

        
        '78 = medarbejder id
        '87 = lto
    
        medarbejderid = Decrypt(request("78"))
        lto = request("87")
       
        'response.Write "<br> medarbjeder " & medarbejderid
        'response.Write "<br> lto " & lto

       
        strSQL = "SELECT login FROM medarbejdere WHERE mid =" & medarbejderid
        oRec.open strSQL, oConn, 3
        if not oRec.EOF then
            brugernavn = oRec("login")
        end if
        oRec.close


         'Tjekker om Medarbeder må skifte password
         call CanEmployeeChangePassword(medarbejderid)

%>
        <script src="js/newpassword_jav.js" type="text/javascript"></script>  
        <div class="container">

            <div class="portlet">
                <div class="portlet-body">
                    <div class="well well-white">

                        <%if cint(accessDenied) = 0 then %>

                        <form action="newpassword.asp?func=changepassword" method="post">
                            <input type="hidden" name="medid" value="<%=medarbejderid %>" />
                            <input type="hidden" name="lto" value="<%=lto %>" />

                            <div class="row">
                                <div class="col-lg-12" style="text-align:center;"><b><%=brugernavn %></b></div>
                            </div>

                            <div class="row">
                                <div class="col-lg-3"></div>
                                <div class="col-lg-2">New password:</div>
                                <div class="col-lg-2"><input type="password" name="FM_newpassword" id="FM_newpassword" class="form-control input-small" /></div>
                            </div>

                            <div class="row">
                                <div class="col-lg-3"></div>
                                <div class="col-lg-2">Repeat:</div>
                                <div class="col-lg-2"><input type="password" name="FM_newpassword_repeat" id="FM_newpassword_repeat" class="form-control input-small" /></div>
                            </div>
                        
                            <div class="row">
                                <div class="col-lg-5"></div>
                                <div class="col-lg-2" style="text-align:center;"><span style="color:darkred;" id="statusbox"></span></div>
                            </div>

                            <br />

                            <div class="row">
                                <div class="col-lg-12" style="text-align:center;"><button type="submit" class="btn btn-success btn-sm"><b>Gem</b></button></div>
                            </div>
                        </form>

                        <%else %>


                        <div class="row" style="text-align:center;">
                            <div class="col-lg-12"><h2 style="color:darkred">Access Denied!</h2><br /><h4 style="color:darkred">Request new link</h4>
                            Hi <b><%=pwmNavn %></b><br />
                            It has been moore than 24 hours ago you requested the new password, or your user is disabled. <br />Ask your TimeOut administrator to request a new password or use the forgot password function at the login page.
                                <br /><br />
                                 logdiffHours: <%=logdiffHours %><br />
                            findesLog: <%=findesLog %>
                            </div>
                           
                        </div>

                        <%end if %>

                    </div>
                </div>
            </div>
        </div>

    <%case "password_updated" 

        lto = request("lto")

    %>

        <div class="container">
            <div class="portlet">
                <div class="portlet-body">
                    <div class="well well-white">
                        <div class="row">
                            <div class="col-lg-12" style="text-align:center;"><h3>Password opdateret</h3>
                                <br /> <h4>Log ind her: <a href="https://timeout.cloud/<%=lto %>">timeout.cloud/<%=lto %></a></h4>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>


    <%end select %>



</div>
</div>
    

<!--#include file="../inc/regular/footer_inc.asp"-->