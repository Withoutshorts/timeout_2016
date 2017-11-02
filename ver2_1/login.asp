

<% thisfile = "login.asp" %>

<!--#include file="inc/regular/header_lysblaa_2015_inc.asp"-->



<%
'response.Buffer = true
'**** Hvis LTO skal køre på gammel version skal det ændres her ****'
'key = request("key")
'lto = request("lto")
'Response.redirect "../ver2_14/login.asp?key="&key&"&lto="&lto


'**** Sletter alle cookies ***'
'Dim x,y

'For Each x in Request.Cookies
'    Response.Write("<p>")
'    If Request.Cookies(x).HasKeys Then
'        For Each y in Request.Cookies(x)
'            'Response.Write(x & ":" & y & "=" & Request.Cookies(x)(y))
'            'Response.Cookies(x).expires = date - 1
'            'Response.Write("<br />")
'        Next
'    Else
'        'Response.Write(x & "=" & Request.Cookies(x) & "<br />")
'        'Response.Cookies(x).expires = date - 1
'    End If
'    Response.Write "</p>"
'Next
%>


<noscript>
  <META HTTP-EQUIV="Refresh" CONTENT="0;URL=nojava.asp">
</noscript>
    



<!--
<META HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">
<META HTTP-EQUIV="EXPIRES" CONTENT="Mon, 22 Jul 2002 11:12:01 GMT">
    -->
    

    




<%
 %>
<!--#include file="inc/connection/conn_db_inc.asp"-->
<!--#include file="inc/errors/error_inc.asp"-->
<!--#include file="inc/regular/global_func.asp"-->






<%


'** Fra email om nyt job TO mobile ***'
if len(trim(request("tomobjid"))) <> 0 then
session("tomobjid") = request("tomobjid")
end if

'*** Sætter lokal dato/kr format. *****
Session.LCID = 1030




session("spmettanigol") = session("spmettanigol") + request("attempt")

	If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
	
	'************************************ Login side *********************************************
	
	strSQL = "SELECT licens FROM licens WHERE id = 1"
	oRec.Open strSQL, oConn, 0, 1, 1
	if not oRec.EOF then
	licensto = oRec("licens")
	end if
	oRec.close


   

   

    if lto = "demo" then
	uval = "Guest"
	'uval = "Support"

    		else
			    if Request.Cookies("login")("usrval") <> "" then
			    uval = Request.Cookies("login")("usrval")
                else
			    uval = ""
                end if
			end if


     '*** Special rdir side, forvalgt initialer TEC / ESN **'
     if lto = "tec" OR lto = "esn" then

        if len(trim(request("usr"))) <> 0 then
        uval = request("usr") 
        uval = replace(uval, "ESN-", "")

        end if

    end if
   

    if lto = "demo" then
			pawval = "gu1234"
            'pawval = "su1234"
			else
			    if Request.Cookies("login")("pwval") <> "" then
			    pawval = Request.Cookies("login")("pwval")
			    else
			    pawval = ""
			    end if
			end if


      if len(pawval) <> 0 then
			    huskCHK = "CHECKED"
			    else
			    huskCHK = ""
			    end if

	
	'*** tjekker om det er PDA eller PC ****
	
   

      

    call browsertype()
    
        
        
    'if instr(request.servervariables("HTTP_USER_AGENT"), "Smartphone") <> 0 then
	
    if browstype_client = "ip" then
    'pixLeft = 20
	'pixTop = 40
	'tboxsize = 100

  
		
			    if request.Cookies("timeoutcloud")("mobileuser") <> "" AND (lto = "sdutek")  then 'SKAL KUN VÆRE DEM DER IKKE SKAL LOGGE PÅ IGEN = sdutek 
                '** AND lto <> "outz" AND lto <> "hestia" AND lto <> "oko" and lto <> "epi"       
                '** Er der cookie og Du allerede er loggget ind, behøver Du ikke logge ind igen på din telefon
                '** For at skifte bruger på mobil SKAL man slette sine cookies ****'
                
                    if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                    response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_14/timetag_web/timetag_web.asp?flushsessionuser=1"
                    else
                    response.redirect "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timetag_web/timetag_web.asp?flushsessionuser=1"
                    end if
				    
                    'response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_14/timetag_web/timetag_web.asp"

                end if
				
			


    %>
     


  

<html xmlns="http://www.w3.org/1999/xhtml">    
<head>
<!--<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />-->
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title><%=login_txt_012 %></title>

 
<script src="login_jav.js" type="text/javascript"></script>



    
    <div class="account-wrapper">
    <div class="account-body">

       <img src="to_2015/img/outzource_logo_4c.jpg" width="200" />
                <br /><label>TimeOut Mobile</label>
    <!--<div id="header"><%=login_txt_010 %></div>-->
    <form id="container" action="login.asp" method="POST">
        
        <div class="form-group">

            
                    <input type="text" id="m_login" class="form-control" name="login" placeholder="<%=login_txt_019 %>" value="<%=uval %>"/>
            </div>
       
        
       
           <div class="form-group">
                <input type="password" id="m_pw" name="pw" class="form-control" placeholder="Password" value="<%=pawval %>"/> 
            
            </div>
       

     
             <div class="form-group clearfix">
                <div class="pull-left">					
              
                    <input id="huskmig" name="huskmig" type="checkbox" value="1" <%=huskCHK %> /> <%=login_txt_011 %>
              
                </div>

                <div class="pull-right">
                <a href="timereg/sendpw.asp?lto=<%=lto %>" target="_blank" style="color:#999999; font-size:11px; text-decoration:underline;">Forgot password</a>
                </div>
            </div>

         <!--   <div class="checkbox pull-left">
               
              <label style="display:inline;">
                    <input id="Checkbox1" name="huskmig" type="checkbox" value="1" <%=huskCHK %> /> <%=login_txt_011 %>
                  
              </label>
             -->
           
           
       
        
            <!---------------- Stempelur ----------------->
         
            <%
            select case lto
                case "epi2017" 'KUN AVI DK intw. Sættes automatisk
                case else
                call stempelurLoginSetting
			end select
            %>
        
                     
                  
            

          
                <input type="submit" class="btn btn-secondary btn-block btn-sm" value="Logind >> "/><br />
                <span style="font-size:10px; color:#999999;">[<%=lto %>]</span>
      
           

    </form>

    </div></div>

    <!--
    <br /><br /><br /><br /><br />
    <div style="padding:20px;"><span style="font-size:8px; color:#999999;"><%=request.servervariables("HTTP_USER_AGENT") %></span></div>
    -->


    <%
	else


    %>

   
	<div class="account-wrapper">
    <div class="account-body">


       
        <%
            
            
          
            
        select case lto 
        case "demo", "intranet - local" 
       ' Response.Write "<b>Velkommen til vores online TimeOut demo.</b>"_
		'	&" <br>"_
	'		&" Login og password til demoen er: <b>guest / 1234 </b> .<br><br>"_
'			&" Husk! I er altid velkommen til at kontakte OutZourCE på "_
	'		&" +45 2684 2000 eller på email: <a href='mailto:timeout@outzource.dk' class=vmenu>timeout@outzource.dk</a>"
			'&" <br><br>God fornøjelse...<br>&nbsp;"
        case "kejd_pb2", "wwf2", "epi_cati"
        Response.write "<h4 style=""color:red;"">DEMO - DEMO - DEMO - DEMO</h4>"
        case else
        end select'OR lto = "intranet - local"  %>
        <form action="login.asp" method="POST">
            



            <%if lto = "demo" then %>


               
                <h2 style="color:black"><img src="to_2015/img/outzource_logo_4c.jpg" width="200" />
                <br />Demo</h2>
                <br />
                


                <div class="form-group clearfix">
                    <div class="pull-left" style="margin-top:10px; color:red">*</div><input type="text" name="FM_ref" id="FM_ref" value="" class="form-control pull-right" placeholder="<%=login_txt_015 %>" style="width:96%">
                </div>

                <div class="form-group clearfix">
                    <div class="pull-left" style="margin-top:10px; color:red">*</div><input type="text" name="FM_firma" id="FM_firma" value="" class="form-control pull-right" placeholder="<%=login_txt_016 %>" style="width:96%">
                </div>
                <div class="form-group clearfix">
                    <div class="pull-left" style="margin-top:10px; color:red">*</div><input type="text" name="FM_email" id="FM_email" value="" class="form-control pull-right" placeholder="<%=login_txt_017 %>" style="width:96%">
                </div>
                <div class="form-group clearfix">
                    <input type="text" name="FM_tel" id="Text1" value="" class="form-control pull-right" placeholder="<%=login_txt_018 %>" style="width:96%">
                </div>

                <input type="hidden" name="login" id="login" value="<%=uval%>" placeholder="<%=login_txt_019 %>" class="form-control" tabindex="1">
                <input type="hidden" name="pw" id="pw" value="<%=pawval %>" placeholder="Password" class="form-control" tabindex="2">

            


            <%else 
            
            
            select case datepart("w", now, 2,2)
                case 1
                dayNameTxt = "Monday"
                case 2
                dayNameTxt = "Tuesday"
                case 3
                dayNameTxt = "Wednesday"
                case 4
                dayNameTxt = "Thursday"
                case 5
                dayNameTxt = "Friday"
                case 6
                dayNameTxt = "Saturday"
                case 7
                dayNameTxt = "Sunday"
            end select


             select case right(datepart("d", now, 2,2), 1)
                case 1
                dayModeTxt = "Beautiful"
                case 2
                dayModeTxt = "Happy"
                case 3
                dayModeTxt = "Nice"
                case 4
                dayModeTxt = "Great"
                case 5
                dayModeTxt = "Super"
                case 6
                dayModeTxt = "Fresh"
                case 7
                dayModeTxt = "Perfect"
                case 8
                dayModeTxt = "Brilliant"
                case 9
                dayModeTxt = "Blessed"
                case 0
                dayModeTxt = "Joyous"
            end select
                    
                
            if month(now) = 1 AND datepart("d", now, 2,2) < 5 then
                dayNameTxt = dayNameTxt & " And a Great " & year(now)
            end if
                
            %>

            <h2 style="color:black">
                 <img src="to_2015/img/outzource_logo_4c.jpg" width="200" />
                <br />
                <span style="font-size:16px; font-weight:lighter;"><i> Have a <%=dayModeTxt &" "& dayNameTxt %></i></span>

            </h2>
            <br />

            

            <input type="hidden" id="lto" value="<%=lto %>" />

            <div class="form-group">
              <input type="Text" name="login" id="login" value="<%=uval%>" placeholder="<%=login_txt_019%>" class="form-control" tabindex="1">
            </div>

            <div class="form-group">
                <input type="Password" name="pw" id="pw" value="<%=pawval %>" placeholder="Password" class="form-control" tabindex="2">
                <input type="hidden" name="attempt" value="1">
            </div>   

            <div class="form-group clearfix">
                <div class="pull-left">					
              
                    <input id="huskmig" name="huskmig" type="checkbox" value="1" <%=huskCHK %> /> <%=login_txt_011 %>
              
                </div>

                <div class="pull-right">
                <a href="timereg/sendpw.asp?lto=<%=lto %>" target="_blank" style="color:#999999; font-size:11px; text-decoration:underline;">Forgot password</a>
                </div>
            </div>
            


            <!---------------- Stempelur ----------------->
            <%

               call stempelurLoginSetting
				
            %>
             



            <%end if %>



          
           

            <br /><br />


            <div class="form-group">              
                <button type="submit" id="loginsubmit" class="btn btn-secondary btn-block btn-sm" tabindex="4">
                Sign in &nbsp; <i class="fa fa-play-circle"></i>
                </button>

              <span style="font-size:10px; color:#999999;">[<%=lto %>]</span>                
            </div> <!-- /.form-group -->           


        </form>
    </div>
</div>
	
		
	
	
		
		
        <%
        'call browsertype()
        'if browstype_client <> "ip" then %>
        
        <script>
		
        
        
        function NewWin(url)    {
		window.open(url, 'Disclaimer', 'width=600,height=580,scrollbars=yes,toolbar=no');    
		}
		</script>
		<br /><br /><br />


		<center>
            <span style="font-size:11px; color:#999999;">
		© 2002 - <%=year(now) &" "& login_txt_029 %>&nbsp;&nbsp;|&nbsp;&nbsp; 
		<%=login_txt_030 %>&nbsp;&nbsp;|&nbsp;&nbsp; <%=login_txt_031 %><br>
		<%=login_txt_032 %>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="javascript:NewWin('timereg/disclaimer.asp');" target="_self" class="vmenu"><u><%=login_txt_033 %></u></a></font>&nbsp;&nbsp;|&nbsp;&nbsp;Optimeret til IE 5+, FireFox 3+ og Chrome 
        
		<% 
		select case lto
		case "infow_demo", "infow", "assurator", "viatech", "pcmanden", "dreist", "rage", "enho"
		%>
		&nbsp;&nbsp;|&nbsp;&nbsp;<b><%=login_txt_034 %></b> salg@informationworker.dk &nbsp;&nbsp;|&nbsp;&nbsp;Tlf: 70212128 
		<%case else%>
		&nbsp;&nbsp;|&nbsp;&nbsp;<b><%=login_txt_034 %></b> support@outzource.dk&nbsp;&nbsp;|&nbsp;&nbsp;Tlf: +45 25 36 55 00&nbsp;&nbsp;|&nbsp;&nbsp;Skype: OutZourCE
		
		<%end select%>
		</span>
        <br /><span style="font-size:10px; color:#999999;"><%=request.servervariables("HTTP_USER_AGENT") %></span>
        <br /><br>&nbsp;

   


               
		</center>

        <%end if %>
		
		
		
		<%'end if ' Iphone
		
		
	
	
else '** POST *****
    
	'************ Hvis der er brugt mere end 3 login forsøg *********
	if session("spmettanigol") > 3 then '3
	%>

	<%
	errortype = 25
	call showError(errortype)
	else
	varLogin = Request.Form("login")
	varPW = Request.Form("pw")


    
	if varlogin = "" Then
	%>

	<%
	errortype = 1
	call showError(errortype)
	
	else
	'***************** login og pw modtages og valideres ***********************************************
		
		if varPW = "" then
		%>

		<%
		errortype = 2
		call showError(errortype)
		
		else
		
        
        email = request.Form("FM_email")
        firma = request.Form("FM_firma")
        ref = request.Form("FM_ref")
        tel = request.Form("FM_tel")

        if ref = "" and lto = "demo" then               
            errortype = 179
	        call showError(errortype)
        else
        if firma = "" and lto = "demo" then
            errortype = 180
	        call showError(errortype)
        else
        if email = "" and lto = "demo" then
            errortype = 181
	        call showError(errortype)
        else
        if instr(email,"@") = 0 and lto = "demo" then
            errortype = 181
	        call showError(errortype)
        else


        '*** HØSTE PW *****'
        hostPw = 0
        hostMid = 0
        if lto = "tec" OR lto = "esn" then
                
               

                strSQL = "SELECT pw AS pw, mid FROM medarbejdere WHERE login = '"& varLogin &"'" 
                
                
            
                oRec.open strSQL, oConn, 3
                if not oRec.EOF then

                'response.write "pw: "& oRec("pw") & "<br>"

               
                if oRec("pw") = "xxx01092014" then
                hostPw = 1
                hostMid = oRec("mid")
                end if

                end if
                oRec.close

                
                if cint(hostPw) = 1 then

                strSQL = "UPDATE medarbejdere SET pw = MD5('"& trim(Request.Form("pw")) &"') WHERE mid = " & hostMid
                oConn.execute(strSQL)

                end if

        end if
        '***** End høste ************'

		
		
		'**** Bruges på timregside til at nulstille bruger ***'
		session("forste") = "j"
		
		'*** Er det en ansat eller en kunde der logger på?***'
		'*** Altid ansat, Kunder har deres  egen login fil **'

        call TimeOutVersion()
	
		
		stransatkunde = "1" 'request("FM_ansat_kunde")
		
		if stransatkunde = "1" then
		'DECODE(
            if cint(hostPw) = 1 then
            strSQL = "SELECT m.mansat, m.mnavn, b.rettigheder, m.mid, m.tsacrm, m.pw, m.timereg FROM medarbejdere m, brugergrupper b WHERE "_
		    &" login = '"& trim(varLogin)&"' AND (mansat <> 2 AND mansat <> 3) AND b.id = m.brugergruppe"
		    
            else
		    strSQL = "SELECT m.mansat, m.mnavn, b.rettigheder, m.mid, m.tsacrm, m.pw, m.timereg FROM medarbejdere m, brugergrupper b WHERE "_
		    &" login = '"& trim(varLogin)&"' AND (mansat <> 2 AND mansat <> 3) AND m.pw = MD5('"& trim(Request.Form("pw")) &"') AND b.id = m.brugergruppe"
		    end if

        else
		strSQL = "SELECT id, kundeid, email, password AS pw, navn AS Mnavn FROM kontaktpers WHERE email='"& request.Form("login") &"'"
		end if
		
		founduser = 0

        'response.write "strSQL "& strSQL
        'response.end
		
		oRec.open strSQL, oConn, 3
		if not oRec.EOF Then
		
		            founduser = 1		
            			
			            '*******************************************************************************************'
			            '*** Sætter sesions var. alt efter om det er en medarbejder eller en kunde der logger på****'
            			
            			
			            if stransatkunde = "1" then
			            strUsrId = Trim(oRec("Mid"))
			            session("user") = Trim(oRec("Mnavn"))
			            session("login") = strUsrId
			            session("mid") = strUsrId
			            session("rettigheder") = oRec("rettigheder")

                        
                            '**** Sætter mobile cookies("") til at forblive logget ind på mobil *****
                            response.Cookies("timeoutcloud")("mobilemid") = session("mid") 
                            response.Cookies("timeoutcloud")("mobileuser") = session("user")
                            response.Cookies("timeoutcloud")("rettigheder") = session("rettigheder")
                            response.Cookies("timeoutcloud").expires = date + 180
                            '************************************************************************


				            startside = oRec("tsacrm")
            				
            				
				            '*** Hent timrregside fra db eller Form *** 
				            'if len(trim(request("FM_0206"))) <> 0 then
				            'treg0206 = request("FM_0206")
				            'else
				            treg0206 = 1 'oRec("timereg") ALTID NY VERSION
				            'end if
            				
            			
            				
            				
			            else

                            '*** Kunde der er loggget på ****'
			                strUsrId = 0
			                thisKpid = oRec("id")
			                session("user") = Trim(oRec("Mnavn"))
			                session("login") = strUsrId
			                session("mid") = oRec("kundeid") 'Kundeid
			                session("rettigheder") = 0
			                startside = 2 'joblog_k.asp
			                treg0206 = 0
			
                        end if
			            '*******************************************************************************************'
            		
            		
            		
            		
            		
            		
		            '********************* Skriver til logfil ********************************************************'
		            if request.servervariables("PATH_TRANSLATED") <> "E:\www\timeout_xp\wwwroot\ver2_1\login.asp" AND request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then
                
                    
                    'response.write "d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\logfile_timeout_"&lto&".txt"
                    'response.end    

				            Set objFSO = server.createobject("Scripting.FileSystemObject")
				            select case lto
				            '*** Bruger case else undt. for nedenstående, der dog burde kunne skiftes unden probs. Minus måske "net" pga lto.
            				
				            case "titoonic"
					            Set objF = objFSO.GetFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_titoonic.txt")
					            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_titoonic.txt", 8)
				            case "skybrud"
					            Set objF = objFSO.GetFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_skybrud.txt")
					            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\log\logfile_timeout_skybrud.txt", 8)
            				
				            case else
					            Set objF = objFSO.GetFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\logfile_timeout_"&lto&".txt")
					            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\logfile_timeout_"&lto&".txt", 8)
            				
				            end select
            				
				            objF.writeLine(session("user") &chr(009)&chr(009)&chr(009)& date &chr(009)& time&chr(009)& request.servervariables("REMOTE_ADDR"))
				            objF.close	
		            end if
		            '*******************************************************************************************'
            		
            		
            		
		            '******************* Sætter sidste login dato *********************************************'
		            strMthNow = month(now)
            		
		            select case strMthNow
		            case 1
		            strMth = "Jan. "
		            case 2
		            strMth = "Feb. "
		            case 3
		            strMth = "Mar. "
		            case 4
		            strMth = "Apr. "
		            case 5
		            strMth = "Maj. "
		            case 6
		            strMth = "Jun. "
		            case 7
		            strMth = "Jul. "
		            case 8
		            strMth = "Aug. "
		            case 9
		            strMth = "Sep."
		            case 10
		            strMth = "Okt. "
		            case 11
		            strMth = "Nov. "
		            case 12
		            strMth = "Dec. "
		            end select
            		
		            strLastLogin =  day(now)& " " & strMth & " " & year(now)
		            session("dato") = strLastLogin
            		
		  end if
		  oRec.close
            		
		
		
		if founduser = 1 then
		
		
    		
		    if stransatkunde = "1" then

                if len(trim(strUsrId)) <> 0 then
                strUsrId = strUsrId
                else
                strUsrId = 0
                end if

            'strUsrId
		    strSQL = "SELECT lastlogin FROM medarbejdere WHERE mid = " & strUsrId 
		    else
		    strSQL = "SELECT lastlogin AS lastlogin FROM kontaktpers WHERE id="& thisKpid 
		    end if
    		
		    oRec.open strSQL, oConn, 3
		    if not oRec.EOF then
		    session("strLastlogin") = oRec("lastlogin")
		    end if
		    oRec.close
		    
		    
		     
		    
		    
		    '************************************************************************'
		    '***** Medarbejder login ***'
            '************************************************************************'
		    if stransatkunde = "1" then 'else kunde login
			
			    oConn.execute("UPDATE medarbejdere SET lastlogin = '"& strLastLogin &"', Mnavn = '"& session("user") &"', timereg = "& treg0206 &" WHERE Mid ="& strUsrId &"")
		        
                '** Tilmeld til nyhedsbrev ***'
		        if request("nyhedsbrev") = "1" then
                oConn.execute("UPDATE medarbejdere SET nyhedsbrev = 1 WHERE Mid ="& strUsrId &"")
                response.cookies("tsa")("nyhansw") = 1
                response.cookies("tsa").expires = date + 180
		        end if
		       
		       '*** Ferie Pl ==> Ferie afholdt, AldersRed, omSorg PL ==> Afholdt **'
                select case lto
                case "tec", "esn"
                highV = 6
                case else
                highV = 2
                end select
             
                   select case lto
                   
                   case "epi", "epi_no", "epi_sta", "epi_ab", "intranet - local", "epi2017"
                
                        if session("rettigheder") = 1 AND ((datepart("w", now, 2,2) = 1 AND _
                        datepart("w", session("strLastlogin"), 2,2) <> 1)) then  'HVER MANDAG    
                          
                            'call opdaterFeriePl(session("rettigheder"), highV)
                        
                        end if

                   case else
                        '*** ÆNDRE TIL kun første login hver dag
                        call opdaterFeriePl(session("rettigheder"), highV)
                   end select

		        
               '*** Konsoliderer timer *****'

               'Response.Write "<br><br>lto:" & lto & "dt: "& datepart("d", now, 2,2)  & " //last: "& datepart("d", session("strLastlogin"), 2,2) &" level: " & session("rettigheder")
               'Response.end

               if session("rettigheder") = 1 AND ((datepart("d", now, 2,2) = 18 AND _
               datepart("d", session("strLastlogin"), 2,2) <> 18) OR (datepart("d", now, 2,2) = 27 AND datepart("d", session("strLastlogin"), 2,2) <> 27)) then
                   select case lto
                   case "xx"
                   case "dencker", "epi", "epi_no", "epi_sta", "epi_ab", "epi2017"
                   call timer_konsolider(lto,0)
                   case "xintranet - local", "xoutz"
                   call timer_konsolider(lto,0)
                   end select
               end if

            

               '*** Indlæser tillæg fra AVI DK
               if ((session("mid") = 1 OR session("mid") = 2694 Or session("mid") = 2452 OR session("mid") = 32821 OR session("mid") = 2663) AND (lto = "epi2017" OR lto = "intranet - local")) then    

                '*EPI
                ddTjkSupplementSL = year(now) & "/" & month(now) & "/"& day(now) 

                ddTjkSupplementST = dateAdd("d", -45, now)
                ddTjkSupplementST = year(ddTjkSupplementST) & "/" & month(ddTjkSupplementST) & "/"& day(ddTjkSupplementST) 
                strSQLepiAViDKsupplement = "SELECT tid, tdato, sttid, sltid FROM timer WHERE taktivitetnavn = 'Data Collection' AND tdato BETWEEN '"& ddTjkSupplementST &"' AND '"& ddTjkSupplementSL &"' AND sttid <> '00:00:00' AND (origin = 11 OR origin = 12) AND overfort = 0"
                'Local
                'strSQLepiAViDKsupplement = "SELECT tid, tdato, sttid, sltid FROM timer WHERE tjobnavn = 'A-SK Restest simpel' AND taktivitetnavn = 'Support U/B' AND tdato BETWEEN '2017-06-15' AND '2017-07-01'"
               
                addSupplement15 = 0
                addTillaeg = 0
                oRec6.open strSQLepiAViDKsupplement, oConn, 3
                while not oRec6.EOF

                addSupplement15 = 0
                addTillaeg = 0
                addTillaeg8 = 0
                addTillaeg20 = 0
                addTillaegW = 0
                'addTillaeg20

                    tregDato = oRec6("tdato")
                    sttidthis = oRec6("tdato") &" "& formatdatetime(oRec6("sttid"), 3)
                    sltidthis = oRec6("tdato") &" "& formatdatetime(oRec6("sltid"), 3)

                    '**** FØR 08:00
                    if formatdatetime(oRec6("sttid"), 3) < "08:00:00" then
                        addSupplement15 = 1

                        if formatdatetime(oRec6("sltid"), 3) > "08:00:00" then
                            sluttidKri = oRec6("tdato") &" 08:00:00"
                        else
                            sluttidKri = sltidthis
                        end if

                        'Response.write "<br>"& oRec6("tid") &"sttidthis:"& sttidthis &" #"& sltidthis & "<br>"
                        addTillaeg8 = dateDiff("n", sttidthis, sluttidKri, 2,2)
                        addTillaeg8 = (addTillaeg8/60) 
                        
                      
                     end if

                     '**** Efter 20:00
                      if formatdatetime(oRec6("sltid"), 3) > "20:00:00" then
                        addSupplement15 = 1

                        if formatdatetime(oRec6("sttid"), 3)  > "20:00:00" then
                           starttidKri = sttidthis
                        else
                           starttidKri = oRec6("tdato") &" 20:00:00"
                        end if

                        'Response.write "<br>"& oRec6("tid") &"sttidthis:"& sttidthis &" #"& sltidthis & "<br>"
                        addTillaeg20 = dateDiff("n", starttidKri, sltidthis,2,2)
                        addTillaeg20 = (addTillaeg20/60) 
                        
                        

                     end if


               
                    '**** Lør / Søn ELLER Hellig ****'
                   
                     call helligdage(oRec6("tdato"), 0, lto)
                     if datePart("w", oRec6("tdato"), 2,2) = "6" OR datePart("w", oRec6("tdato"), 2,2) = "7" OR cint(erHellig) = 1 then
                        addSupplement15 = 1

                        'Response.write "<br>"& oRec6("tid") &"sttidthis:"& sttidthis &" #"& sltidthis & "<br>"
                        addTillaegW = dateDiff("n", sttidthis, sltidthis,2,2)
                        addTillaegW = (addTillaegW/60) 
                        'addTillaegW = 5
                    end if


                  

                    addTillaeg = formatnumber(addTillaeg8/1 + addTillaeg20/1 + addTillaegW/1, 2)
                    addTillaeg = replace(addTillaeg, ".", "") 
                    addTillaeg = replace(addTillaeg, ",", ".")   
                    if cint(addSupplement15) = 1 then

                        strSQLTidCopy = "INSERT INTO timer_tmp_for_copy SELECT * FROM timer WHERE tid = "& oRec6("tid")
                        oConn.execute(strSQLTidCopy)

                        lastTid = 0
                        strSQLTidLast = "SELECT tid FROM timer WHERE tid > 0 ORDER BY tid DESC LIMIT 1"
                        oRec5.open strSQLTidLast, oConn, 3
                        if not oRec5.EOF then

                            lastTid = oRec5("tid") 

                        end if
                        oRec5.close

                        strSQLTidUpdateAddOne = "UPDATE timer_tmp_for_copy SET tid = "& lastTid + 1
                        oConn.execute(strSQLTidUpdateAddOne)
                        
                        strSQLTidInsert = "INSERT INTO timer SELECT * FROM timer_tmp_for_copy WHERE tid = " & lastTid + 1
                        oConn.execute(strSQLTidInsert)

                        strSQLTidInsertDEL = "DELETE FROm timer_tmp_for_copy"
                        oConn.execute(strSQLTidInsertDEL)

                        strSQLTidUpdOfort = "UPDATE timer SET overfort = 1, overfortdt = '"& ddTjkSupplementSL &"' WHERE tid = " & oRec6("tid")
                        oConn.execute(strSQLTidUpdOfort)

                        strSQLTidUpdNew = "UPDATE timer SET tfaktim = 51, taktivitetnavn = 'Supplement 15', timer = "& addTillaeg &", origin = 112, editor = 'Supplement Copy' WHERE tid = " & lastTid + 1
                        oConn.execute(strSQLTidUpdNew)


                    end if
            
                    
                oRec6.movenext
                wend 
                oRec6.close 

               end if

                'Response.write "Overført OK"

		        'Response.end
		        
		        
		       
		        
				
				'*** Opdater login_historik (Stempelur) ****'		
				if len(request("FM_stempelur")) <> 0 then
				intStempelur = request("FM_stempelur")
				else
                    select case lto 
                    case "epi2017"
                    intStempelur = 1
                    case else
				    intStempelur = 0
                    end select
				end if
				
                session("dontLogind") = 0

                '**** Selvom stempelur er slået til skal følgende grupper ikke have et logind **'
                select case lto 
				case "cst"
                        
                        strSQLmansat = " mansat <> 0"
                        call medarbiprojgrp(21, session("mid"), 0, -1)
                        if instr(instrMedidProgrp, "#"& session("mid") &"#,") <> 0 then
                        session("dontLogind") = 1
                        else
                        session("dontLogind") = 0
                        end if

                
              
                        
                case else
                session("dontLogind") = 0
                end select


                '**** Log ikke ind automatisk ****'
                
               

                call erStempelurOn()
               
                if cint(stempelur_hideloginOn) = 1 then 'skriv ikke login, men åben komme/gå tom
				    
                else

                    if session("dontLogind") <> 1 then
				    '**** Tjekker om der skla laves en auto-logud eller om der er et igangværende logind på dagen ***'    
				    call logindStatus(strUsrId, intStempelur, 1, now)
				    end if

                end if
                'end select

                'response.end
				
				

                '******************* Minimumslager notificering ******''
                if session("rettigheder") = 1 then

                    'select case lto
                    'case "dencker", "intranet - local"

                    call minimumslageremail_fn()

                        if cint(minimumslageremail) = 1 then

                            if datePart("w", now, 2,2) = 4 then 'torsdag / OR datePart("w", now, 2,2) = 3  onsdag
                                call matUnderMinLager(lto)
                            end if

                        end if
			
                    'end select	

                end if


				'Response.End
		
		else
		oConn.execute("UPDATE kontaktpers SET lastlogin = '"& strLastLogin &"' WHERE id="& thisKpid  &"")
		end if
		'***********************************'
		'***** Medarbejder kunde login  END '
		
		
        '********************* Sebnder mail hvis det er demo **************************************
        if lto = "demo" then

        email = request("FM_email")
        firma = request("FM_firma")
        ref = request("FM_ref")
        tel = request("FM_tel")

                                    

                                    if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then
                                    

                                    Set myMail=CreateObject("CDO.Message")
                                    myMail.Subject = "Der har været en ny potentiel kunde i demoen"
                                    myMail.From = "timeout_no_reply@outzource.dk"

                                    myMail.To = "<osn@outzource.dk>; <salg@outzource.dk>; <sk@outzource.dk>"                        
                    
                                    if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                    to_url = "https://outzource.dk"
                                    else
                                    to_url = "https://timeout.cloud"
                                    end if
                        
                                    myMail.TextBody = "Hej OutZourCE" & vbCrLf & vbCrLf _
                                    & "Der har været en ny potentiel kunde i demoen : "& vbCrLf & vbCrLf _
                                    & "Firma: "& firma & vbCrLf _
						            & "Navn: "& ref & vbCrLf & vbCrLf _
                                    & "Email: "& email &""& vbCrLf & vbCrLf _
                                    & "Tel: "& tel &""& vbCrLf & vbCrLf _
		                            & "Med venlig hilsen" & vbCrLf _
		                            & session("user") & vbCrLf & vbCrLf
                        
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

                                    if len(trim(email)) <> 0 then
                                    myMail.Send
                                    end if
                                    set myMail=nothing

                                    end if
                
                                    'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		                            ' Sætter Charsettet til ISO-8859-1
		                            'Mailer.CharSet = 2
		                            'Mailer.FromName = "TimeOut | " & ref &"("& firma &")" 
		                            'Mailer.FromAddress = email
		                            'Mailer.RemoteHost = "webout.smtp.nu" '195.242.131.254 '"webmail.abusiness.dk" '"pasmtp.tele.dk"
						            'Mailer.AddRecipient "Søren Karlsen", "sk@outzource.dk"
                                    'Mailer.AddRecipient "Outzource Salg", "salg@outzource.dk"
		                          

                                 
            					    'Mailer.Subject = "Der har været en ny potentiel kunde i demoen"
		                            'strBody = "Hej OutZourCE" & vbCrLf & vbCrLf
                                    'strBody = strBody &"Der har været en ny potentiel kunde i demoen : "& vbCrLf & vbCrLf
                                    'strBody = strBody &"Firma: "& firma & vbCrLf 
						            'strBody = strBody &"Navn: "& ref & vbCrLf & vbCrLf
                                    'strBody = strBody &"Email: "& email &""& vbCrLf & vbCrLf
                                    'strBody = strBody &"Tel: "& tel &""& vbCrLf & vbCrLf
		                            'strBody = strBody &"Med venlig hilsen" & vbCrLf
		                            'strBody = strBody & session("user") & vbCrLf & vbCrLf
            		                
            		
            		
		                            'Mailer.BodyText = strBody
            		
		                            'Mailer.sendmail()
		                            'Set Mailer = Nothing


                                    'end if
        
        
        end if


			
		
		'*********** redirect til den valgte side *************************************************
		'Response.Cookies("usrname")("ansatkunde") = request("FM_ansat_kunde")
		
            '** Den skal nulstilles , så man altid starter med tom liste. PGA performance EPI
		response.cookies("tsa")("visAlleMedarb") = ""

		Response.Cookies("loginlto")("lto") = lto
		Response.Cookies("loginlto").expires = date + 90
       
		
		if len(request("huskmig")) <> 0 then
		Response.Cookies("login")("usrval") = request("login")
		Response.Cookies("login")("pwval") = request("pw")
		Response.Cookies("login").expires = date + 45
		else
		Response.Cookies("login").expires = date - 1
		end if
		
		
		
		
		'*** nulstiller antal logind forsøg, eller kan man ikke logge ind igen
		session("spmettanigol") = 0

		'Response.Cookies("loginlto")("lto") = lto
			'*** tjekker om det er PDA eller PC ****

   
    'if lto = "outz" OR lto = "hestia" then
            call browsertype()
            browstype_client = browstype_client
    'else
     '       browstype_client = "xx"
    'end if
 

			if browstype_client = "ip" then
			
			    'New server
                if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                response.redirect "http://outzource.dk/timeout_xp/wwwroot/ver2_14/timetag_web/timetag_web.asp?flushsessionuser=1"
                else
                response.redirect "https://timeout.cloud/timeout_xp/wwwroot/ver2_14/timetag_web/timetag_web.asp?flushsessionuser=1"
                end if


			else
				
				'*********************************************************
				'**** Tjekker om det er localhost (udviklingsmaskine) ****
				'**** Hvis nej: SSL 								  ****
				'*********************************************************
				'if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then
				'	if stransatkunde = "1" then
				'		select case startside 
				'		case 1 
				'			if lto = "kasters" then
				'			response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/kunder.asp?menu=crm"
				'			else
				'			response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal"
				'			end if
				'		case 2
				'		response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/jbpla_w.asp?menu=res"
				'		case else '0
				'			'if lto = "bowe" then
				'			'Response.write "treg0206" & treg0206
				'			if cint(treg0206) = 1 then
				'			response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/timereg_2006_fs.asp"
				'			'response.redirect "http://81.19.249.35/timeout_xp/wwwroot/ver2_1/timereg/timereg_2006_fs.asp"
				'			else
				'			response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/timereg.asp?showstempelur=1"
				'			end if
				'		end select
				'	else
				'		response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/joblog_k.asp"
				'	end if
				'else

                    '***** KUNDE/EKSTERN *****'
					if stransatkunde = "1" then

                    if lto = "demo" then
                    
                                response.redirect "to_2015/home_dashboard_demo.asp"
                    else

						        select case startside 
						        case 1 
						            if lto = "essens" OR lto = "fkfk" then
						            response.redirect "timereg/crmhistorik.asp?menu=crm&ketype=e&func=hist&selpkt=hist"
						            else
						            response.redirect "timereg/crmkalender.asp?menu=crm&shokselector=1&ketype=e&selpkt=kal"
						            end if
						        case 2
						        response.redirect "timereg/ressource_belaeg_jbpla.asp?menu=webblik"
                                case 3

                                call sonmaniuge(now)

                                'Response.Write "timereg/logindhist_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge
                                'Response.end
                                response.redirect "timereg/logindhist_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge
                                'response.redirect "timereg/logindhist_2011.asp"    

                                case 4
						        response.redirect "timereg/webblik_joblisten.asp?menu=webblik&fromvmenu=1"
                                case 5

                                    if lto = "nt" OR lto = "intranet - local" then
                                    response.redirect "to_2015/job_nt.asp"
                                    else
						            response.redirect "timereg/jobs.asp?menu=job&shokselector=1&fromvemenu=j"
                                    end if
                    
                                case 6



                                call sonmaniuge(now)


                                'Response.Write "timereg/ugeseddel_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge
                                'Response.end
                                'response.redirect "timereg/ugeseddel_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge
                                response.redirect "to_2015/ugeseddel_2011.asp?usemrn="&session("mid")&"&varTjDatoUS_man="&mandagIuge&"&varTjDatoUS_son="&sondagIuge
                                
                                case 7 ' DASHBOARD
						        response.redirect "to_2015/medarbdashboard.asp"

                                case 8 ' Kunder
						        response.redirect "to_2015/kunder.asp?visikkekunder=1"

                                case 9 ' Favorit
                                call sonmaniuge(now)
						        response.redirect "to_2015/favorit.asp?varTjDatoUS_man="&mandagIuge
                                
                                case else '0



							            if cint(treg0206) = 1 then
                            
                            

                                            'Response.write timeout_version
                                            'Response.end

                                            if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\login.asp" then

                                            
                                            select case timeout_version
                                            case "ver2_1"
							                response.redirect "timereg/timereg_akt_2006.asp"
                                            case "ver3_99"
                                            response.redirect "../ver3_99/timereg/timereg_akt_2006.asp"
                                            case "ver2_10"
                                            'response.redirect "../ver2_10/timereg/timereg_akt_2006.asp?hideallbut_first=1"
                                            response.redirect "../ver2_10/timereg/timereg_akt_2006.asp"
                                            case "ver2_14"
                                            'response.redirect "../ver2_10/timereg/timereg_akt_2006.asp?hideallbut_first=1"
                                            response.redirect "../ver2_14/timereg/timereg_akt_2006.asp"
                                            case else
                                            'response.redirect "../ver2_10/timereg/timereg_akt_2006.asp?hideallbut_first=1"
                                            response.redirect "../ver2_14/timereg/timereg_akt_2006.asp"
                                            end select
							                'response.redirect "timereg/timereg_2006_fs.asp"

                                        else

                                        'select case lto
                                        'case "dencker", "intranet - local"
                                        'response.redirect "http://localhost/timeout_xp/timereg/timereg_akt_2006.asp"
                                        response.redirect "http://localhost/timereg/timereg_akt_2006.asp"
                                        'case else
                                        'response.redirect "http://localhost/timeout_xp/timereg/timereg_akt_2006.asp?hideallbut_first=1"
                                        'end select

                                        end if


							            else
							            response.redirect "timereg/timereg.asp?showstempelur=1"
							            end if
						        end select
                    end if 'demo
					else
						response.redirect "timereg/joblog_k.asp"
					end if
				'end if
				
				
			end if
		'********************************************************************************************
		
		
		
		else 'founduser
		%>

		<%
		errortype = 4
		call showError(errortype)
		
		End if
		
		
		
		
		
		'call closeDB
                            end if '* Validering '
		                end if '* Validering '
                     end if '* Validering '
                    end if '* Validering '
		End if '* Validering '
		End if '* Validering '
	end if '* Validering '
	

end if



sub stempelurLoginSetting

             session("stempelur") = 0
			    strSQL = "SELECT stempelur FROM licens WHERE id = 1"
			    oRec.open strSQL, oConn, 3 
			    if not oRec.EOF then
			
				    session("stempelur") = oRec("stempelur")
			
			    end if
			    oRec.close  
			
			    if session("stempelur") <> 0 then


                if browstype_client <> "ip" then
                %>
                    

                    <div class="form-group clearfix">
                          <div class="pull-left">
                    <%=login_txt_021 %><br />
				    <label class="radio-inline" style="text-align:left">

                <%
                    else
                    %>
                       
                           <div class="pull-left clearfix">
                                 <%=login_txt_021 %><br />
                          
                              
                        <%
                end if
                
                strSQL = "SELECT id, navn, faktor, forvalgt FROM stempelur ORDER BY navn"
				oRec.open strSQL, oConn, 3 
				st = 0
                while not oRec.EOF 
				
				if oRec("forvalgt") = 1 then
				chk = "CHECKED"
				else
				chk = ""
				end if


                %>
                
                        <input type="radio" name="FM_stempelur" value="<%=oRec("id")%>" <%=chk%>> <%=oRec("navn")%>
                               <%  if browstype_client <> "ip" then %>
                               <br />                       
                               <%else %>
                               &nbsp;&nbsp;
                               <%end if %>
				
				<%

                st = st + 1
				oRec.movenext
				wend
				oRec.close 
                
                if browstype_client <> "ip" then
                %>
                </label>
                </div>
                </div>
                <%
                else
                    %>
    <br /> &nbsp;<br />
                    </div>
                <%    
                end if
			
			    else
			    %>
			        <div class="pull-right"><input type="hidden" name="FM_stempelur" id="FM_stempelur" value="0"></div>
			    <%
			    end if


end sub

%>


<!--#include file="inc/regular/footer_login_inc.asp"-->


