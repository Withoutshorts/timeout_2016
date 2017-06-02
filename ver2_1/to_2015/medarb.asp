


<%response.buffer = true 
Session.LCID = 1030
%>
			        

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
   
    func = request("func")

	if func = "dbopr" then
	id = 0
	else
        if len(trim(request("id"))) <> 0 then
	    id = request("id")
        else
        id = 0
        end if
	end if    
    
    media = request("media")
   

    if media <> "eksport" then


    if media <> "print" then    
        call menu_2014
    end if


     %>
    <div id="wrapper">
        <div class="content">

    <%

    end if

    select case func
    case "slet"
    %>
            <!--slet sidens indhold-->
        <div class="container" style="width:500px;">
            <div class="porlet">
            
            <h3 class="portlet-title">
               <u><%=medarb_txt_002 %></u>
            </h3>
            
                <div class="portlet-body">
                    <div style="text-align:center;"><%=medarb_txt_003 %> <b><%=medarb_txt_004 %></b> <%=medarb_txt_005 %><br />
                        <%=medarb_txt_006 %>
                    </div><br />
                   <div style="text-align:center;"><a class="btn btn-primary btn-sm" role="button" href="medarb.asp?menu=tok&func=sletok&id=<%=id%>">&nbsp <%=medarb_txt_007 %> &nbsp</a>&nbsp&nbsp&nbsp&nbsp<a class="btn btn-default btn-sm" role="button" href="medarb.asp?"><%=medarb_txt_008 %></a>
                    </div>
                    <br /><br />
                 </div>

            </div>
        </div>
    
      <%
         case "sletok"
        '*** Her slettes en medarbejder ***

	oConn.execute("DELETE FROM budget_medarb_rel WHERE medid = "& id &"")
	oConn.execute("DELETE FROM medarbejdere WHERE Mid = "& id &"")
	oConn.execute("DELETE FROM progrupperelationer WHERE MedarbejderId = "& id  &"") 'projektgruppeId = 10 AND
	oConn.execute("DELETE FROM timereg_usejob WHERE medarb = "& id  &"") 
    oConn.execute("DELETE FROM timer WHERE tmnr = "& id &"")
	oConn.execute("DELETE FROM ressourcer WHERE mid = "& id &"")
	oConn.execute("DELETE FROM ressourcer_md WHERE medid = "& id &"")
	oConn.execute("DELETE FROM timepriser WHERE medarbid = "& id &"")
	
    	
    Response.Redirect "medarb.asp?menu=medarbejder&lastmedid="&id
                     
					
					
    

    
    case "dbred", "dbopr" 


	'*** Her tjekkes om alle required felter er udfyldt. ***
      


	if len(trim(Request("FM_login"))) = 0 OR len(trim(Request("FM_pw"))) = 0 OR len(trim(Request("FM_mnr"))) = 0 OR len(trim(Request("FM_init"))) = 0 OR len(trim(request("FM_navn"))) = 0 then 
	
	
	errortype = 9
	call showError(errortype)

       
    response.end
	
	else
	        %>
			<!--#include file="../timereg/inc/isint_func.asp"-->
			<%
		    isInt = 0
	        call erDetInt(trim(request("FM_mnr")))

                'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br>SSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS #"& request("FM_mnr") & "# isInt: "& isInt &"<br>"
                'Response.write 
                'Response.end


			if isInt > 0 then
			
			errortype = 22
			call showError(errortype)

           
            response.end
			
			isInt = 0
			else
	        
	        
	        if len(request("ansatdato")) <> 0 AND isDate(request("ansatdato")) then
			ansatdato = request("ansatdato")
			ansatdato = year(ansatdato)&"/"&month(ansatdato)&"/"&day(ansatdato)
			else
		

			errortype = 114
			call showError(errortype)
			Response.end
			end if	
			
			if len(request("opsagtdato")) <> 0 AND isDate(request("opsagtdato")) then
			opsagtdato = request("opsagtdato")
			opsagtdato = year(opsagtdato)&"/"&month(opsagtdato)&"/"&day(opsagtdato)
			else
			

			errortype = 115
			call showError(errortype)
			Response.end
			end if	
			
			
		
		
		
	    
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function
		
			strNavn = SQLBless(Request("FM_navn"))
			strAnsat = Request("FM_ansat")
			
			strInit = SQLBless(request("FM_init"))
			
			if len(request("FM_timereg")) <> 0 then
			intTimereg = request("FM_timereg") 
			else
			intTimereg = 0
			end if
			
			
			
			strLogin = Request("FM_login")
			strPw = Request("FM_pw")
			strMnr = trim(Request("FM_mnr"))
			strEditor = session("user")
			strDato = session("dato")
			strMedinfo = SQLBless(Request("FM_medinfo"))
			strBrugergruppe = Request("FM_bgruppe")

            if len(trim(Request("FM_medtype"))) <> 0 then
			strMedarbejdertype = Request("FM_medtype")
            else
            strMedarbejdertype = 0
            end if

			strEmail = Request("FM_email")
			
			madr = replace(request("FM_adr"), "'", "''")
			mpostnr = replace(request("FM_postnr"), "'", "''")
			mcity = replace(request("FM_city"), "'", "''")
			mland = request("FM_land")
			mtlf = replace(request("FM_tel"), "'", "''")
			mcpr = replace(request("FM_cpr"), "'", "''")
			mkoregnr = replace(request("FM_regnr"), "'", "''")
			
			'*** Kun admin brugere skal sættes til at modtage nyhedsbrev ved opret. ***'
			if func = "dbred" then
			nyhedsbrev = request("nyhedsbrev")
			else
			    if cint(strBrugergruppe) = 3 then
			    nyhedsbrev = 1
			    else
			    nyhedsbrev = 0
			    end if
			end if
			
			dagsdato = year(now) &"/"& month(now) &"/"& day(now)
			
			if len(request("FM_tsacrm")) <> 0 then 
			inttsacrm = request("FM_tsacrm")
			else
			inttsacrm = 0
			end if
			
			if len(request("FM_smiley")) <> 0 then
			intSmil = request("FM_smiley")
			else
			intSmil = 0
			end if
			
			strExch = replace(request("FM_exch"), "\", "#")
			
			sprog = request("FM_sprog")


            '**************** Change version
            if len(trim(request("FM_visskiftversion"))) <> 0 AND request("FM_visskiftversion") <> 0 then
            visskiftversion = request("FM_visskiftversion")
            else
            visskiftversion = 0
            end if

            '***************** Create user ok ****'
             if len(trim(request("FM_opretmedarb"))) <> 0 AND request("FM_opretmedarb") <> 0 then
            create_newemployee = request("FM_opretmedarb")
            else
            create_newemployee = 0
            end if
                 

			
            '** Vis timer som start / stop tid
            if len(trim(request("FM_timer_ststop"))) <> 0 then
            timer_ststop = 1
            else
            timer_ststop = 0
            end if

                
			
			'*** Her tjekkes for dubletter i db ***
			strSQL = "SELECT Mnr FROM medarbejdere WHERE Mnr ="& strMnr &" AND Mid <> "& id
			oRec.open strSQL, oConn, 3	
			if oRec.EOF then
			strMnrOK = "y"
			end if
			oRec.close
			
			strSQL = "SELECT login FROM medarbejdere WHERE login ='"& strLogin &"' AND Mid <> "& id
			oRec.open strSQL, oConn, 3	
			if oRec.EOF then
			strLoginOK = "y"
			end if
			oRec.close
			
			'Checker for dubletter i DB før record opdateres eller indsættes
			if strMnrOK = "y" AND strLoginOK = "y" then
			        

                '** Finder medarbejdertypegrp (hvis findes) **'
                    '** Gemmes på medarbejder pga perfomance på statistikker
                    intMedarbejdertype_grp = 0
                    strSQlmtyp = "SELECT mgruppe FROM medarbejdertyper WHERE id = "& strMedarbejdertype
                    oRec2.open strSQlmtyp, oConn, 3
                    if not oRec2.EOF then
                    intMedarbejdertype_grp = oRec2("mgruppe")
                    end if
                    oRec2.close
			        '****'
			        
			        
			        
			        if func = "dbred" then
			        
			        
			        '*** Medarbejdertype historik ***'
			        nyMedarbType = 0
			        strSQLmth = "SELECT medarbejdertype FROM medarbejdere WHERE mid = "& id 
			        oRec.open strSQLmth, oConn, 3
			        if not oRec.EOF then
			        
			            if cint(oRec("medarbejdertype")) <> cint(strMedarbejdertype) then
			            
			                    
			                    '*** Er det første ændring? ***'
			                    
			                    typeHist = 0
			                    strSQL1 = " SELECT mid FROM medarbejdertyper_historik WHERE mid = "& id
			                    oRec2.open strSQL1, oConn, 3
			                    if not oRec2.EOF then
			                        
			                        typeHist = 1
			                    
			                    end if
			                    oRec2.close
			                    
			                    
			                    '*** Opretter den forladte type ved første sktift,  **'
			                    '*** så man ved hvilken type medarbejder            **'
			                    '*** har været fra ansættelses dato                 **'
			                    '*** til første skift.                              **'
			                    
			                    if typeHist = 0 then
			                    
			                    strSQLmthupd1 = "INSERT INTO medarbejdertyper_historik (mid, mtype, mtypedato) "_
			                    &" VALUES ("& id &", "& oRec("medarbejdertype") &", '"& ansatdato &"') "
			                    oConn.execute(strSQLmthupd1)
			                    
			                    
			                    end if
			                    
			                    
			            
			            strSQLmthupd3 = "INSERT INTO medarbejdertyper_historik (mid, mtype, mtypedato) "_
			            &" VALUES ("& id &", "& strMedarbejdertype &", '"& dagsdato &"') "
			            
			            oConn.execute(strSQLmthupd3)
			            
			            end if

                    nyMedarbType = oRec("medarbejdertype")
			        
			        end if
			        oRec.close
			        
			        
			        
			                   
			        
			        

                    '** Opdaterer datoer på eksisterende ****'
                    mtyphist_ids = split(request("FM_mtyphist_id"), ", ")
                    mtyphist_datoer = split(request("FM_mtyphist_dato"), ", ")

                        
                    useOpdStDatoTp = year(now) & "-" & month(now) & "-" & day(now)
                    useOpdMtypeGrp = 0
                    aktiveI = 0

                    for i = 0 to UBOUND(mtyphist_ids)
                                    

                                    if isDate(mtyphist_datoer(i)) = true then

                                    mtyphist_datoer(i) = year(mtyphist_datoer(i)) & "-"& month(mtyphist_datoer(i)) & "-" & day(mtyphist_datoer(i))
                                    strSQLmthupd2 = "UPDATE medarbejdertyper_historik SET mtypedato = '"& mtyphist_datoer(i) &"' WHERE id = "& mtyphist_ids(i)
                                    'Response.write strSQLmthupd2
                                    'Response.end
			                        oConn.execute(strSQLmthupd2)

                                    end if
                        
                                aktiveI = i
                                
                             

                    next
                            
                            
                                  
                                
                                       
                
                                '*** Kost og timepriser opdatering ***********'
                                '*** SKal eksisterede kostprsier opdateres ***'
                                if len(trim(request("FM_opd_mty_kp"))) <> 0 then

                                useOpdStDatoTp = mtyphist_datoer(aktiveI)
                                useOpdMtypeGrp = nyMedarbType
                        
                                nyKost = 0
                                nyKurs = 100
                                nyValuta = 1

                                strSQLsmtyKp = "SELECT m.kostpris FROM medarbejdertyper AS m WHERE m.id = "& useOpdMtypeGrp
                            
                                'response.write "strSQLsmtyKp: "& strSQLsmtyKp
                                'response.flush 

                                oRec6.open strSQLsmtyKp, oConn, 3 
                                if not oRec6.EOF then
                            
                                nyKost = replace(oRec6("kostpris"), ",", ".") 
                                        
                                end if
                                oRec6.close

                                strSQLupdKp = "UPDATE timer SET kostpris = "& nyKost &" WHERE tmnr = "& id &" AND tDato >= '"& useOpdStDatoTp &"'"

                                'response.write "<br>strSQLupdKp: "& strSQLupdKp
                                'response.end

                                oConn.execute(strSQLupdKp)        

                                end if 
                
                
                                '*** SKal eksisterede timepriser opdateres ***'
                                if len(trim(request("FM_opd_mty_tp"))) <> 0 then
                                
                                useOpdStDatoTp = mtyphist_datoer(aktiveI)
                                useOpdMtypeGrp = nyMedarbType


                                nyTp = 0
                                nyKurs = 100
                                nyValuta = 1

                                strSQLsmtyKp = "SELECT m.timepris, m.tp0_valuta, v.kurs FROM medarbejdertyper AS m "_
                                &"LEFT JOIN valutaer AS v ON (v.id = tp0_valuta) WHERE m.id = "& useOpdMtypeGrp
                            
                                'response.write "strSQLsmtyKp: "& strSQLsmtyKp
                                'response.flush 

                                oRec6.open strSQLsmtyKp, oConn, 3 
                                if not oRec6.EOF then
                            
                                nyTp = replace(oRec6("timepris"), ",", ".") 
                                nyValuta = oRec6("tp0_valuta") 
                                nyKurs = replace(oRec6("kurs"), ",", ".")           

                                end if
                                oRec6.close

                                strSQLupdTp = "UPDATE timer SET timepris = "& nyTp &", valuta = "& nyValuta &", kurs = "& nyKurs &" WHERE tmnr = "& id &" AND tDato >= '"& useOpdStDatoTp &"'"
                            
                                'if session("mid") = 1 then
                                'response.write "<br>strSQLupdTp: "& strSQLupdTp
                                'response.end
                                'end if

                                oConn.execute(strSQLupdTp)      

                                
                
                                end if            
                                '**********************************************************************************
                                



                       
                      'if session("mid") = 1 then
                      'response.end
                      'end if









                    '*** dato skal altid følge ansatdato på første type **'
			        strSQLmthupd1 = "SELECT id FROM medarbejdertyper_historik WHERE mid =  "& id &" ORDER BY id"
			                    
			        oRec2.open strSQLmthupd1, oConn, 3
			        if not oRec2.EOF then
			                        
			                strSQLmthupd2 = "UPDATE medarbejdertyper_historik SET mtypedato = '"& ansatdato &"' WHERE id = "& oRec2("id")
			            oConn.execute(strSQLmthupd2)
                                    
			                    
			        end if
			        oRec2.close

                                

                               
                            

                    
			        
			        
                    '*** Hvis medarbejder DE-aktiveres Sætter opsigelsesdato til dd. PGHA normtids beregning ***'
                    'Response.write "strAnsat" & strAnsat
                    if strAnsat = "2" then
                    
                    'Response.write "her"
                    useToBeAnsat = "0"
                    strSQLtjk = "SELECT mansat FROM medarbejdere WHERE mid = "& id
                    oRec5.open strSQLtjk, oConn, 3
                    if not oRec5.EOF then

                    useToBeAnsat = oRec5("mansat")
                    'Response.write "<br>useToBeAnsat"& useToBeAnsat 

                    end if
                    oRec5.close
                    

                    if len(trim(request("FM_deakdd_cancel"))) <> 0 then
                    deakdd_cancel = 1
                    else
                    deakdd_cancel = 0
                    end if

                    if useToBeAnsat <> "2" AND deakdd_cancel = 0 then
                    opsagtdato = year(now)&"/"&month(now) &"/"& day(now)
                    end if
                    
                    
                    
                    end if
                    '*** end 

                   
                    'Response.end
                    


					
					strSQL = "UPDATE medarbejdere SET"_
					&" Mnavn = '"& strNavn &"',"_
					&" Mnr = "& strMnr &","_
					&" Mansat = '"& strAnsat &"',"_
					&" login = '"& strLogin &"',"
					
					if strPw <> "KEEPTHISPW99" then
					strSQL = strSQL &" pw = MD5('"& strPw &"'),"
					end if
					
					strSQL = strSQL &" brugergruppe = "& strBrugergruppe &","_
					&" Medarbejdertype = "& strMedarbejdertype &","_
					&" editor = '"& strEditor &"',"_
					&" dato = '"& strDato &"',"_
					&" Medarbejderinfo = '"& strMedinfo &"', "_
					&" Email = '"& strEmail &"',"_
					&" tsacrm = "& inttsacrm &", smilord = "& intSmil &", exchkonto = '"& strExch &"', "_
					&" init = '"& strInit &"', timereg = "& intTimereg &","_
					&" ansatdato = '"& ansatdato &"', "_
					&" opsagtdato = '"& opsagtdato &"', sprog = "& sprog &", nyhedsbrev = "& nyhedsbrev &", "_
					&" madr = '"& madr & "', mpostnr = '"& mpostnr &"', mcity = '"& mcity &"', mland = '"& mland &"', "_
					&" mtlf = '"& mtlf &"', mcpr = '"& mcpr &"', mkoregnr = '"& mkoregnr &"', "_
                    &" visskiftversion = "& visskiftversion &", medarbejdertype_grp = "& intMedarbejdertype_grp &", timer_ststop = "& timer_ststop &", create_newemployee = "& create_newemployee &""_
			        &" WHERE Mid = "& id &""
					
					
					'response.flush 
                    'response.write strSQL

                    oConn.execute(strSQL)
                    
					if cint(level) = 1 then 'Kun Admin må ændre projektgrupper ved rediger medarb.
					    '*** projektgruppe relationer ***'
					    call prgrel(id, func)
                    end if

                    'Response.end
					
					
					'**** timereg_usejob ****'
                    del = 1 '1 slet fra timereg_usejob / 0: tilføj kun
					call tilfojtilTU(id, del) 
					
					
					
					'*** Når der ændres PW ****'
					'*** Sender mail til medarb. at profil er blevet opdateret ***'

                  	
                    if cint(strAnsat) = 1 then 'kun mail til aktive brugere

					if strPw <> "KEEPTHISPW99" then
					
					if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\medarb_red.asp" then

                        Set myMail=CreateObject("CDO.Message")
                        myMail.Subject="TimeOut - Medarbejder profil opdateret"
                        myMail.From="timeout_no_reply@outzource.dk"

                        if len(trim(strEmail)) <> 0 then
                        myMail.To= ""& strNavn &"<"& strEmail &">"
                        end if

                        if lto = "mi" then
                        myMail.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\help_and_faq\TimeOut_indtasttimer_rev_20120106.pdf" 
                        else
                        myMail.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\help_and_faq\TimeOut_indtasttimer_rev_20130102.pdf" 
                        end if
                    
                        if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                        to_url = "https://outzource.dk"
                        else
                        to_url = "https://timeout.cloud"
                        end if
                        
                        myMail.TextBody= "" & medarb_txt_009 &" "& strNavn & vbCrLf _ 
					    & medarb_txt_118 & vbCrLf _ 
					    & medarb_txt_012 &" "& strLogin & medarb_txt_013 &" "& strPw & vbCrLf & vbCrLf _ 
					    & medarb_txt_014  & vbCrLf _ 
					    & medarb_txt_015 & vbCrLf & vbCrLf _ 
					    & medarb_txt_119&" "&to_url&"/"&lto&""& vbCrLf & vbCrLf _ 
					    & medarb_txt_120& vbCrLf & vbCrLf & strEditor & vbCrLf 

                        
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
					
					end if '** Ændre PW **'

                    end if '** bruger aktiv

					
                  
  
	                
					
                    
					
					end if
					
					
					
					
					
		        '*** Opret medarbejer i DB **'
			    if func = "dbopr" then
    					
    					
                        
    					
    					
						    strSQLminsert = "INSERT INTO medarbejdere"_
						    &" (Mnavn, Mnr, Mansat, login, pw, Brugergruppe, Medarbejdertype, "_
						    &" editor, dato, Medarbejderinfo, Email, tsacrm, smilord, exchkonto, init, "_
						    &" timereg, ansatdato, opsagtdato, sprog, nyhedsbrev, "_
						    &" madr, mpostnr, mcity, mland, mtlf, mcpr, mkoregnr, visskiftversion, medarbejdertype_grp, timer_ststop, create_newemployee) VALUES ("_
						    &" '"& strNavn &"',"_
						    &" "& strMnr &","_
						    &" '"& strAnsat &"',"_
						    &" '"& strLogin &"',"_
						    &" MD5('"& strPw &"'),"_
						    &" "& strBrugergruppe &","_
						    &" "& strMedarbejdertype &","_
						    &" '"& strEditor &"',"_
						    &" '"& strDato &"',"_
						    &" '"& strMedinfo &"',"_
						    &" '"& strEmail &"',"_
						    &" "& inttsacrm &", "& intSmil &", "_
						    &" '"& strExch &"', '"& strInit &"',"_
						    &" "& intTimereg &","_
						    &" '"& ansatdato &"', '"& opsagtdato &"', "& sprog &", "_
						    &" "& nyhedsbrev &", "_
						    &" '"& madr & "', '"& mpostnr &"', '"& mcity &"', '"& mland &"', "_
					        &" '"& mtlf &"', '"& mcpr &"', '"& mkoregnr &"', "& visskiftversion &", "& intMedarbejdertype_grp &", "& timer_ststop &", "& create_newemployee &")"
    						
                            'Response.write strSQLminsert
                            'Response.flush
                            oConn.execute(strSQLminsert)
						    strSQL = "SELECT Mid FROM medarbejdere WHERE Mnr = " & strMnr &""
						    oRec.open strSQL, oConn, 3
						    if not oRec.EOF then
							    intMid = oRec("Mid")
						    end if
						    oRec.close
    						
						    '*** Medarbejder  type historik ****
						    strSQLmthupd = "INSERT INTO medarbejdertyper_historik (mid, mtype, mtypedato) "_
			                &" VALUES ("& intMid &", "& strMedarbejdertype &", '"& ansatdato &"') "
			                oConn.execute(strSQLmthupd)
    						
						    
                          
    						'*** projektgruppe relationer ***'
					        call prgrel(intMid, func)


					        '**** Sætter guiden aktive job aktive ****'
					        'guideasyids = 0
                            'forvalgt = 0 'off / 1 = on

                            'Response.Write "guidjobids:"& guidjobids

                            'Response.end
					        'call setGuidenUsejob(intMid, guidjobids, 1, guideasyids, forvalgt, 1)

                            '**** timereg_usejob ****'
                            call positiv_aktivering_akt_fn() 'wwf
                            if cint(positiv_aktivering_akt_val) <> 1 then
					            call tilfojtilTU(intMid, 0) 
					        end if
    						
						    '*** Opdaterer timepriser på stamaktiviteter ***
    						
						    medtypeid = 0
						    strSQL = "SELECT mid FROM medarbejdere WHERE Medarbejdertype = "& strMedarbejdertype & " AND mid <> " & intMid
						    oRec.open strSQL, oConn, 3 
						    while not oRec.EOF 
						    medtypeid = oRec("mid")
						    oRec.movenext
						    wend
						    oRec.close 

                            '**** HVIS der ikke er andre medarbejdere af typen bruges medarbejdertypens standard timepriser
                             if cdbl(medtypeid) = 0 then
                                if len(trim(request("FM_medarbejdertype_follow_tp"))) <> 0 then
                                 medarbejdertype_follow_tp = request("FM_medarbejdertype_follow_tp")
                                else
                                 medarbejdertype_follow_tp = 0
                                end if

                                        '*** Vælger 1 medarb. af typen
                                        strSQL = "SELECT mid FROM medarbejdere WHERE Medarbejdertype = "& medarbejdertype_follow_tp & " LIMIT 1"
						                oRec.open strSQL, oConn, 3 
						                if not oRec.EOF then
						                medtypeid = oRec("mid")
						                oRec.movenext
						                end if
						                oRec.close 

                             end if

                            ct = 0
                            strAktIDsWrt = ""
                          
                            
                                    '************** HENTER ÅBEN JOB OG TILBUD ***********
                                    strJobids = "(jobid = 0 "
                                    whSQL = "(jobstatus = 1 OR jobstatus = 3)"
                                    strSQLjob = "SELECT j.id AS jid FROM job j "_
		                            &" WHERE "& whSQL &" GROUP BY j.id"

                                    oRec2.open strSQLjob, oConn, 3
		                            while not oRec2.EOF
                 
                                        strJobids = strJobids & " OR jobid = " & oRec2("jid")

                                    oRec2.movenext
                                    wend 
                                    oRec2.close

                       
                                    strJobids = strJobids & ")"

                            strSQL = "SELECT aktid, jobid, timeprisalt, 6timepris, 6valuta FROM timepriser WHERE "& strJobids &" AND medarbid = "& medtypeid &" GROUP BY aktid, timeprisalt" 
						  
                            'Response.write strSQL & "<br>"
    						
						    oRec.open strSQL, oConn, 3 
						    while not oRec.EOF 

                                tprisUse = replace(tprisUse, ".", "")
                                tprisUse = replace(oRec("6timepris"), ",", ".")

                               
                                timeprisAlt = oRec("timeprisalt") 'kolononne valgt
                               
							    strSQLinsTp = "INSERT INTO timepriser ( "_
							    &" jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
							    &" VALUES ("& oRec("jobid") &", "& oRec("aktid") &", "& intMid &", "& timeprisAlt &", "& tprisUse &", "& oRec("6valuta") &")"

                                'Response.Write strSQLinsTp & "<br>"
                                'Response.flush

                                oConn.execute(strSQLinsTp)
    						    


                                strAktIDsWrt = strAktIDsWrt & " AND id <> " & oRec("aktid")
                                ct = ct + 1
						    oRec.movenext
						    wend
						    oRec.close 
    					
                        'response.end	
    					'Response.write strAktIDsWrt & "<br><br> "
                        

                        '*** Indlæser default timepris på de STAM-aktiviteter der EVT ikke bliver fundet *****'
                        '*** finder tp på type ***
                        tprisUse = 0
                        valutaUse = 1
                        strSQLmtyp = "SELECT timepris, tp0_valuta FROM medarbejdertyper WHERE id = " & strMedarbejdertype

                        oRec.open strSQLmtyp, oConn, 3 
						if not oRec.EOF then

                          tprisUse = replace(oRec("timepris"), ".", "")
                          tprisUse = replace(tprisUse, ",", ".")

                          valutaUse = oRec("tp0_valuta")

                        end if
                        oRec.close

                        strSQLstamAktiF = "SELECT id FROM aktiviteter WHERE job = 0 AND aktfavorit <> 0 " & strAktIDsWrt
                        'Response.Write strSQLstamAktiF
                        'Response.flush

                        oRec.open strSQLstamAktiF, oConn, 3 
						while not oRec.EOF 

                        strSQLinsTpAktiF = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris, 6valuta) "_
                        &" VALUES (0, "& oRec("id") &", "& intMid &", 6, "& tprisUse &", "&valutaUse&")" 

                        'Response.Write strSQLinsTpAktiF
                        'Response.flush

                        oConn.execute(strSQLinsTpAktiF)

                        oRec.movenext
						wend
						oRec.close 

                        'Response.Write "<br>ct:"& ct
                        'Response.end

    					lastmedid = intMid	





                         if cint(strAnsat) = 1 then 'kun mail til aktive brugere

                                
                              if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\medarb.asp" then


                            
                                    Set myMail=CreateObject("CDO.Message")
                                    myMail.Subject="Timeout - Medarbejder profil opdateret"
                                
                                            select case lto
                                            case "kejd_pb"
                                            myMail.From = "joedan@okf.kk.dk"
                                            case else
					                        myMail.From = "timeout_no_reply@outzource.dk"
				                            end select

                                    'myMail.To=strEmail
                                    if len(trim(strEmail)) <> 0 then
                                    myMail.To= ""& strNavn &"<"& strEmail &">"
                                    end if

                                    if lto = "mi" then
                                    myMail.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\help_and_faq\TimeOut_indtasttimer_rev_20120106.pdf" 
                                    else
                                    myMail.AddAttachment "D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\help_and_faq\TimeOut_indtasttimer_rev_20130102.pdf" 
                                    end if
                    
                              
                                    mailtextbody= "" & medarb_txt_009 &" "& strNavn & vbCrLf _ 
					                & medarb_txt_010 & vbCrLf & vbCrLf _
					                & medarb_txt_011 & vbCrLf _ 
					                & medarb_txt_012 &" "& strLogin & " " & medarb_txt_013 & strPw & vbCrLf & vbCrLf _ 
					                & medarb_txt_014  & vbCrLf _ 
					                & medarb_txt_015 & vbCrLf & vbCrLf 

                                    if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
					                mailtextbody = mailtextbody & medarb_txt_119 &": https://outzource.dk/"&lto&""& vbCrLf & vbCrLf 
                                    else
                                    mailtextbody = mailtextbody & medarb_txt_119 &": http://timeout.cloud/"&lto&""& vbCrLf & vbCrLf 
                                    end if

					                mailtextbody = mailtextbody & medarb_txt_120& vbCrLf & vbCrLf & strEditor & vbCrLf 

                        
                                    myMail.TextBody= mailtextbody

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
                    


    					

                     
    					
					                   ' if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\medarb_red.asp" then
    					
    					        
                                                'if len(trim(strEmail)) <> 0 then

					                            'If Mailer.SendMail Then
        				
						                '            call medarb_opr_ok
    						
					                            'Else
    				                            'Response.Write "Fejl...<br>" & Mailer.Response
  					                            'End if

                                                'else

                                                    '***call medarb_opr_ok***

                                                'end if
    					
					                    'else
    						
						                    '***call medarb_opr_ok***
    						
					                    'end if
    						
			                    end if
			                else
			                    '*** Hvis dublet fandtes ***
			

			                    if strMnrOK <> "y" then
			                    errortype = 10
			                    call showError(errortype)
                                response.end
			                    else
			
			                    if strLoginOK <> "y" then
			                    errortype = 11
			                    call showError(errortype)
                                response.end

						                    end if
					                    end if
				                    end if
			                    end if
		                end if

                %>
                <div class="container" style="width:500px;">
            <div class="porlet">
            
            <h3 class="portlet-title">
               <u><%=medarb_txt_016 %></u>
            </h3>
            
                <div class="portlet-body">
                    <div><%=medarb_txt_017 %> 
                    </div><br />
                    <div>
                       
                        <%if request("medarbtypligmedarb") = "1" then 'Medarbejdertype:medarbejder 1:1%> 
                        <b><a href="medarbtyper.asp?lastmedid=<%=id %>&mtypeIdforvlgt=<%=strMedarbejdertype%>&mtypenavnforvlgt=Mtyp: <%=strNavn %>&func=opret"><%=medarb_txt_018 %> >></a>    
                        <%else%>
                        <b><a href="medarb.asp?menu=medarbejder&lastmedid=<%=id %>"><%=medarb_txt_019 %> >></a>
                        <%end if %>
                       
                    </div>
                    <br /><br />
                 </div>

            </div>
        </div>
                <%
               

        'Response.write "Medarbejder oprettet. <a href='medarb.asp?menu=medarbejder&lastmedid="&id&"'>Videre..</a>"
        'Response.Redirect "medarb.asp?menu=medarbejder&lastmedid="&id

         Response.end
                
                
        
    %>


<%
    case "red", "opret"

    %>
    <script src="js/medarb_jav.js" type="text/javascript"></script>        
    <%
     
	if func = "red" then

        HeaderTxt = "Redigér"


		strSQL = "SELECT mid, mnavn, mnr, mansat, login, pw, lastlogin, brugergruppe, "_
		&" medarbejdertype, type, navn, medarbejdere.editor, medarbejdere.dato, "_
		&" Medarbejderinfo, email, tsacrm, smilord, exchkonto, init, timereg, ansatdato, "_
		&" opsagtdato, sprog, nyhedsbrev,  "_
		&" madr, mpostnr, mcity, mland, mtlf, mcpr, mkoregnr, visskiftversion, timer_ststop, create_newemployee "_
		&" FROM medarbejdere, brugergrupper, medarbejdertyper "_
		&" WHERE mid = "& id &" AND brugergrupper.id = brugergruppe AND medarbejdertyper.id = medarbejdertype"
		
		oRec.open strSQL, oConn, 3
		
		if not oRec.EOF then
			
			strNavn = oRec("Mnavn")
			strAnsat = oRec("Mansat")
			strLogin = oRec("login")
			strPw = oRec("pw")
			strMnr = oRec("Mnr")
			strEditor = oRec("editor")
			strDato = oRec("dato")
			strMedinfo = oRec("Medarbejderinfo")
			dbfunc = "dbred"
			varsubval = "Opdater"
			strEmail = oRec("email")
			intCRM = oRec("tsacrm")
			intSmil = oRec("smilord")
            
            if len(trim(oRec("exchkonto"))) <> 0 then
			strExch = replace(oRec("exchkonto"), "#", "\")
            else
            strExch = ""
            end if
			
            strId =oRec("mid")
            strInit = oRec("init")
			intTimereg = oRec("timereg")
			ansatdato = formatdatetime(oRec("ansatdato"), 2) 'day(oRec("ansatdato")) &"-"& month(oRec("ansatdato")) &"-"& year(oRec("ansatdato"))
			opsagtdato = formatdatetime(oRec("opsagtdato"),2) 'day(oRec("opsagtdato")) &"-"& month(oRec("opsagtdato")) &"-"& year(oRec("opsagtdato"))
			sprog = oRec("sprog")
			nyhedsbrev = oRec("nyhedsbrev")
			
			madr = oRec("madr")
			mpostnr = oRec("mpostnr")
			mcity = oRec("mcity")
			mland = oRec("mland")
			mtlf = oRec("mtlf")
			mcpr = oRec("mcpr")
			mkoregnr = oRec("mkoregnr")

            visskiftversion = oRec("visskiftversion")
            
            medarbejdertype = oRec("medarbejdertype")

            timer_ststop = oRec("timer_ststop")

            create_newemployeeThisMid = oRec("create_newemployee")
			
		end if
		oRec.close



        else

        HeaderTxt = "Opret"

        strSQL = "SELECT Mnr FROM medarbejdere ORDER BY Mnr"
		oRec.open strSQL, oConn, 3
		if not oRec.EOF then
		oRec.movelast
		
		strNavn = ""
		strAnsat = ""
		strLogin = ""
		strPw = ""
		strMnr = (oRec("Mnr") + 1)
		strEditor = ""
		strDato = ""
		strMedinfo = ""
		dbfunc = "dbopr"
		varSubVal = "Opret" 
		strEmail = ""
		strCRMcheckedTSA = "CHECKED"
		intSmil = 1
		strExch = ""
		strInit = ""
		intTimereg = "1"
		ansatDato = date()
		opsagtdato = "01-01-2044"
		nyhedsbrev = 0
        visskiftversion = 0

        medarbejdertype = 0
        timer_ststop = 0
        create_newemployeeThisMid = 0

       
		end if
		oRec.close


        select case lto 
        case "esn", "tec", "intranet - local"
        intCRM = 3
        case else
        intCRM = 0
        end select

	end if
	



          strCRMcheckedCRM = ""
		strCRMcheckedTSA = ""
		strCRMcheckedRes = ""
         strCRMcheckedTSA_3 = ""
        strCRMcheckedTSA_4 = ""
        strCRMcheckedTSA_5 = ""
        strCRMcheckedTSA_6 = ""
        strCRMcheckedTSA_7 = ""
        strCRMcheckedTSA_8 = ""
		
		select case intCRM 
		case 1
		strCRMcheckedCRM = "checked"
		case 2
		strCRMcheckedRes = "checked"
		case 3
        strCRMcheckedTSA_3 = "CHECKED"
        case 4
        strCRMcheckedTSA_4 = "CHECKED"
        case 5
        strCRMcheckedTSA_5 = "CHECKED"
        case 6
        strCRMcheckedTSA_6 = "CHECKED"
        case 7
        strCRMcheckedTSA_7 = "CHECKED"
        case 8
        strCRMcheckedTSA_8 = "CHECKED"
		case else
		strCRMcheckedTSA = "CHECKED"
		end select


	
            if cint(visskiftversion) = 1 then
            visskiftversionCHK = "CHECKED"
            else
            visskiftversionCHK = ""
            end if

            if cint(timer_ststop) = 1 then
            timer_ststopCHK = "CHECKED"
            else
            timer_ststopCHK = ""
            end if


            if cint(create_newemployeeThisMid) = 1 then
            create_newemployeeCHK = "CHECKED"
            else
            create_newemployeeCHK = ""
            end if
            

		if intSmil = 1 then
			chkSmil = "CHECKED"
		else
			chkSmil = ""
		end if
		
		
		if intTimereg = "1" then
		strTregChk0 = ""
		strTregChk1 = "CHECKED"
		else
		strTregChk0 = "CHECKED"
		strTregChk1 = ""
		end if

        

%>

    

     <div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u><%=HeaderTxt & " Medarbejder"%></u>
        </h3>
        

    <form action="medarb.asp?func=<%=dbfunc%>" method="post">
    <input type="hidden" name="FM_smiley" id="Hidden1" value="1">
	<input type="hidden" name="FM_timereg" id="Hidden2" value="1">
	<input type="hidden" name="id" value="<%=id%>">
    
            <div class="row">
            <div class="col-lg-10">&nbsp</div>
            <div class="col-lg-2 pad-b10">
            <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
            </div>
            </div>

        <div class="portlet-body">
         <section>
                <div class="well well-white">
                    <div class="row">
                        <div class="col-lg-12">
                            <h4 class="panel-title-well"><%=medarb_txt_021 %></h4>
                        </div>
                    </div>

                    <div class="row">
                         <div class="col-lg-1">&nbsp;</div>

                        <div class="col-lg-2 pad-t5"><%=medarb_txt_022 %>:&nbsp<span style="color:red;">*</span>
                           

                        </div>
                        <div class="col-lg-3">   
                            <input name="FM_navn" type="text" class="form-control input-small" value="<%=strNavn %>" />
                        </div>
                        <!-- Sortering -->

                        <div class="col-lg-1 pad-t5"><%=medarb_txt_023 %>:&nbsp<span style="color:red;">*</span></div>
                        <div class="col-lg-2">   
                            <input name="FM_init" type="text" class="form-control input-small" value="<%=strInit%>" />

                        </div>
                                
                           
                      
                        <!-- Id -->
                        <%if level <= 2 OR level = 6 then%>
                        <div class="col-lg-1 pad-t5"><%=medarb_txt_024 %>:&nbsp<span style="color:red;">*</span></div>
                        <div class="col-lg-1">   
                            <input name="FM_Mnr" type="text" class="form-control input-small" value="<%=strMnr%>" /> 
                           
                        </div>
                        <%else %>
                        <input type="hidden" name="FM_Mnr" value="<%=strMnr%>">
                        <%end if %>
                    </div>
                    
                
            <!--Log ind-->
           
                     <div class="row">
                                <div class="col-lg-1">   
                                    &nbsp;
                                </div>
                                <div class="col-lg-2 pad-t5"><%=medarb_txt_025 %>:&nbsp<span style="color:red;">*</span></div>
                                <div class="col-lg-3">   
                                    <input name="FM_login" type="text" class="form-control input-small" value="<%=strLogin %>" />
                                </div>
                                <div class="col-lg-1 pad-t5"><%=medarb_txt_026 %>:&nbsp<span style="color:red;">*</span></div>
                                <div class="col-lg-4">   
                                    <%if func = "red" then 
		                                strPw = "KEEPTHISPW99"
		                                else
		                                strPw = ""
		                                end if %>

                                    <input name="FM_pw" id="pw" type="password" class="form-control input-small" value="<%=strPw%>" />
                                </div>     
                     </div>

                    <div class="row">
                        <div class="col-lg-1">&nbsp</div>
                        <div class="col-lg-2"><%=medarb_txt_027 %>:</div>
                        <div class="col-lg-3">
                            <%
		                    select case strAnsat
		                    case "2"
		                    chk1 = ""
		                    chk2 = "CHECKED"
		                    chk3 = ""
		                    case "3"
		                    chk3 = "CHECKED"
		                    chk2 = ""
		                    chk1 = ""
		                    case else
		                    chk1 = "CHECKED"
		                    chk2 = ""
		                    chk3 = ""
		                    end select 
		
	                    %>
		                        <input type="radio" name="FM_ansat" id="FM_ansat1" value="1" <%=chk1%>> <%=medarb_txt_029 %> <%if func <> "red" then %>
                                <span style="font-size:10px;">(<%=medarb_txt_028 %>)</span>
                                <%end if %>
		                    <br /><input type="radio" name="FM_ansat" id="FM_ansat2" value="2" <%=chk2%>> <%=medarb_txt_030 %> 
		                    <br /><input type="radio" name="FM_ansat" id="FM_ansat3" value="3" <%=chk3%>> <%=medarb_txt_031 %>
		                    <br />

                        </div>


                        <div class="col-lg-6">

                        <div id="deaktivnote" style="visibility:hidden; display:none; width:450px; padding:20px; background-color:#cccccc;">
                                <%=medarb_txt_032 %><br /><br />
                                <input type="checkbox" value="1" name="FM_deakdd_cancel" /> <%=medarb_txt_033 %> <b><%=medarb_txt_034 %></b>

                            </div>
                         </div>

                    </div>
            
                </div>
             </section>

            <section>
                <!-- Accordion -->
                <div class="panel-group accordion-panel" id="accordion-paneled">
                    <!-- PersonData -->
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseOne">
                            <%=medarb_txt_035 %>
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseOne" class="panel-collapse collapse">
                        <div class="panel-body">
                                <div class="row">
                                 <div class="col-lg-1">&nbsp;</div>
                        <div class="col-lg-2"><%=medarb_txt_036 %>:</div>
                        <div class="col-lg-3">   
                                    <input name="FM_adr" type="text" class="form-control input-small" value="<%=madr%>"  />
                            </div>
                                    <div class="col-lg-6">&nbsp;</div> 
                        </div>
                <div class="row">
                                <div class="col-lg-1">   
                                    &nbsp;
                                    </div>
                    <div class="col-lg-2"><%=medarb_txt_037 %>:</div>
                    <div class="col-lg-1">   
                                    <input name="FM_postnr" type="text" class="form-control input-small" value="<%=mpostnr%>" />           
                    </div>
                    <div class="col-lg-1"><%=medarb_txt_038 %>:</div>
                    <div class="col-lg-3">   
                                    <input name="FM_city" type="text" class="form-control input-small" value="<%=mcity%>" />
                    </div>
                    <div class="col-lg-4">&nbsp;</div>   
                </div>

                    <div class="row">
                                    <div class="col-lg-1">   
                                            &nbsp;
                                            </div>
                        <div class="col-lg-2"><%=medarb_txt_039 %>:</div>
                        <div class="col-sm-3">
                        <select class="form-control input-small" name="FM_land" style="width:200px;">
						            <!--#include file="../timereg/inc/inc_option_land.asp"-->
                                <%if func = "red" then%>
                        		<option SELECTED><%=mland%></option>
		                        <%else%>
		                        <option SELECTED><%=medarb_txt_040 %></option>
		                        <%end if%>

		                </select>

                         </div>
                    </div>
                    <div class="row">
                                    <div class="col-lg-1">   
                                            &nbsp;
                                            </div>
                        <div class="col-lg-2"><%=medarb_txt_041 %>:</div>
                        <div class="col-lg-2"> 
                        <input name="FM_tel" type="text" class="form-control input-small" value="<%=mtlf%>"/>
                                  </div>
                        <div class="col-lg-7">&nbsp;</div>    
                    </div>
                         <div class="row">
                                <div class="col-lg-1">   
                                    &nbsp;
                                    </div>
                    <div class="col-lg-2"><%=medarb_txt_042 %>:</div>
                    <div class="col-lg-3">   
                                    <input name="FM_email" type="text" class="form-control input-small" value="<%=strEmail%>" />
                                    <span style="color:#999999; font-size:10px; line-height:10px;"><%=medarb_txt_043 %></span>               
                    </div>
                                    <div class="col-lg-5">&nbsp;</div>     
                </div>
                    <div class="row">
                                <div class="col-lg-1">   
                                    &nbsp;
                                    </div>
                    <div class="col-lg-2"><%=medarb_txt_044 %>:</div>
                    <div class="col-lg-3">   
                                    <input name="FM_cpr" type="text" class="form-control input-small" value="<%=mcpr%>" />           
                    </div>
                                    <div class="col-lg-6">&nbsp;</div>
                </div>
                    <div class="row">
                                <div class="col-lg-1">   
                                    &nbsp;
                                    </div>
                    <div class="col-lg-2"><%=medarb_txt_045 %>:</div>
                    <div class="col-lg-3">   
                                    <input name="FM_regnr" type="text" class="form-control input-small" value="<%=mkoregnr%>" />  
                                          <div class="col-lg-8">   
                                          &nbsp;
                                         </div>         
                   
                    </div>
                </div>
                    
                          </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                      </div> <!-- /.panel -->

                    

                        <%
                            
                     'call meStamdata(session("mid"))       
                     if cint(level) <= 2 OR cint(level) = 6 OR cint(meCreate_newemployee) = 1 then %>
                    <!-- Medarbejder -->
                    <div class="panel panel-default">
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTwo">
                            <%=medarb_txt_046 %>
                            </a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseTwo" class="panel-collapse collapse">
                          <div class="panel-body">
                        <%end if %>

                            
                              <%if cint(level) = 1 OR cint(meCreate_newemployee) = 1 then %>
                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2"><%=medarb_txt_047 %>:</div>
                                  <div class="col-lg-2">
                                        <div class='input-group date'>
                                              <input class="form-control input-small" type="text" name="ansatdato" id="ansatdato" value="<%=ansatdato%>"/>
                                              <!--<span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>-->
                                        </div>

                                  </div>
                                  <div class="col-lg-1"><%=medarb_txt_048 %>:</div>
                                  <div class="col-lg-2">
                                      <div class='input-group date'>
                                      <input class="form-control input-small" type="text" name="opsagtdato" id="opsagtdato" value="<%=opsagtdato%>" />
                                           <!--<span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>-->
                                        </div>
                                  </div>
                                        <div class="col-lg-4">&nbsp</div>
                              </div>
                              
                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2"><%=medarb_txt_049 %>:</div>
                                  <div class="col-lg-2"><select class="form-control input-small" name="FM_sprog" style="width:160px;">
                                     <%

                                    select case lto
                                    case "outz", "hidalgo"
                                         sprogAntalIds = 8
                                    case else
                                         sprogAntalIds = 3
                                    end select
		
			                        strSQL = "SELECT sproglabel, id FROM sprog WHERE id < "& sprogAntalIds
			                        oRec.open strSQL, oConn, 3
			                        while not oRec.EOF 
			                         
                                     if cint(sprog) = oRec("id") then
			                         sprogSEL = "SELECTED"
			                         else
			                         sprogSEL = ""
			                         end if
			                         %>
			                        <option value="<%=oRec("id")%>" <%=sprogSEL %>><%=oRec("sproglabel")%></option>
			                        <%
			                        oRec.movenext 
			                        wend
			                        oRec.close
	                            %></select>
                                  </div>

                              </div>
                              <% 
	                            if cint(nyhedsbrev) = 0 then
	                            chkny0 = "CHECKED"
	                            chkny1 = ""
	                            else
	                            chkny0 = ""
	                            chkny1 = "CHECKED"
	                            end if 
	                            
                                  
                              if cint(level) = 1 then%>	
                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2 pad-t20 pad-b20"><%=medarb_txt_050 %>:</div>
                                  <div class="col-lg-2 pad-t20 pad-b20"><%=medarb_txt_051 %> <input id="nyhedsbrev" name="nyhedsbrev" type="radio" value="1" <%=chkny1 %> />&nbsp;&nbsp;<%=medarb_txt_008 %> 
                                    <input id="nyhedsbrev" name="nyhedsbrev" type="radio" value="0" <%=chkny0 %> /> </div>
                              </div>

                              

                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2"><%=medarb_txt_052 %>:</div>
                                  <div class="col-lg-6">
                                      <input class="form-control input-small" type="text" name="FM_exch" value="<%=strExch%>" placeholder="<%=medarb_txt_053 %>"/>
                                  </div>
                                  
                              </div>
                              <br /><br />
                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2 pad-b20"><%=medarb_txt_054 %>:</div>
                                  <div class="col-lg-6">
                                     
                                      <input type="checkbox" name="FM_visskiftversion" value="1" <%=visskiftversionCHK %>/> <%=medarb_txt_055 %><br />
                                      <input type="checkbox" name="FM_opretmedarb" value="1" <%=create_newemployeeCHK %>/> <%=medarb_txt_056 %>
                                  </div> 
                               <div class="col-lg-6">&nbsp</div>  
                              </div>
                              <%else %>
                                
                                    <input type="hidden" name="FM_exch" value="<%=strExch%>"/>
                                    <input type="hidden" name="nyhedsbrev" value="<%=nyhedsbrev%>"/>
                                    <input type="hidden" name="FM_visskiftversion" value="<%=visskiftversion%>"/>
                                    <input type="hidden" name="FM_opretmedarb" value="<%=create_newemployee%>"/>

                              <%end if 'level %>


                              <%else %>

                                    <input type="hidden" name="ansatdato" value="<%=ansatdato%>"/>
                                    <input type="hidden" name="opsagtdato" value="<%=opsagtdato%>" />
                                    <input type="hidden" name="FM_sprog" value="<%=sprog%>"/>
                                    <input type="hidden" name="FM_exch" value="<%=strExch%>"/>
                                    <input type="hidden" name="nyhedsbrev" value="<%=nyhedsbrev%>"/>
                                    <input type="hidden" name="FM_visskiftversion" value="<%=visskiftversion%>"/>
                                    <input type="hidden" name="FM_opretmedarb" value="<%=create_newemployee%>"/>

                              <%end if 'level%>



                              <%if cint(level) = 1 OR cint(meCreate_newemployee) = 1 then %>
                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2">
                                     
                                      <a href="medarbtyper.asp" target="_blank"><%=medarb_txt_057 %>:</a>
                                      
                                      &nbsp&nbsp<a data-toggle="modal" href="#styledModalSstGrp20"><span class="fa fa-info-circle"></span></a>
                                      
                                  </div>
                                    
                                    <div id="styledModalSstGrp20" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                                        <div class="modal-dialog">
                                            <div class="modal-content" style="border:none !important;padding:0;">
                                              <div class="modal-header">
                                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                  <h5 class="modal-title"><%=medarb_txt_058 %></h5>
                                                </div>
                                                <div class="modal-body">
                                                <%=medarb_txt_059 %>

                                                </div>
                                            </div>
                                        </div>
                                 </div>

                                  <div class="col-lg-8">

                                  
                                  <% if func = "red" then

	                                        antalhist = 0
	                                        strSQLantalth = "SELECT COUNT(id) AS antalhist FROM medarbejdertyper_historik WHERE mid = "& id &" GROUP BY mid"
	                                        oRec.open strSQLantalth, oConn, 3
	                                        if not oRec.EOF then
	                                        
	                                            antalhist = oRec("antalhist") 
	            
	                                        end if
	                                        oRec.close
	        
	                                    if cint(antalhist) >= 30 then
	                                    dis = "DISABLED"
	                                    else
	                                    dis = ""
	                                    end if
	        
	                                  else
	                                    dis = ""
	                                 end if %>


                                     
                                  <select class="form-control input-small" <%=dis%> name="FM_medtype" id="FM_medtype" style="width:540px;">
                                     
                                    <%	strSQL = "SELECT medarbejdertype, mt.type, mt.id, mgruppe, "_
                                    &" mtg.navn AS mtgnavn "_ 
                                    &" FROM medarbejdere AS m, medarbejdertyper AS mt"_
                                    &" LEFT JOIN medarbtyper_grp AS mtg ON (mtg.id = mgruppe) WHERE mid = "& id &" AND mt.id = medarbejdertype" 'AND mid = 99999999
			                        %>
            
            
		                        <%
		                                if func = "red" then
		
		
            
                                            oRec.open strSQL, oConn, 3
			                                if not oRec.EOF then
			        
                                            if oRec("mgruppe") <> 0 then
                                            mtgNavn = " [hovedgruppe:  "& oRec("mtgnavn") &"]"
                                            else
                                            mtgNavn = ""
                                            end if
                                            %>
			                                <option value="<%=oRec("id")%>" SELECTED><%=oRec("type")%> <span style="font-size:8px;"><%=mtgNavn%></span></option>
			                                <%
			                                end if
			                                oRec.close

		                                end if
		
		                        'strSQL = "SELECT id, type FROM medarbejdertyper ORDER BY type"
		                        strSQL = "SELECT mt.type, mt.id, mgruppe, "_
                                &" mtg.navn AS mtgnavn "_ 
                                &" FROM medarbejdertyper AS mt"_
                                &" LEFT JOIN medarbtyper_grp AS mtg ON (mtg.id = mgruppe) WHERE mt.id <> 0 ORDER BY mtg.navn, type"
                                
                                        
                                oRec.open strSQL, oConn, 3
		                        While not oRec.EOF
                
                                    if oRec("mgruppe") <> 0 then
                                    mtgNavn = " [ hovedgruppe:  "& oRec("mtgnavn") &"]"
                                    else
                                    mtgNavn = ""
                                    end if

                                    if cint(medarbejdertype) = oRec("id") then
                                    mtypSEL = "SELECTED"
                                    else
                                    mtypSEL = ""
                                    end if

                                    if lastMtypHgrp <> oRec("mtgNavn") then
                                    %>  
                                        <option value="0" disabled></option>
                                        <option value="0" disabled><%=medarb_txt_060 %>: <%=oRec("mtgNavn")%></option>
                                    <%
            
                                    end if

                                        %>
		                        <option value="<%=oRec("id")%>" <%=mtypSEL%>><%=oRec("type")%></option>
		                        <%

                                lastMtypHgrp = oRec("mtgNavn")
		                        oRec.movenext
		                        Wend
		                        oRec.close
		                        %>
		                        </select> 
                                      
                                      
                                      <%if func = "red" AND level = 1 then %>
                                      <a href="medarbtyper.asp?func=red&id=<%=medarbejdertype %>" target="_blank"><%=medarb_txt_061 %></a>
                                      <%else
                                          
                                      call medarbtypligmedarb_fn()%>
                                      <input type="hidden" name="medarbtypligmedarb" value="<%=medarbtypligmedarb %>" />   
                                      <%end if %>

                                     
                                        <%if func = "opret" AND level = 1  then
                                        %>
                                        <br /><span style="color:red;"><%=medarb_txt_062 %></span> <%=medarb_txt_063 %><br />
                                        <%=medarb_txt_064 %>
                                        <select name="FM_medarbejdertype_follow_tp" class="form-control input-small" style="width:200px;">
                                        <%
                                         strSQL = "SELECT mt.type, mt.id FROM medarbejdertyper mt WHERE mt.id <> 0 ORDER BY type"
                                
                                        
                                                oRec.open strSQL, oConn, 3
		                                        While not oRec.EOF
                
                                                 %>
		                                        <option value="<%=oRec("id")%>"><%=oRec("type")%></option>
		                                        <%

                              
		                                        oRec.movenext
		                                        Wend
		                                        oRec.close
		                                        %>
		                                        </select> 
                                
                                      <%end if %>

                              <%
                              else

		                        strSQL = "SELECT medarbejdertype, type, id FROM medarbejdere, medarbejdertyper WHERE Mid = "& id &" AND medarbejdertyper.id = medarbejdertype"

		                        oRec.open strSQL, oConn, 3
		                        if not oRec.EOF then%>
		                        <input type="hidden" name="FM_medtype" value="<%=oRec("id")%>">
		                        <%
		                        end if
		                        oRec.close
	                        
                            end if
                            %>





                                 <%if cint(level) = 1 then %>
                                <br /><br /><b><%=medarb_txt_065 %>:</b> (<%=medarb_txt_066 %>)<br />
                               <%=medarb_txt_067 %><br />
                                <%=medarb_txt_068 %><br /><br />
		                        <%end if
		
			                    strSQL = "SELECT mtypedato, mt.type AS mttype, mth.id AS mtypeid FROM medarbejdertyper_historik AS mth, "_
			                    &" medarbejdertyper mt WHERE mth.mid = "& id &" AND mt.id = mth.mtype ORDER BY mth.id"

                                'Response.Write strSQL
                                'Response.flush
                                mth = 1
			                    oRec.open strSQL, oConn, 3
			                    while not oRec.EOF 
			
                                if cint(level) = 1 then %>
			                    
			                    <%=oRec("mttype")%> - <%=medarb_txt_129 %>
                                <input id="Hidden4" type="hidden" name="FM_mtyphist_id" value="<%=oRec("mtypeid") %>" />
            
                                    <%if cint(mth) = 1 then %>
                                    <input id="Hidden6" type="hidden" name="FM_mtyphist_dato" value="<%=oRec("mtypedato") %>" /><%=oRec("mtypedato") %><br />
                                    <%else %>
                                    <input id="Text9" type="text" name="FM_mtyphist_dato" style="font-size:9px; width:80px;" value="<%=oRec("mtypedato") %>" /><br />
                                    <%end if %>


                                <%else %>
                                <input id="Hidden5" type="hidden" name="FM_mtyphist_id" value="0" />
                                <input id="Text10" type="hidden" name="FM_mtyphist_dato" value="<%=oRec("mtypedato") %>" /><!--<%=oRec("mtypedato") %><br />-->
                                <%end if %>
			                    <% 

                                mth = mth + 1
			                    oRec.movenext
			                    wend 
			                    oRec.close
			
                                    if cint(level) = 1 then

			                        if mth = 1 then
			                        %>
			                        (Ingen)
			                        <%
			                        end if

                                    end if 'level

                                
		
		

                            if cint(level) = 1 AND func = "red" then
                            %>
                            <br /><br /><br />
                            <%=medarb_txt_069 %><br />
                                <span style="color:#999999;"><%=medarb_txt_070 %></span><br />
                                <input type="checkbox" name="FM_opd_mty_kp" id="FM_opd_mty_kp" value="1"/> <%=medarb_txt_071 %> <b><%=medarb_txt_072 %></b> 
                                 <br />
                                 <input type="checkbox" name="FM_opd_mty_tp" id="FM_opd_mty_tp" value="1"/> <%=medarb_txt_071 %> <b><%=medarb_txt_073 %></b> (<%=medarb_txt_074 %>)
                               <br />&nbsp;
        
                                <%
                            end if
                              
                              if cint(level) = 1 OR cint(meCreate_newemployee) = 1 then %>
                             </div>
                             </div><!-- ROW -->
                              <%end if %>


                              <%if cint(level) = 1 OR cint(meCreate_newemployee) = 1 then %>
                               <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2"><br /><br /><%=medarb_txt_075 %>:&nbsp&nbsp<a data-toggle="modal" href="#styledModalSstGrp3"><span class="fa fa-info-circle"></span></a></div>
                                    <div id="styledModalSstGrp3" class="modal modal-styled fade" style="top:60px;">
                                                    <div class="modal-dialog">
                                                        <div class="modal-content" style="border:none !important;padding:0;">
                                                            <div class="modal-header">
                                                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                                <h5 class="modal-title"><%=medarb_txt_075 %></h5>
                                                            </div>
                                                            <div class="modal-body">
                                                                <%=medarb_txt_076 %> 
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>

                                  <div class="col-lg-4"><br /><br />
                                  <select class="form-control input-small" name="FM_bgruppe">

		                      <%end if 'level


                                        brugergpValgt = 7 'TSA NIV 2
		                                '** Finder den allerede valgte brugergruppe **
		                                if func = "red" then
			                                strSQL = "SELECT brugergruppe, navn, id FROM medarbejdere, brugergrupper WHERE Mid = "& id &" AND brugergrupper.id = brugergruppe"
			                                oRec.open strSQL, oConn, 3
			                                if not oRec.EOF then
			                                
                                            if cint(level) = 1 then%>
			                                <option value="<%=oRec("id")%>" SELECTED><%=oRec("navn")%></option>
			                                <%end if

                                            brugergpValgt = oRec("id")
			                                end if
			                                oRec.close
		                                end if



                                        if cint(level) = 1 OR cint(meCreate_newemployee) = 1 then
                                       
		
		                                '** Hvis der oprettes ny medarbejder sættes brugergruppe 2 () til default ***
		                                if level <> 1 then
		                                strWHEREKlaus = " WHERE rettigheder = 7"
		                                else
		                                strWHEREKlaus = " WHERE rettigheder <> 0" '= 2 OR rettigheder = 6 "
		                                end if
		
		                                strSQL = "SELECT id, navn FROM brugergrupper "& strWHEREKlaus &" Order by id"
		                                oRec.open strSQL, oConn, 3
		                                While not oRec.EOF

                                        if oRec("id") = 7 AND func <> "red" then%>
		                                <option value="<%=oRec("id")%>" SELECTED><%=oRec("navn")%></option>
		                                <%else%>
		                                <option value="<%=oRec("id")%>"><%=oRec("navn")%></option>
		                                <%end if


		                                oRec.movenext
		                                Wend
		                                oRec.close
		                                %>
		                                </select>

                                       <%else%>
                                             <input type="hidden" name="FM_bgruppe" value="<%=brugergpValgt%>"/>
                                        <%end if%>

                                    <%if cint(level) = 1 OR cint(meCreate_newemployee) = 1 then %>
                                  </div>
                                   <div class="col-lg-5">&nbsp</div>
                              </div>
                              <%end if %>


                              <%if cint(level) = 1 OR cint(meCreate_newemployee) = 1 then %>
                              <div class="row">
                                  <div class="col-lg-1">&nbsp</div>
                           
                                   
                                  <div class="col-lg-2 pad-t10">
                                    
                                      <a href="projektgrupper.asp?" target="_blank"><%=medarb_txt_077 %>:</a>
                                      &nbsp&nbsp<a data-toggle="modal" href="#styledModalSstGrp40"><span class="fa fa-info-circle"></span></a>

                                        <div id="styledModalSstGrp40" class="modal modal-styled fade" style="top:60px;">
                                                    <div class="modal-dialog">
                                                        <div class="modal-content" style="border:none !important;padding:0;">
                                                            <div class="modal-header">
                                                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                                                <h5 class="modal-title"><%=medarb_txt_077 %></h5>
                                                            </div>
                                                            <div class="modal-body">
                                                                 <%=medarb_txt_078 %> 
                                                                <%=medarb_txt_079 %><br /><br />
                                                                <%=medarb_txt_080 %> <u><%=medarb_txt_081 %></u> <%=medarb_txt_082 %>  
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>

                                      <span style="font-size:10px;"></span>
                                  </div>
                                  
                                   <div class="col-lg-4 pad-t10">
                                   <input id="Hidden3" name="FM_progrp" type="hidden" value="10" /><!-- Alle gruppen -->
                                   <select id="FM_progrp" name="FM_progrp" multiple style="height:200px;" class="form-control input-small" <%=projGrpDis %>>
            
                                    <%
    
                                    selprogrp = ""
    
                                    if func = "opret" then
                                    thisMid = 0
                                    else
                                    thisMid = id
                                    end if
    
                                    strNOTpgids = ""
    
                                                        '*** Henter projektgrupper grupper medarbejder er med i ***'
                                                        strSQLpg = "SELECT Mid, MedarbejderId, p.navn AS pgnavn, p.id AS pid, pr.teamleder, pr.notificer FROM progrupperelationer pr "_
                                                        &" LEFT JOIN medarbejdere AS m ON (Mid = MedarbejderId) "_
                                                        &" LEFT JOIN projektgrupper AS p ON (p.id = pr.projektgruppeId) WHERE Mid = "& thisMid &" ORDER BY p.navn"
	                                                    oRec2.open strSQLpg, oConn, 3 
	                                                    pgr = 1
	
	                                                    while not oRec2.EOF 
	
	                                                    if len(trim(oRec2("pid"))) <> 0 then
	                                                    strNOTpgids = strNOTpgids & " AND pg.id <> " & oRec2("pid")
	                                                    end if
	
	                                                    if oRec2("pid") <> 10 then
        
                            
                                                            if oRec2("teamleder") <> 0 then
                                                            teaml = "_1"
                                                            teamlTxt = " ("& medarb_txt_083 &")"
                                                            else
                                                            teaml = ""
                                                            teamlTxt = ""
                                                            end if

                                                            if oRec2("notificer") <> 0 then
                                                            notf = "-1"
                                                            notfTxt = " ("& medarb_txt_084 &")"
                                                            else
                                                            notf = ""
                                                            notfTxt = ""
                                                            end if


	                                                    %>
                                                        <option value=<%=oRec2("pid") & teaml & notf %> SELECTED><%=oRec2("pgnavn") %> <%=teamlTxt %> <%=notfTxt %></option>
                                                        <%
    
                                                        selprogrp = selprogrp & ","& oRec2("pid")
    
	                                                    end if
	
	                                                    pgr = pgr + 1
	                                                    oRec2.movenext
	                                                    wend
	                                                    oRec2.close
	
	


	
	                                    '*** Henter resterende grupper ***'
	                                    strSQLrpg = "SELECT pg.navn, pg.id AS pid, opengp FROM projektgrupper pg WHERE pg.id <> 10 " & strNOTpgids & " ORDER BY pg.navn"
	                                    'Response.Write strSQLrpg
	                                    'Response.flush
	                                    oRec2.open strSQLrpg, oConn, 3 
	                                    while not oRec2.EOF 

                                        if cint(level) = 1 OR cint(meCreate_newemployee) = 1 OR oRec2("opengp") = 1 then
                                        %>
                                        <option value=<%=oRec2("pid")%>><%=oRec2("navn") %></option>

                                        <%else
                                
                                            if lto <> "wwf" then%>
                                            <option value=<%=oRec2("pid")%> DISABLED><%=oRec2("navn") %></option>

                                            <%end if
                                         end if

	                                    oRec2.movenext
	                                    wend
	                                    oRec2.close
	                                    %>
                                        </select>


                                   </div>
                                <div class="col-lg-7">&nbsp</div>      
                              </div>
                              <%else 
                                  
                                 '** Denne bruges IKKE mere da KUN administratorer må ændre projektgrupper%>

                                 <%'call medariprogrpFn(id) %>
                                <input name="FM_progrp" type="hidden" value="0" /> <!-- <%=replace(medariprogrpTxt, "#", "") %> --> 
                              <%end if %>


                               <%if cint(level) <= 2 OR cint(level) = 6 OR cint(meCreate_newemployee) = 1 then %>
                              <div class="row">
                                   <div class="col-lg-12"><br />&nbsp</div>
                                </div>
                              <div class="row">
                                  
                                  <div class="col-lg-1">&nbsp</div>
                                  <div class="col-lg-2"><b><%=medarb_txt_123 %>:</b></div>
                                  <div class="col-lg-5">
                                     
                                             <%if cint(level) <= 2 OR cint(level) = 6 then%> 
                                             <input type="radio" name="FM_tsacrm" value="0" <%=strCRMcheckedTSA%>> <%=medarb_txt_124 %> 
                                      <br />&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp  <input type="checkbox" name="FM_timer_ststop" value="1" <%=timer_ststopCHK%> /> <%=medarb_txt_085 %> <br>
          
                                          <br /> <input type="radio" name="FM_tsacrm" value="6" <%=strCRMcheckedTSA_6%>> <%=medarb_txt_086 %><br>
                                              <%end if %>


                                        <%if level = 1 then %>

                                        <%call erStempelurOn()
                
                                        if cint(stempelurOn) = 1 then %>
                                            <input type="radio" name="FM_tsacrm" value="3" <%=strCRMcheckedTSA_3%>> <%=medarb_txt_087 %>  <br>
                                        <%end if %>
		
		                             <input type="radio" name="FM_tsacrm" value="2" <%=strCRMcheckedRes%>> <%=medarb_txt_088 %><br>
                                    <input type="radio" name="FM_tsacrm" value="4" <%=strCRMcheckedTSA_4%>> <%=medarb_txt_089 %><br>
                                     <input type="radio" name="FM_tsacrm" value="5" <%=strCRMcheckedTSA_5%>> <%=medarb_txt_090 %><br>
                                      <input type="radio" name="FM_tsacrm" value="7" <%=strCRMcheckedTSA_7%>> <%=medarb_txt_091 %><br>
                                      <input type="radio" name="FM_tsacrm" value="8" <%=strCRMcheckedTSA_8%>> <%=medarb_txt_092 %><br>

                                     <%if licensType = "CRM" then%>
                                     <input type="radio" name="FM_tsacrm" value="1" <%=strCRMcheckedCRM%>> <%=medarb_txt_093 %><br> 
                                    <%end if%>

                                      <%end if 'level = 1 %>

                                  </div>
                            <%else %>
                                    <input type="hidden" name="FM_tsacrm" value="<%=intCRM %>">
                              <%end if %>
                              </div>



                                
              
                         
                               <%if cint(level) = 1 then %>

                          </div> <!-- /.panel-body -->
                        </div> <!-- /.panel-collapse -->
                      </div> <!-- /.panel -->
                    
                    <%end if %>
                   

                      <%if func = "red" then %>
                              <br /><br /><br />
                <div style="font-weight: lighter;"><%=medarb_txt_094 %> <b><%=strDato%></b> <%=medarb_txt_095 %> <b><%=strEditor%></b></div>

                            <%if func = "red" then %>
                            <span style="font-size:9px; color:#999999;"> 
                            TimeOut Id: <%=id %>
                            <%end if %></span><br /><br />&nbsp;

                <%else %>
                <br /><br />&nbsp;
                <%end if %>
   
             

                     </div> <!-- /.accordion -->
                
               </section>

                  <div style="margin-top:15px; margin-bottom:15px;">
                    <button type="submit" class="btn btn-success btn-sm pull-right"><b><%=medarb_txt_020 %></b></button>
                     
         
                    <div class="clearfix"></div>
                 </div>


        </div> <!-- /.portlet-body -->
        </form>
      </div> <!-- /.portlet -->
    </div> <!-- /.container -->

            <%

    case else 'list

                %>
                <script src="js/medarb_list_201612_jav.js" type="text/javascript"></script>
                <%


                if len(trim(request("lastmedid"))) <> 0 then
                lastmedid = request("lastmedid")
                else
                lastmedid = 0
                end if

                if len(trim(request("visikkemedarbejdere"))) <> 0 then
                visikkemedarbejdere = 1
                else
                visikkemedarbejdere = 0
                end if


                if len(trim(request("sogsubmitted"))) <> 0 then
                sogsubmitted = 1
                else
                sogsubmitted = 0
                end if

   
	            if cint(sogsubmitted) = 1 then
                ''**AND len(request("ktype")) <> 0 
	            '** request("ktype") tjekker at form e submitted **'
	            thiskri = request("FM_soeg")
	            useKri = 1

                response.cookies("medarb_2015")("medarbsogekri") = thiskri
	           
	            else
	        
	                    if request.cookies("medarb_2015")("medarbsogekri") <> "" then
	                    thiskri = request.cookies("medarb_2015")("medarbsogekri")
	                    useKri = 1
	                    visikkemedarbejdere = 0
	                    else
	                    thiskri = ""
	                    useKri = 0
	                    end if
	        
	            end if

                if cint(sogsubmitted) = 1 then
                    if len(trim(request("FM_vispasogluk"))) <> 0 then
                    vispasogluk = 1
                    else
                    vispasogluk = 0
                    end if
                else
                    if request.cookies("medarb_2015")("vispasogluk") <> "" then
                    vispasogluk = request.cookies("medarb_2015")("vispasogluk")
                    else
                    vispasogluk = 0
                    end if
                end if

                response.cookies("medarb_2015")("vispasogluk") = vispasogluk
                response.cookies("medarb_2015").expires = date + 1

                if cint(vispasogluk) = 1 then
                vispasoglukCHK = "CHECKED"
                else
                vispasoglukCHK = ""
                end if

   
                
         if media <> "eksport" then %>



 <!--Medarbejder liste-->
           
<script src="js/Medarbejder.js" type="text/javascript"></script>

    <div class="container">
      <div class="portlet">
        <h3 class="portlet-title">
          <u><%=medarb_txt_096 %></u>
        </h3>
        
          <%if level = 1 then 
              
              
              strSQL = "SELECT licens.key, klienter FROM licens WHERE id = 1"
		    oRec.Open strSQL, oConn, 0, 1, 1
		    if not oRec.EOF then
		    'ikl = left(oRec("key"), 4)
		    'akl = right(ikl, 2)
		
		    akl = oRec("klienter")
		
		        call antalmedarb_fn(2)

		    end if
		    oRec.close
              
              if cint(kli) < cint(akl) then '* kun admin brugere kan oprette medarb.
              
              %>
          <form action="medarb.asp?menu=medarber&func=opret" method="post">
                <section>
                         <div class="row">
                             <div class="col-lg-10">&nbsp;</div>
                             <div class="col-lg-2">
                            <button class="btn btn-sm btn-success pull-right"><b><%=medarb_txt_097 %> +</b></button><br />&nbsp;
                            </div>
                        </div>
                </section>
         </form>
          <%end if %>

          <%end if %>

         
          <section>

                 <form action="medarb.asp?vismedarbejdere=1&sogsubmitted=1" method="POST">
                <div class="well">
                         <h4 class="panel-title-well"><%=medarb_txt_098 %></h4>
                         <div class="row">
                             <div class="col-lg-3">
                                 <%=medarb_txt_099 %>: <br />
			                    <input type="text" name="FM_soeg" id="FM_soeg" class="form-control input-small" placeholder="% Wildcard" value="<%=thiskri%>">
                               </div>
                            <div class="col-lg-9"><br />
                            <button type="submit" class="btn btn-secondary btn-sm pull-right"><b><%=medarb_txt_100 %> >></b></button>
                            </div>
                             </div>
                    <div class="row">
                        <div class="col-lg-3"><input id="jq_vispasogluk" name="FM_vispasogluk" type="checkbox" <%=vispasoglukCHK %> value="1" /> <%=medarb_txt_101 %></div>
                    </div>
                            
                    </div>
              </form>
 

        <div class="portlet-body">

           <table id="medarb_list" class="table dataTable table-striped table-bordered table-hover ui-datatable">
       
           
            <thead>
              <tr>
                <th style="width: 22%"><%=medarb_txt_102 %></th>
                <th style="width: 5%"><%=medarb_txt_103 %></th>
                <th style="width: 15%"><%=medarb_txt_104 %></th>
                <th style="width: 18%"><%=medarb_txt_105 %></th>
                <th style="width: 15%"><%=medarb_txt_106 %></th>
                <th style="width: 10%"><%=medarb_txt_114 %></th>
                <th style="width: 10%"><%=medarb_txt_108 %></th>
                <th style="width: 5%"><%=medarb_txt_109 %></th>
                <th style="width: 5%"><%=medarb_txt_110 %></th>

              </tr>
            </thead>
              <tbody>

                <%

                end if 'media


                    'sort = Request("sort")
	                ekspTxt = ""
	
		
		            '** Søge kriterier **
		            if cint(visikkemedarbejdere) <> 1 then
		
			            if usekri = 1 then '**SøgeKri
			            sqlsearchKri = " (mnavn LIKE '%"& thiskri &"%' OR init LIKE '"& thiskri &"%' OR (mnr LIKE '"& thiskri &"%' AND mnr <> '0'))" 
			            else
			            sqlsearchKri = " (mid <> 0)"
			            end if
			
		            else
		            sqlsearchKri = " (mid = 0)"
		            end if

                    if cint(vispasogluk) = 1 then
                    vispasoglukKri = " AND mansat <> -1"
                    else
                    vispasoglukKri = " AND mansat = 1"
                    end if


                     '** Admin må søge
                    strSQLAdminRights = ""
                     if cint(meCreate_newemployee) = 1 AND level <> 1 then 'Hvis opret medab = OK må man godt søge i dem man står som editor på

                     '   if cint(meCreate_newemployee) = 1 then
                         'strSQLAdminRights = " AND m.editor = '"& meNavn &"'"
                     '   end if

                     else

                          if level <> 1 then 
                             strSQLAdminRights = " AND mid = " & session("mid") 
                          end if

                     end if
                     %>
                



                   <%strSQL = "SELECT mnavn, mid, email, lastlogin, init, mnr, mansat, brugergruppe, medarbejdertype, mt.type AS mtypenavn, b.navn AS brugergruppenavn, ansatdato, m.editor FROM medarbejdere AS m "_
                    &"LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) "_
                    &"LEFT JOIN brugergrupper AS b ON (b.id = m.brugergruppe) WHERE "& sqlsearchKri &" "& vispasoglukKri &" "& strSQLAdminRights &" ORDER BY mnavn LIMIT 4000" 
        
                    'response.write strSQL
                    'response.flush

                    oRec.open strSQL, oConn, 3
        
                    while not oRec.EOF
                    
                       if len(trim(request("jq_visalle"))) <> 0 then
                      jq_vispasluk = request("jq_visalle")
                      else
                      jq_vispasluk = 0
                      end if

                      if jq_vispasluk <> "1" then
                      visAlleSQLval = " AND mansat = 1 "
                      else
                      visAlleSQLval = " AND mansat <> -1 "
                      end if

                       

                       mtypenavn = oRec("mtypenavn")
                       mBrugergruppe = oRec("brugergruppenavn")
                   
                       call mstatus_lastlogin

                          
                     if media <> "eksport" then
                       
                       
                       if cdbl(lastmedid) = cdbl(oRec("mid")) then
                       trBgCol = "#FFFFE1" 
                       else
                       trBgCol = ""
                       end if%>

                        <tr style="background-color:<%=trBgCol%>;">
                            
                            <td>
                                 <% 
                                  editOk = 0   
                                  if cint(meCreate_newemployee) = 1 AND level <> 1 then 'Hvis opret medab = OK må man godt søge på alle, men kun redigere i dem man selv har oprettet

                                        if oRec("editor") = session("user") then
                                        editOk = 1
                                        end if

                                 else 

                                     if cint(level) = 1 then
                                      editOk = 1
                                     else
                                      editOk = 0
                                     end if

                                 end if

                                     'Response.write editOk & " " & meCreate_newemployee & " session(user): " & session("user") & " level: " & level & " editor: "& oRec("editor")

                                 if cint(editOk) = 1 then%>
                                  <a href="medarb.asp?func=red&id=<%=oRec("mid") %>"><%=left(oRec("mnavn"), 30) %>
                                 <%else %>
                                  <%=left(oRec("mnavn"), 30) %>
                                 <%end if %>

                                   <%if len(trim(oRec("init"))) <> 0 then%>
                                   &nbsp;[<%=oRec("init") %>]
                                   <%end if %>
                                
                                </a></td>
                            <td><%=mStatus %></td>
                            <td><%=left(mtypenavn, 15) %></td>
                            <td><%=left(mBrugergruppe, 20) %></td>
                            <td><a href="mailto:<%=oRec("email") %>"><%=oRec("email") %></a></td>
                            <td><%=oRec("ansatdato") %></td>
                            <td><%=lastLoginDateFm%></td>
                            <td style="text-align:center;"> <a href="../timereg/joblog.asp?menu=timereg&FM_medarb=<%=oRec("Mid")%>&FM_job=0&selmedarb=<%=oRec("Mid")%>" target="_blank"><span class="fa fa-external-link"></span></a></td>
                            <td style="text-align:center;">
                                <%if cint(level) = 1 then %>
                                <a href="medarb.asp?menu=tok&func=slet&id=<%=oRec("mid")%>"><span style="color:darkred;" class="fa fa-times"></span></a>
                                <%else %>
                                &nbsp;
                                <%end if %>
                            </td>
                        </tr>
               
                  <% end if 'media
                      
                      ekspTxt = ekspTxt & oRec("mnavn") & ";" & oRec("mnr") & ";" & oRec("init") & ";" & mStatus & ";" & oRec("mtypenavn") & ";" & oRec("brugergruppenavn") & ";" & oRec("email") & ";"& oRec("ansatdato") &"; xx99123sy#z"
                      oRec.movenext
                      wend
                      
                      oRec.close 
                      
                      
                      
             if media <> "eksport" then          
            %>


            </tbody>
               <tfoot>
                   <tr>
                    <th><%=medarb_txt_102 %></th>
                    <th><%=medarb_txt_103 %></th>
                    <th><%=medarb_txt_104 %></th>
                    <th><%=medarb_txt_105 %></th>
                    <th><%=medarb_txt_106 %></th>
                    <th><%=medarb_txt_114 %></th>
                    <th><%=medarb_txt_108 %></th>
                    <th><%=medarb_txt_109 %></th>
                    <th><%=medarb_txt_110 %></th>
                   </tr>
               </tfoot>
          </table>

        </div> <!-- /.portlet-body -->
        </section>
        <%end if %>



<%
'************************************************************************************************************************************************
'******************* Eksport **************************' 

                if media = "eksport" then

    
   
    
                    call TimeOutVersion()
    
    
                    ekspTxt = ekspTxt 'request("FM_ekspTxt")
	                ekspTxt = replace(ekspTxt, "xx99123sy#z", vbcrlf)
	
	
	
	                filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	                filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				                Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				                if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\medarb.asp" then
					                Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                else
					                Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					                Set objNewFile = nothing
					                Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				                end if
				
				
				
				                file = "medarbexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				                '**** Eksport fil, kolonne overskrifter ***
				
			
				                'strOskrifter = "Medarbejder; Nr.; Init; Status; Medarbejertype; Brugergruppe; Email; Ansatdato"
				                strOskrifter = medarb_txt_002&"; "& medarb_txt_111&"; "& medarb_txt_126&"; "& medarb_txt_127&"; "& medarb_txt_128&"; "& medarb_txt_112&"; "& medarb_txt_106&"; "& medarb_txt_114&"; "
				
				
				
				                objF.WriteLine(strOskrifter & chr(013))
				                objF.WriteLine(ekspTxt)
				                objF.close
				
				                %>
				                <div style="position:absolute; top:100px; left:200px; width:300px;">
	                            <table border=0 cellspacing=1 cellpadding=0 width="100%">
	                            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	                            <img src="../ill/outzource_logo_200.gif" />
	                            </td>
	                            </tr>
	                            <tr>
	                            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
                
                             
	                            <a href="../inc/log/data/<%=file%>" target="_blank" ><%=medarb_txt_115 %> >></a>
	                            </td></tr>
	                            </table>
                                </div>
	            
	          
	            
	                            <%
                
                
                                Response.end
	                            'Response.redirect "../inc/log/data/"& file &""	
				



                end if  
 '************************************************************************************************************************************************
              
          %>




          <section>
                <div class="row">
                     <div class="col-lg-12">
                        <b><%=medarb_txt_116 %></b>
                        </div>
                    </div>
                    <form action="medarb.asp?media=eksport" method="Post" target="_blank">
                  
                    <div class="row">
                     <div class="col-lg-12 pad-r30">
                         
                    <input id="Submit5" type="submit" value="<%=medarb_txt_117 %>" class="btn btn-sm" /><br />
                    <!--Eksporter viste kunder og kontaktpersoner som .csv fil-->
                         
                         </div>


                </div>
                </form>
                
            </section>

      </div><!-- /.portlet -->
    </div> <!-- /.container -->








<%end select %>

  </div> <!-- .content -->
</div> <!-- /#wrapper -->


<!--#include file="../inc/regular/footer_inc.asp"-->