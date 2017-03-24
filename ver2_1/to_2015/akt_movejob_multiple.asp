

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->


<%

    func = request("func")
	if func = "dbopr" then
	id = 0
	else
	id = request("id")
	end if
	
		function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, "'", "''")
		SQLBless = tmp
		end function


    call menu_2014()

    %>

       

    <div id="wrapper">
        <div class="content">    

    <%
    select case func


    case "flyt"
	'*** Her indsættes en ny type i db ****
		if len(request("FM_aid")) = 0 OR len(trim(request("FM_newakt"))) = 0 OR isdate(request("FM_fradato") = false) then
		
		errortype = 186
		call showError(errortype)
		
		else
		

                    aid = request("FM_aid") 'Eksisterende
                    newakt = request("FM_newakt") 'akt eller job (ved flyt)

                    
                    fraDato = request("FM_fradato")
                    fraDatoSQL = year(fraDato) &"/"& month(fraDato) &"/"& day(fraDato) 

                    forefter = request("forefter")


                    jobnavn = "Job Not found"
                    jobnr = 0
                    kkundenavn = "Customer Not found on update multible jobs"
                    kknr = 0

                    '**** NYT JOB ****
                    if newakt < 0 then 
                    newjobid = newakt * -1    
                    strSQLj = "SELECT j.id, jobnavn, jobnr, kkundenavn, kkundenr, kid FROM job j "_
                    &" LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE j.id = "& newjobid
                    else
                    strSQLj = "SELECT jobnavn, jobnr, kkundenavn, kkundenr, kid, fakturerbar FROM aktiviteter a "_
                    &" LEFT JOIN job j ON (j.id = a.job) "_
                    &" LEFT JOIN kunder k ON (k.kid = j.jobknr) WHERE a.id = "& newakt
                    end if
                    'Response.write strSQLj
                    'Response.flush
        
                    oRec.open strSQLj, oConn, 3
                    if not oRec.EOF then


                    jobnavn = replace(oRec("jobnavn"), "'", "''")
                    jobnr = oRec("jobnr")
                    kkundenavn = replace(oRec("kkundenavn"), "'", "''")
                    kknr = oRec("kkundenr")
                    kid = oRec("kid")

                    if newakt > 0 then 
                    fakturerbar = oRec("fakturerbar")
                    end if            

                    end if
                    oRec.close

                

                    '*** Flyt aktivitet *****'
                    if newakt < 0 then            
                    newakt = newakt * -1

                    strSQLjob = "UPDATE aktiviteter SET "_
					& " job = "& newakt &""_
				    & " WHERE id = "& aid
                								
				    'Response.write strSQLtimer
					'Response.flush
                								
					oConn.execute(strSQLjob)
            
                                '*** Flytter ALLE timer
                                '*** Opdaterer timereg tabellen ****
					            strSQLtimer = "UPDATE timer SET "_
					            & " Tknavn = '"& kkundenavn &"', Tknr = "& kid &", tjobnr = '"& jobnr &"', tjobnavn = '"& jobnavn &"'"_
				                & " WHERE taktivitetid = "& aid 
                								
				                'Response.write strSQLtimer
					            'Response.flush
                								
					            oConn.execute(strSQLtimer)


                    else        '*** MERGE / Sammenlæg


                              '*** Flytter timer pr. dato
                                '*** Opdaterer timereg tabellen ****
					            strSQLtimer = "UPDATE timer SET "_
					            & " Tknavn = '"& kkundenavn &"', Tknr = "& kid &", tjobnr = '"& jobnr &"', tjobnavn = '"& jobnavn &"', taktivitetid = "& newakt &", tfaktim = "& fakturerbar &""_
				                & " WHERE taktivitetid = "& aid & " AND tdato "& forefter &" '"& fraDatoSQL &"'"
                								
				                'Response.write strSQLtimer
					            'Response.flush
                								
					            oConn.execute(strSQLtimer)

                    end if


					
								                
					
                
        		
					
		

		
		Response.redirect "akt_movejob_multiple.asp"
		end if


    case else
	'*** Her indlæses form til rediger/oprettelse af ny type ***
	




%>


<div class="container">
    <div class="portlet">
            <h3 class="portlet-title">
            <u>Flyt aktivitet og/eller timer til nyt job eller aktivitet</u>
            </h3>

        <form action="akt_movejob_multiple.asp?func=flyt" method="post">
             <input type="hidden" name="id" value="<%=id%>">
        <div class="row">
            <div class="col-lg-10">&nbsp</div>
            <div class="col-lg-2 pad-b10">
            <button type="submit" class="btn btn-success btn-sm pull-right"><b>Opdatér</b></button>
            </div>
        </div>

        <div class="portlet-body">
            <section>
                <div class="well well-white">
           
           
          

      
             <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-6">
                    Flyt/Sammenlæg aktivitet og timer til valgte job/aktivitet.<br />
                    Ved sammenlæg beholdes aktivitet på det oprindelig job. Ved opret ny, flyttes aktiviteten og alle timer. 
                    Projektgrupper forbliver uændret.
                    <br /><br /><b>Vælg aktivitet:</b> (på aktive job - vælg 1)<br />
                    <%

                    whSQL = "(jobstatus = 1)"
                    strSQLjob = "SELECT jobnavn, jobnr, j.id AS jobid, kid, kkundenavn, kkundenr, a.navn AS aktnavn, a.id AS aktid, FORMAT(SUM(timer), 2) AS sumtimer FROM aktiviteter a "_
		            &" LEFT JOIN job j ON (j.id = a.job) "_
                    &" LEFT JOIN kunder ON (kid = jobknr) "_
                    &" LEFT JOIN timer ON (taktivitetid = a.id) "_
                    &" WHERE "& whSQL &" GROUP BY taktivitetid ORDER BY kkundenavn, jobnavn, a.navn"
		
		            'Response.Write strSQLjob
		            'Response.flush
		
		            %>
		
            <select name="FM_aid" size=10 style="width:700px;">
          
              
		        <%
		
		        oRec2.open strSQLjob, oConn, 3
		        lastkid = 0
                antaljob = 0
                lastjid = 0
		        while not oRec2.EOF


                    if lastkid <> oRec2("kid") then

                        if antaljob <> 0 then
                        %>
                        <option DISABLED></option>
                        <%end if

                    %><option DISABLED><%=oRec2("kkundenavn") %> (<%=oRec2("kkundenr") %>)</option>
                    <%end if


                    if lastjid <> oRec2("jobid") then

                        if antaljob <> 0 then
                        %>
                        <option DISABLED></option>
                        <%end if

                    %><option DISABLED><%=oRec2("jobnavn") %> (<%=oRec2("jobnr") %>)</option>
                    <%end if %>
                        
                              

                 <%if isNull(oRec2("sumtimer")) <> true then
                     sumtimer = replace(replace(oRec2("sumtimer"), ",",""), ".", ",")
                 else
                     sumtimer = 0
                 end if

                     %>

		          <option value="<%=oRec2("aktid") %>"><%=oRec2("aktnavn") & " .......................... ("& sumtimer &" t.)"%></option>
           
		        <%
                lastjid = oRec2("jobid")
                lastKid = oRec2("kid")    
                antaljob = antaljob + 1 
		        oRec2.movenext
		        Wend
		        oRec2.close
		
		        %>
		         </select>
		         <br />
                Antal aktive job og tilbud: <%=antaljob %>
                    
                </div>
            </div>
            <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-6"><br /><br /><b>Flyt til job/akt.:</b> (sammenlæg: vælg 1, eller opret som ny)<br />
                    <select name="FM_newakt" size=10 style="width:700px;">
                  
                    <%   whSQL = "(jobstatus = 1)"
                    strSQLjobnew = "SELECT jobnavn, jobnr, j.id AS jobid, kid, kkundenavn, kkundenr, a.navn AS aktnavn, a.id AS aktid, FORMAT(SUM(timer), 2) AS sumtimer FROM aktiviteter a "_
		            &" LEFT JOIN job j ON (j.id = a.job) "_
                    &" LEFT JOIN kunder ON (kid = jobknr) "_
                    &" LEFT JOIN timer ON (taktivitetid = a.id) "_
                    &" WHERE "& whSQL &" GROUP BY taktivitetid ORDER BY kkundenavn, jobnavn, a.navn"
			        oRec2.open strSQLjobnew, oConn, 3
			        while not oRec2.EOF
				
                            if lastkid <> oRec2("kid") then

                                if antaljob <> 0 then
                                %>
                                <option DISABLED></option>
                                <%end if

                            %><option DISABLED><%=oRec2("kkundenavn") %> (<%=oRec2("kkundenr") %>)</option>
                            <%end if


                            if lastjid <> oRec2("jobid") then

                                if antaljob <> 0 then
                                %>
                                <option DISABLED></option>
                                <%end if

                            %><option DISABLED><%=oRec2("jobnavn") %> (<%=oRec2("jobnr") %>)</option>
                                 <option value="-<%=oRec2("jobid") %>">>> Opret som ny aktivitet</option>
                            <%end if %>
                        
                              
                         <%if isNull(oRec2("sumtimer")) <> true then
                     sumtimer = replace(replace(oRec2("sumtimer"), ",",""), ".", ",")
                 else
                     sumtimer = 0
                 end if

                     %>

		                  <option value="<%=oRec2("aktid") %>"><%=oRec2("aktnavn") & " .......................... ("& sumtimer &" t.)"%></option>
           
		                <%
                        lastjid = oRec2("jobid")
                        lastKid = oRec2("kid")    
                        antaljob = antaljob + 1 

			        oRec2.movenext
			        wend
			        oRec2.close %>
                    </select>
                </div>
            </div>
               
                    
                    <div class="row">
                <div class="col-lg-1">&nbsp</div>
                <div class="col-lg-6"><br /><br /><b>Flyt timer pr. dato</b> <br />
                Ved sammenlæg: Vælg fra/til hvilken dato timer skal flyttes, ved opret ny flyttes alle timer.
                    <br />
                   <select name="forefter">
                       <option value=">=">Efter</option>
                       <option value="<=">Før</option>
                       
                   </select> <input type="text" name="FM_fradato" value="<%=day(now) &"/"& month(now) &"/"& year(now) %>" />

                

            </div>
            </div>
                    
        </div>


             

         </section>

            
        </div>
        </form>
    </div>
</div>



    


<%end select  %>

</div>
</div>

<!--#include file="../inc/regular/footer_inc.asp"-->