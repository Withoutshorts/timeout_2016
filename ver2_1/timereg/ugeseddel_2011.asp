<%response.buffer = true %>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/treg_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->


<%
 '**** S�gekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_sogjobogkunde"


                '*** S�G kunde & Job            
                
                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                jobkundesog = request("jq_newfilterval")
                else
                filterVal = 0
                jobkundesog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")

                if filterVal <> 0 then
            
                 lastKid = 0
                
                  strJobogKunderTxt = strJobogKunderTxt &"<span style=""color:red; font-size:9px; float:right;"" class=""luk_jobsog"">[X]</span>"    
                         


                strSQL = "SELECT j.id AS jid, j.jobnavn, j.jobnr, j.jobstatus, k.kkundenavn, k.kkundenr, k.kid FROM timereg_usejob AS tu "_ 
                &" LEFT JOIN job AS j ON (j.id = tu.jobid) "_
                &" LEFT JOIN kunder AS k ON (k.kid = j.jobknr) "_
                &" WHERE tu.medarb = "& medid &" AND (j.jobstatus = 1 OR j.jobstatus = 3) AND "_
                &" (jobnr LIKE '"& jobkundesog &"%' OR jobnavn LIKE '"& jobkundesog &"%' OR "_
                &" kkundenavn LIKE '"& jobkundesog &"%' OR kkundenr = '"& jobkundesog &"' OR k.kinit = '"& jobkundesog &"')  AND kkundenavn <> ''"_
                &" GROUP BY j.id ORDER BY kkundenavn, jobnavn LIMIT 50"       
    

                 'response.write "strSQL " &strSQL
                 'response.end

                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
                if lastKid <> oRec("kid") then
                strJobogKunderTxt = strJobogKunderTxt &"<br><br><b>"& oRec("kkundenavn") &" "& oRec("kkundenr") &"</b><br>"
                end if 
                 
                strJobogKunderTxt = strJobogKunderTxt & "<input type=""hidden"" id=""hiddn_job_"& oRec("jid") &""" value="""& oRec("jobnavn") & " ("& oRec("jobnr") &")"">"
                strJobogKunderTxt = strJobogKunderTxt & "<input type=""checkbox"" class=""chbox_job"" id=""chbox_job_"& oRec("jid") &""" value="& oRec("jid") &"> "& oRec("jobnavn") & " ("& oRec("jobnr") &")" &"<br>" 
                
                lastKid = oRec("kid") 
                oRec.movenext
                wend
                oRec.close

              


                    '*** ��� **'
                    call jq_format(strJobogKunderTxt)
                    strJobogKunderTxt = jq_formatTxt


                    response.write strJobogKunderTxt

                end if    


        
          case "FN_sogakt"

               
                '*** S�g Aktiviteter 
                

                if len(trim(request("jq_newfilterval"))) <> 0 then
                filterVal = 1 
                aktsog = request("jq_newfilterval")
                else
                filterVal = 0
                aktsog = "6xxxxxfsdf554"
                end if
        
                medid = request("jq_medid")
                aktid = request("jq_aktid")
    
                if len(trim(request("jq_jobid"))) <> 0 then        
                jobid = request("jq_jobid")
                else
                jobid = 0
                end if

                'positiv aktivering
                pa = 0
                if len(trim(request("jq_pa") )) <> 0 then
                pa = request("jq_pa") 
                else
                pa = 0
                end if
        
                'pa = 0
            

                if filterVal <> 0 then
            
                 
    
                strAktTxt = strAktTxt &"<span style=""color:#999999; font-size:9px; float:right;"" class=""luk_aktsog"">X</span>"    
                         

               if pa = "1" then
               strSQL= "SELECT a.id AS aid, navn AS aktnavn FROM timereg_usejob LEFT JOIN aktiviteter AS a ON (a.id = tu.aktid) "_
               &" WHERE tu.medarb = "& usemrn &" AND tu.jobid = "& jobid &" AND aktid <> 0 AND a.navn LIKE '%"& aktsog &"%' AND aktstatus = 1 ORDER BY navn"   


               else


                        '*** Finder medarbejders projektgrupper 
                        '** Medarbejder projektgrupper **'
                        medarbPGrp = "#0#" 
                        strMpg = "SELECT projektgruppeId, medarbejderId, teamleder FROM progrupperelationer WHERE medarbejderId = "& medid & " GROUP BY projektgruppeId"

                        oRec5.open strMpg, oConn, 3
                        while not oRec5.EOF
                        medarbPGrp = medarbPGrp & ",#"& oRec5("projektgruppeId") &"#"         
        
                        oRec5.movenext
                        wend
                        oRec5.close 


           




               strSQL= "SELECT a.id AS aid, navn AS aktnavn, projektgruppe1, projektgruppe2, projektgruppe3, "_
               &" projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter AS a "_
               &" WHERE a.job = " & jobid & " AND navn LIKE '%"& aktsog &"%' AND aktstatus = 1 ORDER BY navn"      
    

               
            
                end if

                 'response.write "strSQL " &strSQL
                 'response.end


                oRec.open strSQL, oConn, 3
                while not oRec.EOF
        
                 showAkt = 0
                if instr(medarbPGrp, "#"& oRec("projektgruppe1") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe2") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe3") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe4") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe5") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe6") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe7") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe8") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe9") &"#") <> 0 _
                OR instr(medarbPGrp, "#"& oRec("projektgruppe10") &"#") <> 0 then
                showAkt = 1
                end if 


                
                
                if cint(showAkt) = 1 then 
                 
                strAktTxt = strAktTxt & "<input type=""hidden"" id=""hiddn_akt_"& oRec("aid") &""" value="""& oRec("aktnavn") &""">"
                strAktTxt = strAktTxt & "<input type=""checkbox"" class=""chbox_akt"" id=""chbox_akt_"& oRec("aid") &""" value="& oRec("aid") &"> "& oRec("aktnavn") &"<br>" 
                
                end if
                
                oRec.movenext
                wend
                oRec.close

              


                    '*** ��� **'
                    call jq_format(strAktTxt)
                    strAktTxt = jq_formatTxt


                    response.write strAktTxt

                end if    




        end select
        response.end
        end if




tloadA = now
 

if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	
	else
	
	function SQLBless(s)
		dim tmp
		tmp = s
		tmp = replace(tmp, ",", ".")
		SQLBless = tmp
	end function
	
	func = request("func")
	id = request("id")
	thisfile = "ugeseddel_2011.asp"
	media = request("media")




    if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> 0 then
    nomenu = 1
    else
    nomenu = 0
    end if

    if len(trim(request("usemrn"))) <> 0 then
    usemrn = request("usemrn")
    else
    usemrn = session("mid")
    end if
    


            
         

    level = session("rettigheder")
 
    rdir = request("rdir")
    select case rdir
    case "logindhist"
    rdirfile = "logindhist_2011.asp"
    case "godkenduge"
    rdirfile = "godkenduge.asp"
    case else
	rdirfile = "ugeseddel_2011.asp"
	end select
    

    if len(trim(request("medarbsel_form"))) <> 0 then

        'response.write "HER"
        'response.end

    if len(trim(request("FM_visAlleMedarb"))) <> 0 then
    visAlleMedarbCHK = "CHECKED"
    visAlleMedarb = 1
    else
    visAlleMedarbCHK = ""
    visAlleMedarb = 0
    end if

    else

                    if request.cookies("tsa")("visAlleMedarb") = "1" then
                    visAlleMedarbCHK = "CHECKED"
                    visAlleMedarb = 1
                    else
                    visAlleMedarbCHK = ""
                    visAlleMedarb = 0
                    'usemrn = session("mid")
                    end if
    end if
    response.cookies("tsa")("visAlleMedarb") = visAlleMedarb



    
    call ersmileyaktiv()
	

    call medarb_teamlederfor
    
        
    
        
    varTjDatoUS_man = request("varTjDatoUS_man")

    'if len(trim(request("varTjDatoUS_son"))) <> 0 then
	'varTjDatoUS_son = request("varTjDatoUS_son")
    'else
    varTjDatoUS_son = dateAdd("d", 6, varTjDatoUS_man)
    'end if

    'Response.Write varTjDatoUS_man
    'Response.end

    datoMan = day(varTjDatoUS_man) &"/"& month(varTjDatoUS_man) &"/"& year(varTjDatoUS_man)
    datoSon = day(varTjDatoUS_son) &"/"& month(varTjDatoUS_son) &"/"& year(varTjDatoUS_son)

    next_varTjDatoUS_man = dateadd("d", 7, varTjDatoUS_man)
	next_varTjDatoUS_son = dateadd("d", 7, varTjDatoUS_son)
    next_varTjDatoUS_man = year(next_varTjDatoUS_man) &"/"& month(next_varTjDatoUS_man) &"/"& day(next_varTjDatoUS_man)
	next_varTjDatoUS_son = year(next_varTjDatoUS_son) &"/"& month(next_varTjDatoUS_son) &"/"& day(next_varTjDatoUS_son)


    prev_varTjDatoUS_man = dateadd("d", -7, varTjDatoUS_man)
	prev_varTjDatoUS_son = dateadd("d", -7, varTjDatoUS_son)
    prev_varTjDatoUS_man = year(prev_varTjDatoUS_man) &"/"& month(prev_varTjDatoUS_man) &"/"& day(prev_varTjDatoUS_man)
	prev_varTjDatoUS_son = year(prev_varTjDatoUS_son) &"/"& month(prev_varTjDatoUS_son) &"/"& day(prev_varTjDatoUS_son)





    call positiv_aktivering_akt_fn()
    call smileyAfslutSettings()

	'Response.Write "media " & media
	
	
	'*** S�tter lokal dato/kr format. *****
	Session.LCID = 1030
	
	
    select case func 
	case "-"

    case "opdaterstatus"


        
        'response.write "request(ids)" & request("ids")
        
          ujid = split(request("ids"), ",")

                if len(trim(request("FM_godkendt"))) <> 0 then
                uGodkendt = request("FM_godkendt")
                else
                uGodkendt = 0
                end if
	
    for u = 0 to UBOUND(ujid)
	
	            editor = session("user")


                'if len(trim(request("FM_godkendt_"& trim(ujid(u))))) <> 0 then
                'uGodkendt = request("FM_godkendt_"& trim(ujid(u)))
                'else
                'uGodkendt = 0
                'end if
              

				strSQL = "UPDATE timer SET godkendtstatus = "& uGodkendt &", "_
				&"godkendtstatusaf = '"& editor &"' WHERE tid = " & ujid(u)
				
			    'Response.write strSQL &"<br>"
				
				oConn.execute(strSQL)
				
	next

       'response.end

    'Response.Redirect "ugeseddel_2011.asp?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&varTjDatoUS_son="&varTjDatoUS_son&"&nomenu="&nomenu

    Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	
	case "godkendugeseddel"
	
	'*** Godkender ugeseddel ***'
	
	    varTjDatoUS_manSQL = year(varTjDatoUS_man)&"/"&month(varTjDatoUS_man)&"/"&day(varTjDatoUS_man)
        varTjDatoUS_sonSQL = year(varTjDatoUS_son)&"/"&month(varTjDatoUS_son)&"/"&day(varTjDatoUS_son)
	    
	    strSQLup = "UPDATE timer SET godkendtstatus = 1, godkendtstatusaf = '"& session("user") &"' WHERE tmnr = "& usemrn 
	    if cint(SmiWeekOrMonth) = 0 then
        strSQLup = strSQLup & " AND tdato BETWEEN '"& varTjDatoUS_manSQL &"' AND '" & varTjDatoUS_sonSQL & "'" 
        else
        varTjDatoUS_man_mth = datepart("m", varTjDatoUS_man,2,2)
        strSQLup = strSQLup & " AND MONTH(tdato) = '"& varTjDatoUS_man_mth & "'" 
        end if

        strSQLup = strSQLup & " AND godkendtstatus <> 1" 

	    oConn.execute(strSQLup)
	    
        'if session("mid") = 1 then
	    'Response.Write strSQLup
	    'Response.end
        'end if
        
        
        '*** Godkend uge status ****'
        call godekendugeseddel(thisfile, session("mid"), usemrn, varTjDatoUS_man)
	    
	
	
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	

    case "afvisugeseddel"
	
	'*** Afviser ugeseddel ***'
    if len(trim(request("FM_afvis_grund"))) <> 0 then
    txt = replace(request("FM_afvis_grund"), "'", "")
    else
    txt = ""
    end if

	call afviseugeseddel(thisfile, session("mid"), usemrn, varTjDatoUS_man, varTjDatoUS_son, txt)
	    
	
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	
    
    
    case "adviserugeafslutning"

    call afslutugereminder(thisfile, session("mid"), usemrn, varTjDatoUS_man, varTjDatoUS_son, txt)

  
	Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son
	

    case "slet_tip"

        %>
        	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
        <%

        id = request("id")

        slttxt = "Du er ved at slette en timeregistrering uden match.<br> Er Du sikker p� Du vil g�re dette?"
        slturl = rdirfile & "?func=slet_tip_ok&id="&id&"&usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son

       call sltque(slturl,slttxt,slturlalt,slttxtalt,110,90)

    

    case "slet_tip_ok"

        id = request("id")

        sqlInTimerImpTemp = "DELETE FROM timer_import_temp WHERE id = "& id
        oConn.Execute(sqlInTimerImpTemp)

    Response.Redirect rdirfile & "?usemrn="&usemrn&"&varTjDatoUS_man="&varTjDatoUS_man&"&showadviseringmsg=1&nomenu="&nomenu '&"&varTjDatoUS_son="&varTjDatoUS_son

    case else

   

    call akttyper2009(2)
	
	
	
	if media <> "export" then
	
	
	if media <> "print" then
	
	
	%>

	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <SCRIPT language=javascript src="inc/smiley_jav.js"></script>
    <SCRIPT language=javascript src="inc/ugeseddel_2011_jav.js"></script>


    




     <%
          if media <> "print" then
    %>

     <div id="loadbar" style="position:absolute; display:; visibility:visible; top:160px; left:300px; width:300px; background-color:#ffffff; border:10px #cccccc solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	</td></tr></table>

	</div>

    <% end if
         
         
         
         call browsertype()
    


        
        if browstype_client <> "ip" then
            
            
	  
         
         
                 if cint(nomenu) <> 1 then
         
                   leftPos = 90
	            topPos = 102
         
         
                call menu_2014() %>
	
	
	              <% 
                   else   
              
                        leftPos = 20
	                    topPos = 20

                   end if


          else

          
	    leftPos = 20
	    topPos = 99

          end if

    
    else 
	
	leftPos = 20
	topPos = 20
	
	%>
	<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	<%end if 'print%>
	
	<%end if 'eksport
        
        
        %>
	
	
	
	<!-------------------------------Sideindhold------------------------------------->
	<div id="sindhold" style="position:absolute; left:<%=leftPos%>px; top:<%=topPos%>px; visibility:visible;">
      
        <%if cint(stempelurOn) = 1 then 
            lft = 710
        else
            lft = 800
        end if%>
          
    <div style="position:absolute; background-color:#ffffff; border:0px #5582d2 solid; padding:4px; width:75px; top:-20px; left:<%=lft%>px; z-index:0;"><a href="<%=lnkTimeregside%>" class="vmenu"><%=left(tsa_txt_031, 7) %>.</a></div>
    
    <%if cint(stempelurOn) = 1 then %>
    <div style="position:absolute; background-color:#ffffff; border:0px #5582d2 solid; padding:4px; top:-20px; width:95px; left:800px; z-index:0;"><a href="<%=lnkLogind%>" class="vmenu"><%=tsa_txt_340 %></a></div>
    <%end if
        
         
    tTop = 2
	tLeft = 0
	tWdth = 900
	
                 

	
	call tableDiv(tTop,tLeft,tWdth)    
            
            
    if media <> "print" AND len(trim(strSQLmids)) > 0 then 'Hvis man er level 1 eller teamleder vil len(trim(strSQLmids)) ALTID V�RE > 16 %>
	<form id="filterkri" method="post" action="ugeseddel_2011.asp">
        <input type="hidden" name="FM_sesMid" id="FM_sesMid" value="<%=session("mid") %>">
        <input type="hidden" name="medarbsel_form" id="medarbsel_form" value="1">
        <input type="hidden" name="nomenu" id="nomenu" value="<%=nomenu %>">
        <input type="hidden" name="varTjDatoUS_man" id="varTjDatoUS_man" value="<%=varTjDatoUS_man %>">
        
        <table>
       <tr bgcolor="#ffffff">
	<td valign=top> <b><%=tsa_txt_077 %>:</b> <br />
    <input type="CHECKBOX" name="FM_visallemedarb" id="FM_visallemedarb" value="1" <%=visAlleMedarbCHK %> /> <%=tsa_txt_388 %> (<%=tsa_txt_357 %>)
   
	<br />
				<% 
				call medarb_vaelgandre
                %>
        </td>
           </tr></table>
        </form>
       
	
	<%
   
       end if 'media
	
        if datePart("d", now, 2,2) < 4 OR datePart("d", now, 2,2) > 29 then 

	   if cint(smilaktiv) <> 0 AND media <> "print" then
        %> <br /><br /><div style="width:600px;"><%call afslutMsgTxt %>
        </div><%
            end if

        end if
    

    
   ugeseddelvisning = 1

   perInterval = 6 'dateDiff("d", varTjDatoUS_man, varTjDatoUS_son, 2,2) 
   perIntervalLoop = perInterval

   'response.write "perIntervalLoop " & perIntervalLoop
   for l = 0 to perIntervalLoop 
        
        if l = 0 then
        varTjDatoUS_use = varTjDatoUS_man
        showheader = 1
        showtotal = 0
        else
        varTjDatoUS_use = dateAdd("d", l, varTjDatoUS_man)
        showheader = 0
        showtotal = 0
         end if 

       if l = perIntervalLoop then
       showtotal = 1
       end if

         varTjDatoUS_son = varTjDatoUS_man

        call ugeseddel(usemrn, varTjDatoUS_use, varTjDatoUS_use, ugeseddelvisning, showtotal, showheader)  
   
   next
	
	
    if browstype_client <> "ip" then

        if media <> "print" then

        %>
        <div style="position:absolute; top:0px; left:945px; width:200px; background-color:#FFFFFF; padding:0px 20px 20px 20px;">
        <h4>Ugeseddel Resum�<br /><span style="font-size:9px; font-weight:lighter;">Normtimer, realiseret og timer uden match</span></h4>

        <% call normRealGrafWeekPage(usemrn, varTjDatoUS_man)  %>
            



     

        </div><!-- ugesededel resume -->

        <% end if ''IP


        

    if media <> "print" then '<> print

ptop = 0
pleft = 1200
pwdt = 220

call eksportogprint(ptop,pleft, pwdt)
%>

<form action="ugeseddel_2011.asp?usemrn=<%=usemrn%>&varTjDatoUS_man=<%=varTjDatoUS_man %>&varTjDatoUS_son=<%=varTjDatoUS_son %>&media=print" method="post" target="_blank">
<tr> 
    <td valign=top align=center>
   <input type=image src="../ill/printer3.png" />
    </td>
    <td class=lille><input id="Submit5" type="submit" value=" Print venlig >> " style="font-size:9px; width:130px;" /></td>
</tr>
</form>
    <tr><td colspan="2">

           <!-- hvis level = 1 OR teamleder -->
        <%if (level <=2 OR level = 6) AND cint(smilaktiv) <> 0 then %>
            <br /><br />

        <%
        '** Tjekker for Uge 53. SKAL i virkeligheden v�re om s�ndag er i et andet �r end mandag - da �r s� skifter.
        if datepart("ww", varTjDatoUS_son, 2,2) = 53 then
        tjkAar = year(varTjDatoUS_son) + 1
        else
        tjkAar = year(varTjDatoUS_son)
        end if    
            
        call erugeAfslutte(tjkAar, datepart("ww", varTjDatoUS_son, 2,2), usemrn) %>
        <%call godkendugeseddel(fmlink, usemrn, varTjDatoUS_man, rdir) %>
    
        <%end if 'level %>
    

             </td>
        </tr>



</table>

      

      

</div>
<%else

Response.Write("<script language=""JavaScript"">window.print();</script>")

end if 
    
%>

  

  
    <%end if 'print%>

  </div><!--tablediv-->
    <br /><br />&nbsp;


    </div><!--sidediv-->
    <br /><br />&nbsp;
    <%
	
	
	end select
	
	
	end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->
