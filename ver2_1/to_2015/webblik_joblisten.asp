<%response.buffer = true%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->

<!--#include file="../inc/regular/global_func.asp"-->


<%

'section for ajax calls
if Request.Form("AjaxUpdateField") = "true" then
Select Case Request.Form("control")
case "FM_sortOrder"
Call AjaxUpdate("job","risiko","")

case "FN_updaterplan"

 

    dothis = request("dothis")
    ab_id = 0 'request("id")
    ab_name = 0

    if session("lto") = "dencker" then
    ab_medid = 123 'Ufordelte mid Dencker
    else
    ab_medid = 0
    end if
    
    ab_aktid = request("aktid")
    ab_jobid = request("jobid")
    ab_serie = 0
    ab_date = year(now) & "/" & month(now) & "/" & day(now)
    ab_startdate = year(request("datothis")) & "/" & month(request("datothis")) & "/" & day(request("datothis")) & " 08:00:00"
    ab_enddate = ab_startdate
    ab_editor = session("user")

    'response.write "HER 2 "& ab_name &","& ab_medid &","& ab_aktid &","& ab_jobid &", dothis: "& dothis &","& ab_serie &","& ab_date &","& ab_startdate &","& ab_enddate &","& ab_editor &","& ab_date 
    'response.end

   '*** ��� **'
   call jq_format(ab_editor)
   ab_editor = jq_formatTxt

    if dothis = "1" then 
     '** INSERT 
    strSQLIns = "INSERT INTO akt_bookings (ab_name, ab_medid, ab_aktid, ab_jobid, ab_serie, ab_date, ab_startdate, ab_enddate, ab_editor, ab_editor_date) "_
    &" VALUES ("& ab_name &","& ab_medid &","& ab_aktid &","& ab_jobid &","& ab_serie &",'"& ab_date &"','"& ab_startdate &"','"& ab_enddate &"','"& ab_editor &"','"& ab_date &"')"

    'response.write strSQLIns
    'response.end
    oConn.Execute(strSQLIns)

    else

    strSQLDel = "DELETE FROM akt_bookings WHERE ab_jobid = " & ab_jobid & " AND ab_aktid = "& ab_aktid & " AND ab_startdate = '"& ab_startdate &"'"
    oConn.Execute(strSQLDel)

    end if
    

     '** Job Rekvnr ****'
     'strSQLUpdj = "UPDATE job SET rekvnr = '"& jobrekvnr &"' WHERE id = "& jobid 
     'oConn.Execute(strSQLUpdj)


case "FN_updaterekvnrkom"

        jobid = request("jobid") 
        jobrekvnr = request("newRekvnr")
        jobrekvnr = replace(jobrekvnr, "'", "''")
        
        '** Job Rekvnr ****'
        strSQLUpdj = "UPDATE job SET rekvnr = '"& jobrekvnr &"' WHERE id = "& jobid 
        oConn.Execute(strSQLUpdj)

case "FN_updatejobkom"
        
                 jobid = request("jobid")   
                 jobkom = request("jobkom") 
                 jobkom = replace(jobkom, "''", "") 
                 jobkom = replace(jobkom, "'", "")
                 jobkom = replace(jobkom, "<", "")
                 jobkom = replace(jobkom, ">", "")

                 oldKom = ""

                 strSQLUpdj = "SELECT kommentar FROM job WHERE id = "& jobid
                 oRec3.open strSQLUpdj, oConn, 3
                 if not oRec3.EOF then 
                 oldKom = oRec3("kommentar")
                 end if
                 oRec3.close

                 jobkom = "<span style=""color:#999999;"">" & formatdatetime(now, 2) &", "& session("user") &":</span>" & jobkom & "<br>"& oldKom
                
                 '** Job kommentar ****'
                 strSQLUpdj = "UPDATE job SET kommentar = '"& jobkom &"' WHERE id = "& jobid 
                 oConn.Execute(strSQLUpdj)
                 



case "FN_updatejobkom_ryd"
                

                 jobid = request("jobid")   
                 
                 '** Job kommentar ****'
                 strSQLUpdj = "UPDATE job SET kommentar = '' WHERE id = "& jobid 
                 oConn.Execute(strSQLUpdj)
                 

case "FN_updatejobstatus"
                

                 jobid = request("jobid")   
                 jobstatus = request("jobstatus")

                 medarbejderid = session("mid")
                 
                 '** Luk Job ****'
                 strSQLUpdj = "UPDATE job SET jobstatus = "& jobstatus &" WHERE id = "& jobid 
                 oConn.Execute(strSQLUpdj)

                
                 call lukjobmail(jobstatus, jobid, medarbejderid)



case "FN_updatejobprio"
                

                 jobid = request("jobid")   
                 jobprioitet = request("jobprioitet")

                 '** Job kommentar ****'
                 strSQLUpdj = "UPDATE job SET risiko = "& jobprioitet &" WHERE id = "& jobid 
                 oConn.Execute(strSQLUpdj)

case "FN_updatejovforkalktimer"
                

                 jobid = request("jobid")   
                 jobforklak = request("budgettimer")
                 
                 jobforklak = replace(jobforklak, ",",".")

                 '** Job kommentar ****'
                 strSQLUpdj = "UPDATE job SET budgettimer = "& jobforklak &", ikkebudgettimer = 0 WHERE id = "& jobid 
                 oConn.Execute(strSQLUpdj)


case "FN_updatejobvaluta"

                 jobid = request("jobid")   
                 jo_valuta = request("jo_valuta")
                 
            
                 '** Job valuta ****'
                 strSQLUpdj = "UPDATE job SET jo_valuta = "& jo_valuta &" WHERE id = "& jobid 
                 oConn.Execute(strSQLUpdj)

case "FN_updatejovbrutoms"
                

                 jobid = request("jobid")   
                 jobbruttooms = request("jo_bruttooms")
                 
                 jobbruttooms = replace(jobbruttooms, ",",".")

                 '** Job bruttooms ****'
                 strSQLUpdj = "UPDATE job SET jo_bruttooms = "& jobbruttooms &" WHERE id = "& jobid 
                 oConn.Execute(strSQLUpdj)

 case "FN_updatejobdato"
                

                 jobid = request("jobid")   
                 jobstDato = request("jobstDato")
                 jobslDato = request("jobslDato")             

                 '** Job kommentar ****'
                 strSQLUpdj = "UPDATE job SET jobstartdato = '"& jobstDato &"', jobslutdato = '"& jobslDato &"' WHERE id = "& jobid 
                 oConn.Execute(strSQLUpdj)

case "FN_updateest"            

                 jobid = request("jobid")   
                 rest = request("rest")
                 timerproc = request("timerproc")  
                 ddDato = year(now) &"/"& month(now) &"/"& day(now) &" "& datepart("h", now, 2,2) & ":" & datepart("n", now, 2,2) &":00"           

                 '** Job Restestimat / WIP ****'
                 strSQLUpdj = "UPDATE job SET restestimat = "& rest &", stade_tim_proc = "& timerproc &" WHERE id = "& jobid 
                 oConn.Execute(strSQLUpdj)

                 strSQLUpdjWiphist = "INSERT INTO wip_historik (dato, restestimat, stade_tim_proc, medid, jobid) VALUES ('"& ddDato &"', "& rest &", "& timerproc &", "& session("mid") &", "& jobid &")"
                 oConn.Execute(strSQLUpdjWiphist)


                strHTMLWIPhist = "WIP historik (seneste 5):<br />" 
                'Response.Write strHTMLWIPhist
                'response.end
                                
                                strSelWIp = "SELECT w.dato, medid, init, stade_tim_proc, restestimat FROM wip_historik w "_
                                &" LEFT JOIN medarbejdere ON (mid = medid) "_
                                &" WHERE jobid = "& jobid &" ORDER BY dato DESC limit 5"
                                oRec10.open strSelWip, oConn, 3
                                t = 0
                                while not oRec10.EOF 

                                if oRec10("stade_tim_proc") = 1 then
                                restestimatTxt = oRec10("restestimat") & " %"
                                else
                                restestimatTxt = oRec10("restestimat") & " t. rest"
                                end if
                                

                                if t = 0 then
                                strHTMLWIPhist = strHTMLWIPhist & "<span style=""color:orange"">" & left(oRec10("dato"), 16) &" <b>"& restestimatTxt &"</b> - "& oRec10("init") &"</span><br />"
                                else
                                strHTMLWIPhist = strHTMLWIPhist & left(oRec10("dato"), 16) &" <b>"& restestimatTxt &"</b> - "& oRec10("init") &"<br />"
                                end if


                               t = t + 1
                               oRec10.movenext
                               wend
                               oRec10.close


                  Response.Write strHTMLWIPhist

       


End select
Response.End
end if


%>
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../timereg/inc/dato2.asp"-->

<!--#include file="../timereg/inc/isint_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!--#include file="inc/webblik_joblisten_inc.asp"-->

<!--#include file="../inc/regular/topmenu_inc.asp"-->
<%


if len(trim(request("nomenu"))) <> 0 AND request("nomenu") <> "0" then
nomenu = 1
else
nomenu = 0
end if

func = request("func")

'****************************'
'*** Opdaterer job liste ****'
'****************************'

if func = "opdaterjobliste" then


ujid = split(request("FM_jobid"), ",")

'Response.write request("FM_jobid")
'Response.end

ustaar = split(request("FM_start_aar"))
ustmd = split(request("FM_start_mrd"))
ustdag = split(request("FM_start_dag"))

uslaar = split(request("FM_slut_aar"))
uslmd = split(request("FM_slut_mrd"))
usldag = split(request("FM_slut_dag"))

'uudgifter = split(request("FM_udgifterpajob"), ",")

    'Response.write request("FM_restestimat")
    'Response.end 

    if len(trim(request("FM_restestimat"))) <> 0 then
    urest = split(request("FM_restestimat"), ",")
    utimer_proc = split(request("FM_stade_tim_proc"), ",")
    opdEstimat = 1
    else
    urest = 0
    utimer_proc = 0
    opdEstimat = 0
    end if

urisiko = split(request("FM_risiko"), ",")
ustatus = split(request("FM_jobstatus"), ",")

'ukomm = split(trim(request("FM_kommentar")), ", #, ")
'nykomm = split(trim(request("FM_kommentar_ny")), ", #, ")



for u = 0 to UBOUND(ujid)
	'Response.write uudgifter(u) & "<br>"
	
	ustdato = replace(ustaar(u) & "/" & ustmd(u) & "/" & ustdag(u), ",", "")
	usldato = replace(uslaar(u) & "/" & uslmd(u) & "/" & usldag(u), ",", "")
	
	
			
			
    call erDetInt(trim(urisiko(u))) 
    if isInt > 0 OR instr(trim(urisiko(u)), ".") <> 0 then
    	

    	%>
		<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
		
		<%
    	
	    errortype = 123
	    call showError(errortype)

    isInt = 0
    Response.end

    else

     '*** S�t lukkedato (skal v�re f�r det skifter status) ***'
     call lukkedato(ujid(u), ustatus(u))
               
    
	strSQLjobupd = "UPDATE job SET jobstartdato = '"& ustdato &"', "_
	&" jobslutdato = '"& usldato &"', "
	if cint(opdEstimat) = 1 then
    strSQLjobupd = strSQLjobupd &" restestimat = "& urest(u) &", stade_tim_proc = "& utimer_proc(u) &", "
    end if
    
    if request("FM_sorter") <> 1 AND request("FM_sorter") <> 31 AND request("FM_sorter") <> 7 then
    strSQLjobupd = strSQLjobupd &" risiko = "& trim(urisiko(u)) &", "
    end if

	strSQLjobupd = strSQLjobupd &" jobstatus = "& ustatus(u) &" WHERE id = " & ujid(u)

	oConn.execute(strSQLjobupd)
	
	'Response.write strSQLjobupd & "<br>"
	'Response.end

        '**** SyncDatoer ****'
        jobnrThis = 0
        syncslutdato = 0
        strSQLj = "SELECT jobnr, syncslutdato FROM job WHERE id =  "& ujid(u)
        oRec5.open strSQLj, oConn, 3 
        if not oRec5.EOF then

        jobnrThis = oRec5("jobnr")
        syncslutdato = oRec5("syncslutdato")

        end if
        oRec5.close


     if ustatus(u) = 0 AND syncslutdato = 1 then
            call syncJobSlutDato(ujid(u), jobnrThis, syncslutdato)
     end if

	end if
	
next

'Response.Write "her"
'Response.flush
Response.Redirect "webblik_joblisten.asp"

end if



if len(trim(request("fromvmenu"))) <> 0 then
fromvmenu = request("fromvmenu")
else
fromvmenu = 0
end if

thisfile = "webblik_joblisten.asp"

if len(request("FM_parent")) <> 0 then
parent = request("FM_parent")
else
parent = 0
end if

if len(request("oldone")) <> 0 then
oldones = request("oldone")
else
oldones = 0
end if

if len(request("lastid")) <> 0 then
lastId = request("lastid")
else
lastId = 0
end if

print = request("print")



if len(request("realfakpertot")) <> 0 then
realfakpertot = request("realfakpertot")
Response.Cookies("webblik")("realfakpertot") = realfakpertot
else
    if request.Cookies("webblik")("realfakpertot") <> "" then
    realfakpertot = request.Cookies("webblik")("realfakpertot")
    else
    realfakpertot = 0
    end if
end if


if realfakpertot <> 0 then
realfakpertot1CHK = "CHECKED"
realfakpertot0CHK = ""
realfakpertotTxt = "Kun i periode"
else
realfakpertot1CHK = ""
realfakpertot0CHK = "CHECKED"
realfakpertotTxt = "Total, uanset valgt periode"
end if
            
 

if len(request("realfakpertot")) <> 0 then 'er der s�gt

if len(trim(request("historisk_wip"))) <> 0 then
historisk_wip = 1
else
historisk_wip = 0
end if

Response.Cookies("webblik")("historisk_wip") = historisk_wip
else
    if request.Cookies("webblik")("historisk_wip") <> "" then
    historisk_wip = request.Cookies("webblik")("historisk_wip")
    else
    historisk_wip = 0
    end if
end if


if historisk_wip <> 0 then
historisk_wipCHK = "CHECKED"
historisk_wipTxt = "Historisk WIP benyttet"

else
historisk_wipCHK = ""
historisk_wipTxt = "Aktuel WIP benyttet"
end if



if len(request("realfakpertot")) <> 0 then 'er der s�gt

if len(trim(request("timertilfak"))) <> 0 then
timertilfak = 1
else
timertilfak = 0
end if

Response.Cookies("webblik")("timertilfak") = timertilfak
else
    if request.Cookies("webblik")("timertilfak") <> "" then
    timertilfak = request.Cookies("webblik")("timertilfak")
    else
    timertilfak = 0
    end if
end if


if timertilfak <> 0 then
timertilfakCHK = "CHECKED"
timertilfakTxt = "Timer til fakturering"
else
timertilfakCHK = ""
timertilfakTxt = ""
end if


            


            if len(trim(request("FM_sorter_visprogrp"))) <> 0 then
            sorter_visprogrp = request("FM_sorter_visprogrp")
            else
                if request.Cookies("webblik")("sorter_visprogrp") <> "" then
                sorter_visprogrp = request.Cookies("webblik")("sorter_visprogrp")
                else
                sorter_visprogrp = "#-1#"
                end if
            end if

            Response.Cookies("webblik")("sorter_visprogrp") = sorter_visprogrp







   if len(request("st_sl_dato")) <> 0 then
    usedatoKri = request("st_sl_dato")  
    else
        if len(request.Cookies("webblik")("datokri")) <> 0 then
        usedatoKri = request.Cookies("webblik")("datokri")
        else
        usedatoKri = 1
        end if  
    end if
    
    response.Cookies("webblik")("datokri") = usedatoKri
    
    
    
   
    st_sl_DatoChk0 = "SELECTED"
    st_sl_DatoChk1 = ""
    st_sl_DatoChk2 = ""
    st_sl_DatoChk3 = ""
    st_sl_DatoChk6 = ""
    st_sl_DatoChk4 = ""
    st_sl_DatoChk5 = ""
    stsldatoSQLKri = "" 'jobstartdato DESC


    select case usedatoKri 
    case 2
 
    st_sl_DatoChk2 = "SELECTED"
    stsldatoSQLKri = "" 'jobstartdato DESC
    printVal = "- Periode ignoreret"

    case 1
   
    st_sl_DatoChk1 = "SELECTED"
    stsldatoSQLKri = "jobslutdato"
    printVal = "- Startdato"        
    
    case 3
   
    st_sl_DatoChk3 = "SELECTED"
    stsldatoSQLKri = "" 'jobstartdato DESC
    printVal = "- Timer, materialer eller fakturaer"

    case 4
  
    st_sl_DatoChk4 = "SELECTED"
    stsldatoSQLKri = "lukkedato" 'jobstartdato DESC
    printVal = "- Lukkedato"
    
    case 5
  
    st_sl_DatoChk5 = "SELECTED"
    stsldatoSQLKri = "" 'jobstartdato DESC
    printVal = "- Timer"

    case 6
   
    st_sl_DatoChk6 = "SELECTED"
    stsldatoSQLKri = "" 'jobstartdato DESC
    printVal = "- Fakturaer"

    case else
    st_sl_DatoChk0 = "SELECTED"
    stsldatoSQLKri = "jobstartdato"
    printVal = "- Slutdato"
    end select 


    if len(trim(request("FM_start_dato"))) <> 0 then

                    strMrd = datepart("m", request("FM_start_dato"), 2,2)
                    strDag = datepart("d", request("FM_start_dato"), 2,2)
                    strAar = datepart("yyyy", request("FM_start_dato"), 2,2) 
                    strDag_slut = datepart("d", request("FM_slut_dato"), 2,2)
                    strMrd_slut = datepart("m", request("FM_slut_dato"), 2,2)
                    strAar_slut = datepart("yyyy", request("FM_slut_dato"), 2,2)


    else

                    'if Request.Cookies("erp_datoer")("st_md") <> "" then

                    'strMrd = Request.Cookies("erp_datoer")("st_md")
		            'strDag = Request.Cookies("erp_datoer")("st_dag")
		            'strAar = Request.Cookies("erp_datoer")("st_aar") 
        		    
	                'strDag_slut = Request.Cookies("erp_datoer")("sl_dag")
	                'strMrd_slut = Request.Cookies("erp_datoer")("sl_md")
	                'strAar_slut = Request.Cookies("erp_datoer")("sl_aar")

                    'else

                    strMrd = month(now) '1
                    strDag = day(now)'1
                    strAar = year(now)

                    endDateinAMonth = dateAdd("m", 1, now)
                    strDag_slut = datepart("d", endDateinAMonth, 2,2)
                    strMrd_slut = datepart("m", endDateinAMonth, 2,2)
                    strAar_slut = datepart("yyyy", endDateinAMonth, 2,2)


                    'end if

    end if


               'response.Cookies("erp_datoer")("st_md") = strMrd
		       'response.Cookies("erp_datoer")("st_dag") = strDag 
		       'response.Cookies("erp_datoer")("st_aar") = strAar
        		    
	           'response.Cookies("erp_datoer")("sl_dag") = strDag_slut
	           'response.Cookies("erp_datoer")("sl_md") = strMrd_slut
	           'response.Cookies("erp_datoer")("sl_aar") = strAar_slut 
   



if len(session("user")) = 0 then
	
	errortype = 5
	call showError(errortype)
	else
	
	'*** S�tter lokal dato/kr format. Skal inds�ttes efter kalender.
	Session.LCID = 1030
	%>
	

   


	




	<%if print <> "j" then 
	
	
	
	%>
    <div id="loadbar" style="position:absolute; display:; visibility:visible; top:360px; left:220px; width:500px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	Forventet loadtid:
	<%
	
	exp_loadtid = 0
	'exp_loadtid = (((len(akttype_sel) / 3) * (len(antalvlgM) / 3)) / 50)  %> 
	ca.: <b>3-45 sek.</b><br /><br />
    Ved visning af mange job (+300), kan der v�re lidt ekstra ventetid efter siden er loadet. F�rst n�r denne boks forsvinder helt er siden klar.<br /><br />
    <a href="webblik_joblisten.asp?fromvmenu=1" style="color:red;">Stop pageload</a>
    <%'response.flush %>
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
	
	</td></tr></table>

	</div>
    <%end if



        if print <> "j" then
            call menu_2014() 
        end if
        %>

   


	

	<script src="js/webblik_jav3.js"></script>
	
	

	

      
	<% 

          call aktBudgettjkOn_fn()


          if len(request("FM_sorter")) <> 0 then
		sorter = request("FM_sorter")
		        select case sorter
		        case 1
		        orderBY = "risiko"
		        case 2
		        orderBY = "jobstartdato"
                case 3
                orderBY = "kkundenavn, jobnavn, jobnr"
                case 31
                orderBY = "kkundenavn, risiko"
		        case 4 
                orderBY = "jobnavn, jobnr"
                case 5 
                orderBY = "jobnr"
                case 6 
                orderBY = "jobnr DESC"
                case 7,8 
                orderBY =  "risiko" 'orderBY = "kkundenavn, jobnavn, jobnr"                
                case else
		        orderBY = "jobslutdato"
		        end select
		else
		    if request.cookies("webblik")("prioitet") <> "" then
		    sorter = request.cookies("webblik")("prioitet")
		        select case sorter
		        case 1
		        orderBY = "risiko"
		        case 2
		        orderBY = "jobstartdato"
                case 3
                orderBY = "kkundenavn, jobnavn, jobnr"
                case 31
                orderBY = "kkundenavn, risiko"
		        case 4 
                orderBY = "jobnavn, jobnr"
                case 5 
                orderBY = "jobnr"              
		        case 6 
                orderBY = "jobnr DESC"
                case 7,8 
                orderBY =  "risiko" '"""kkundenavn, jobnavn, jobnr"        
                case else
		        orderBY = "jobslutdato"
		        end select
		    else
		    sorter = 0
		    orderBY = "jobslutdato"
		    end if
		    
		    
		end if 
		
		
		response.cookies("webblik")("prioitet") = sorter
		
		
		prioCHK0 = ""
		prioCHK1 = ""
		prioCHK2 = ""
        prioCHK3 = ""
        prioCHK31 = ""
        prioCHK4 = ""
		prioCHK5 = ""
        prioCHK6 = ""
        prioCHK7 = ""
        'stCHK8 = ""

        'prioTxt8 = "Projektgruppe (eksl. ''Alle-gruppen'')"
        prioTxt7 = "Projektgruppe - Priroitet (drag'n drop mode)"
        prioTxt6 = "Jobnr (stigende)"
        prioTxt5 = "Jobnr (faldende)"
		prioTxt4 = "Jobnavn"
        prioTxt3 = "Kunde"
		prioTxt31 = "Kunde - Priroitet (drag'n drop mode)"
        prioTxt2 = "Periode - Startdato"
		prioTxt0 = "Periode - Slutdato" 
		prioTxt1 = "Priroitet" 
		
		
		select case cint(sorter)
		case 1
	    prioCHK1 = "SELECTED"
		vlgPrioTxt = prioTxt1
		
		case 2
		
		prioCHK2 = "SELECTED"
		vlgPrioTxt = prioTxt2
	    
        case 3
	    prioCHK3 = "SELECTED"
		vlgPrioTxt = prioTxt3

         case 31
	    prioCHK31 = "SELECTED"
		vlgPrioTxt = prioTxt31

        case 4
		prioCHK4 = "SELECTED"
		vlgPrioTxt = prioTxt4
		
        case 5
		prioCHK5 = "SELECTED"
		vlgPrioTxt = prioTxt5

        case 6
		prioCHK6 = "SELECTED"
		vlgPrioTxt = prioTxt6

         case 7
		prioCHK7 = "SELECTED"
		vlgPrioTxt = prioTxt7

           case 8
		prioCHK8 = "SELECTED"
		vlgPrioTxt = prioTxt8

		case else
		prioCHK0 = "SELECTED"
	    vlgPrioTxt = prioTxt0
		
		end select



            if len(request("FM_start_dato")) <> 0 then

            if len(request("FM_status0")) <> 0 then
	        stat0 = 1
	        stCHK0 = "CHECKED"
	        else
	        stat0 = 0
	        stCHK0 = ""
	        end if
	        
	        if len(request("FM_status1")) <> 0 then
	        stat1 = 1
	        stCHK1 = "CHECKED"
	        else
	        stat1 = 0
	        stCHK1 = ""
	        end if
	        
	        if len(request("FM_status2")) <> 0 then
	        stat2 = 1
	        stCHK2 = "CHECKED"
	        else
	        stat2 = 0
	        stCHK2 = ""
	        end if
	        

            if len(request("FM_status3")) <> 0 then
	        stat3 = 1
	        stCHK3 = "CHECKED"
	        else
	        stat3 = 0
	        stCHK3 = ""
	        end if


             if len(request("FM_status4")) <> 0 then
	        stat4 = 1
	        stCHK4 = "CHECKED"
	        else
	        stat4 = 0
	        stCHK4 = ""
	        end if

             if len(request("FM_status5")) <> 0 then
	        stat5 = 1
	        stCHK5 = "CHECKED"
	        else
	        stat5 = 0
	        stCHK5 = ""
	        end if
	   
	    else
	        

             if request.cookies("webblik")("status0") <> "" then
	                
	                stat0 = request.cookies("webblik")("status0")
	                if cint(stat0) = 1 then
	                stCHK0 = "CHECKED"
	                else
	                stCHK0 = ""
	                end if
	                
	        else
	        stat0 = 0
	        stCHK0 = ""
	        end if


	        if request.cookies("webblik")("status1") <> "" then
	                
	                stat1 = request.cookies("webblik")("status1")
	                if cint(stat1) = 1 then
	                stCHK1 = "CHECKED"
	                else
	                stCHK1 = ""
	                end if
	                
	        else
	        stat1 = 0
	        stCHK1 = ""
	        end if
	        
	        if request.cookies("webblik")("status2") <> "" then
	                
	                stat2 = request.cookies("webblik")("status2")
	                if cint(stat2) = 1 then
	                stCHK2 = "CHECKED"
	                else
	                stCHK2 = ""
	                end if
	                
	       
	        else
	        stat2 = 0
	        stCHK2 = ""
	        end if

            if request.cookies("webblik")("status3") <> "" then
	                
	                stat3 = request.cookies("webblik")("status3")
	                if cint(stat3) = 1 then
	                stCHK3 = "CHECKED"
	                else
	                stCHK3 = ""
	                end if
	                
	       
	        else
	            if lto = "epi" OR lto = "epi_no" OR lto = "epi_sta" OR lto = "epi_ab" then
                stat3 = 1
	            stCHK3 = "CHECKED"
                else
                stat3 = 0
	            stCHK3 = ""
                end if
	        end if
	        


             if request.cookies("webblik")("status4") <> "" then
	                
	                stat4 = request.cookies("webblik")("status4")
	                if cint(stat4) = 1 then
	                stCHK4 = "CHECKED"
	                else
	                stCHK4 = ""
	                end if
	                
	        else
	        stat4 = 0
	        stCHK4 = ""
	        end if

             if request.cookies("webblik")("status5") <> "" then
	                
	                stat5 = request.cookies("webblik")("status5")
	                if cint(stat5) = 1 then
	                stCHK5 = "CHECKED"
	                else
	                stCHK5 = ""
	                end if
	                
	        else
	        stat5 = 0
	        stCHK5 = ""
	        end if

	    end if
	    
        response.cookies("webblik")("status0") = stat0
	    response.cookies("webblik")("status1") = stat1
	    response.cookies("webblik")("status2") = stat2
	    response.cookies("webblik")("status3") = stat3
	    response.cookies("webblik")("status4") = stat4
        response.cookies("webblik")("status5") = stat5


        '** Skulte job **'
               if len(request("st_sl_dato")) <> 0 then
   
               if request("visskjulte") <> 0 then
               visskjulte = 1
               CHKskj = "CHECKED"
               else
               visskjulte = 0
               CHKskj = ""
               end if
    
                if request("visKunFastpris") <> 0 then
               visKunFastpris = 1
               visKunFastprisCHK = "CHECKED"
               else
               visKunFastpris = 0
               visKunFastprisCHK = ""
               end if
    

               if request("skjulnuljob") <> 0 then
               skjulnuljob = 1
               skjulnuljobCHK = "CHECKED"
               else
               skjulnuljob = 0
               skjulnuljobCHK = ""
               end if

                if len(trim(request("visKunGT"))) <> 0 then
                 visKunGT = 1
                   visKunGTCHK = "CHECKED"
                   else
                   visKunGT = 0
                   visKunGTCHK = ""
                   end if
    
                   if len(trim(request("visSimpel"))) <> 0 then
                   visSimpel = request("visSimpel")

                   visSimpelCHK0 = ""
                   visSimpelCHK1 = ""
                   visSimpelCHK2 = ""
                   visSimpelCHK3 = ""

                   select case visSimpel
                   case 0
                   visSimpelCHK0 = "SELECTED"
                   case 1
                   visSimpelCHK1 = "SELECTED"
                   case 2
                   visSimpelCHK2 = "SELECTED"
                   case else
                   visSimpelCHK1 = "SELECTED"
                   end select 

                   else
                    visSimpel = 1
                    visSimpelCHK1 = "SELECTED"
                   end if

        
      
                   if len(trim(request("visKol_jobTweet"))) <> 0 then
                   visKol_jobTweet = request("visKol_jobTweet")
                   visKol_jobTweet_CHK = "CHECKED"
                   else
                   visKol_jobTweet = 0
                   visKol_jobTweet_CHK = ""
                   end if

                   if len(trim(request("visKol_terminer"))) <> 0 then
                   visKol_terminer = request("visKol_terminer")
                   visKol_terminer_CHK = "CHECKED"
                   else
                   visKol_terminer = 0
                   visKol_terminer_CHK = ""
                   end if


                   if len(trim(request("visKol_rekvnr"))) <> 0 then
                   visKol_rekvnr = request("visKol_rekvnr")
                   visKol_rekvnr_CHK = "CHECKED"
                   else
                   visKol_rekvnr = 0
                   visKol_rekvnr_CHK = ""
                   end if


                    if len(trim(request("visKol_prioritet"))) <> 0 then
                   visKol_prioritet = request("visKol_prioritet")
                   visKol_prioritet_CHK = "CHECKED"
                   else
                   visKol_prioritet = 0
                   visKol_prioritet_CHK = ""
                   end if

                    

                   if len(trim(request("visKol_forecast"))) <> 0 then
                   visKol_forecast = request("visKol_forecast")
                   visKol_forecast_CHK = "CHECKED"
                   else
                   visKol_forecast = 0
                   visKol_forecast_CHK = ""
                   end if
                   

                   if len(trim(request("visKol_planlagning"))) <> 0 then
                   visKol_planlagning = request("visKol_planlagning")
                   visKol_planlagning_CHK = "CHECKED"
                   else
                   visKol_planlagning = 0
                   visKol_planlagning_CHK = ""
                   end if

                    


               else
   
   
                    if request.cookies("webblik")("visskj") <> "1" then
                    visskjulte = 0
                    CHKskj = ""
                    else
                    visskjulte = 1
                    CHKskj = "CHECKED"
                    end if

                   if request.cookies("webblik")("viskunfs") <> "1" then
                   visKunFastpris = 0
                   visKunFastprisCHK = ""
                   else
                   visKunFastpris = 1
                   visKunFastprisCHK = "CHECKED"
                   end if

      
                   skjulnuljob = 0
                   skjulnuljobCHK = ""
     

                     if request.cookies("webblik")("viskungt") <> "1" then
                   visKunGT = 0
                   visKunGTCHK = ""
                   else
                   visKunGT = 1
                   visKunGTCHK = "CHECKED"
                   end if



                   if request.cookies("webblik")("visSimpel") <> "" then
       
                   visSimpel = request.cookies("webblik")("visSimpel") 

                   select case visSimpel
                   case 0
                   visSimpelCHK0 = "SELECTED"
                   case 1
                   visSimpelCHK1 = "SELECTED"
                   case 2
                   visSimpelCHK2 = "SELECTED"
                   case 3
                   visSimpelCHK3 = "SELECTED"
                   case else
                   visSimpelCHK1 = "SELECTED"
                   end select 
       
                   else

                       select case lto
                       case "epi", "epi_no, epi_bar", "epi_as", "epi2017", "epi_2017"
                       visSimpel = 2 'udv
                       visSimpelCHK2 = "SELECTED"
                       case else
                       visSimpel = 1 'Standard
                       visSimpelCHK1 = "SELECTED"
                       end select

                   end if


                    '**Job Tweet kolonne **'

                   if request.cookies("webblik")("visKol_jobTweet") <> "" then

                       visKol_jobTweet = request.cookies("webblik")("visKol_jobTweet")

                       if cint(visKol_jobTweet) = 1 then
                       visKol_jobTweet_CHK = "CHECKED"
                       else
                       visKol_jobTweet_CHK = ""
                       end if

                   else

                       select case lto
                        case "epi", "epi_no, epi_bar", "epi_as", "epi2017", "epi_2017"
                       visKol_jobTweet_CHK = ""
                       case else
                        visKol_jobTweet_CHK = "CHECKED"
                       end select

                   end if



                    '***** Terminer

                     if request.cookies("webblik")("visKol_terminer") <> "" then

                       visKol_terminer = request.cookies("webblik")("visKol_terminer")

                       if cint(visKol_terminer) = 1 then
                       visKol_terminer_CHK = "CHECKED"
                       else
                       visKol_terminer_CHK = ""
                       end if

                   else

                      
                       visKol_terminer_CHK = ""
                       
                   end if


                   '***** Vis Rekvnr
        
                    if request.cookies("webblik")("visKol_rekvnr") <> "" then

                       visKol_rekvnr = request.cookies("webblik")("visKol_rekvnr")

                       if cint(visKol_rekvnr) = 1 then
                       visKol_rekvnr_CHK = "CHECKED"
                       else
                       visKol_rekvnr_CHK = ""
                       end if

                   else

                      
                       visKol_rekvnr_CHK = ""
                       
                   end if

                 

                   '**** Vise forecast

                     if request.cookies("webblik")("visKol_forecast") <> "" then

                       visKol_forecast = request.cookies("webblik")("visKol_forecast")

                       if cint(visKol_forecast) = 1 then
                       visKol_forecast_CHK = "CHECKED"
                       else
                       visKol_forecast_CHK = ""
                       end if

                   else

                      
                       visKol_forecast_CHK = ""
                       
                   end if


                   
                   '**** Vise planlagning

                     if request.cookies("webblik")("visKol_planlagning") <> "" then

                       visKol_planlagning = request.cookies("webblik")("visKol_planlagning")

                       if cint(visKol_planlagning) = 1 then
                       visKol_planlagning_CHK = "CHECKED"
                       else
                       visKol_planlagning_CHK = ""
                       end if

                   else

                      
                       visKol_planlagning_CHK = ""
                       
                   end if

                    


        
               end if
   

               response.cookies("webblik")("visKol_rekvnr") = visKol_rekvnr
               response.cookies("webblik")("visKol_forecast") = visKol_forecast
               response.cookies("webblik")("visKol_planlagning") = visKol_planlagning
               response.cookies("webblik")("visKol_terminer") = visKol_terminer
               response.cookies("webblik")("visKol_jobTweet") = visKol_jobTweet
               response.cookies("webblik")("visSimpel") = visSimpel
               response.cookies("webblik")("viskungt") = visKunGT
               response.cookies("webblik")("skjulnuljob") = skjulnuljob
               response.cookies("webblik")("visskj") = visskjulte
               response.cookies("webblik")("viskunfs") = visKunFastpris



       
        select case lto 
        case "epi", "epi_no", "epi_sta", "intranet - local", "epi_ab"
        ltoLimit = "0,5000" '5000
        ltoLimitTXT = 2500
        case else
        ltoLimit = "0,5000" '1000
        ltoLimitTXT = 1000
        end select
	
	  
		if len(request("FM_kunde")) <> 0 then 'Er der submitted
		    
		    if len(trim(request("jobnr_sog"))) <> 0 then
			jobnr_sog = trim(request("jobnr_sog"))
            show_jobnr_sog = jobnr_sog
            else
            jobnr_sog = "-1"
			jobKri = ""
			show_jobnr_sog = ""
            end if

         else
                if request.cookies("webblik")("jobnrnavn") <> "" then
			    jobnr_sog = request.cookies("webblik")("jobnrnavn")
                show_jobnr_sog = jobnr_sog
                else
                jobnr_sog = "-1"
			    jobKri = ""
			    show_jobnr_sog = ""
                end if
                 
         end if
			    

            if jobnr_sog <> "-1" then '** ER DER S�GT eller er feltet tomt? 
                
                if instr(jobnr_sog, ",") <> 0 then '** Komma **'

	                jobKri = " j.jobnr = 0 "
	                jobSogValuse = split(jobnr_sog, ",")
	                for i = 0 TO UBOUND(jobSogValuse)
	                    jobKri = jobKri & " OR j.jobnr = '"& trim(jobSogValuse(i)) &"'"   
	                next
	    
	                jobKri = " AND ("&jobKri&")"

                else

                    if instr(jobnr_sog, "--") <> 0 then '** Interval
	                jobSogValuse = split(jobnr_sog, "--")
	                    
                        call erDetInt(jobSogValuse(0))
                        isInt = isInt
                        call erDetInt(jobSogValuse(1))
                        isInt = isInt  
                                    
                        if isInt = 0 then 'numerisk
                        jobKri = " AND (j.jobnr >= "& trim(jobSogValuse(0)) &" AND j.jobnr <= " & trim(jobSogValuse(1)) & ")"
                        else
                        jobKri = " AND (j.jobnr >= '"& trim(jobSogValuse(0)) &"' AND j.jobnr <= '" & trim(jobSogValuse(1)) & "')"  
                        end if   
                    else
                    
                    if instr(jobnr_sog, ">") <> 0 OR instr(jobnr_sog, "<")  <> 0 then '** Afgr�nsning
	                
                            
                            if instr(jobnr_sog, ">") <> 0 then
                            jobnr_sog = replace(jobnr_sog, ">", "")
                            
                            call erDetInt(jobnr_sog)
                           
                                    
                            if isInt = 0 then 'numerisk
                            jobKri = " AND (j.jobnr > "& trim(jobnr_sog) &")"  
                            else 
                            jobKri = " AND (j.jobnr > '"& trim(jobnr_sog) &"')"  
                            end if 
                            
                            else
                            
                            jobnr_sog = replace(jobnr_sog, "<", "")

                            call erDetInt(jobnr_sog)
                           
                                    
                            if isInt = 0 then 'numerisk
                            jobKri = " AND (j.jobnr < "& trim(jobnr_sog) &")"   
                            else 
                            jobKri = " AND (j.jobnr < '"& trim(jobnr_sog) &"')" 
                            end if
                        
                            end if
	               
                    else
                    jobKri = "AND (j.jobnr LIKE '"& jobnr_sog &"%' OR j.jobnavn LIKE '"& jobnr_sog &"%' OR project_tier = '"& jobnr_sog &"') "
                    end if
                    end if

			    end if


             end if

       jobnrnavn = show_jobnr_sog 
	  response.cookies("webblik")("jobnrnavn") = jobnrnavn
	
	
	
    if len(request("FM_kunde")) <> 0 then
			
			if request("FM_kunde") = 0 then
			valgtKunde = 0
			sqlKundeKri = "jobknr <> 0"
			else
			valgtKunde = request("FM_kunde")
			sqlKundeKri = "jobknr = "& valgtKunde &""
			end if
			
	else
		
			if len(request.cookies("webblik")("kon")) <> 0 AND request.cookies("webblik")("kon") <> 0 then
			valgtKunde = request.cookies("webblik")("kon")
			sqlKundeKri = "jobknr = "& valgtKunde &""
			else
			valgtKunde = 0
			sqlKundeKri = "jobknr <> 0"
			end if
			
	end if
	
	
	response.Cookies("webblik")("kon") = valgtKunde
    














    	
	'*******************************************************
	'**** Igangv�rende Job *********************************
	'*******************************************************

        select case visSimpel
        case 1
        containerWdtpx = 1200
        case 2
        containerWdtpx = 1400
        case else
        containerWdtpx = 0
        end select

        if cint(visKol_terminer) = 1 then
        containerWdtpx = containerWdtpx + 100
        end if

        if cint(visKol_jobTweet) = 1 then
        containerWdtpx = containerWdtpx + 100
        end if

        if cint(visKol_rekvnr) = 1 then
        containerWdtpx = containerWdtpx + 100
        end if

        if cint(visKol_prioritet) = 1 then
        containerWdtpx = containerWdtpx + 100
        end if

        if cint(visKol_forecast) = 1 then
        containerWdtpx = containerWdtpx + 50
        end if

        if cint(visKol_planlagning) = 1 then
        containerWdtpx = containerWdtpx + 50
        end if

        
        
        


        select case visSimpel
        case 1, 2
        containerWdt = "style=width:"& containerWdtpx &"px"
        case else
        containerWdt = ""
        end select

       


        if len(trim(request("FM_medarb_jobans"))) <> 0 then
        medarb_jobans = request("FM_medarb_jobans")
        else
                if request.cookies("webblik")("jk_ans") <> "" then
                medarb_jobans = request.cookies("webblik")("jk_ans")
                else
                    if lto = "dencker" OR lto = "hidalgo" then
                    medarb_jobans = session("mid")
                    else
                    medarb_jobans = 0
                    end if
                end if
        end if

        response.cookies("webblik")("jk_ans") = medarb_jobans
       
        

        
        if len(trim(request("jobansv"))) <> 0 AND request("jobansv") <> 0 then
            jobansv = 1
            jobansvCHK = "CHECKED"
        else
            if request.cookies("webblik")("jobansv") = "1" AND len(request("FM_kunde")) = 0 then
            jobansv = 1
            jobansvCHK = "CHECKED"
            else
            jobansv = 0
            jobansvCHK = ""
            end if
        end if

        if len(trim(request("salgsansv"))) <> 0 AND request("salgsansv") <> 0 then
            salgsansv = 1
            salgsansvCHK = "CHECKED"
        else
            if request.cookies("webblik")("salgsansv") = "1" AND len(request("FM_kunde")) = 0 then
            salgsansv = 1
            salgsansvCHK = "CHECKED"
            else
            salgsansv = 0
            salgsansvCHK = ""
            end if
        end if


         if len(trim(request("kansv"))) <> 0 AND request("kansv") <> 0 then
            kansv = 1
            kansvCHK = "CHECKED"
        else
            if request.cookies("webblik")("kansv") = "1" AND len(request("FM_kunde")) = 0 then
            kansv = 1
            kansvCHK = "CHECKED"
            else 
            kansv = 0
            kansvCHK = ""
            end if
        end if
            

            response.cookies("webblik")("jobansv") = jobansv
            response.cookies("webblik")("salgsansv") = salgsansv
            response.cookies("webblik")("kansv") = kansv

       
      
    
    '**** Projektgrupper ******    
    if level <= 2 OR level = 6 then
        
                if len(request("FM_start_dato")) <> 0 then
        
                    if len(trim(request("FM_ignorer_pg"))) <> 0 OR request("FM_sorter") = 7 OR request("FM_sorter") = 8 then
	                IgnrProjGrp = 1
                    strPgrpSQLkri = ""
	                else
	                IgnrProjGrp = 0
		            call hentbgrppamedarb(session("mid"))
		            end if
	             
	             
	             else  
	                
	                
	                
	                    if request.cookies("webblik")("ignorer_pg") <> "" then
	                    IgnrProjGrp = request.cookies("webblik")("ignorer_pg")
	                        if cint(IgnrProjGrp) = 1 then 
    	                        strPgrpSQLkri = ""
	                        else
    	                        call hentbgrppamedarb(session("mid"))
	                        end if
                	    
	                    else
	                    IgnrProjGrp = 0
		                call hentbgrppamedarb(session("mid"))
		                end if
	               
	            end if
                	
	                
        
		if cint(IgnrProjGrp) = 1 then
		chkThis = "CHECKED"
		else
		chkThis = ""
		end if

    
    else
    IgnrProjGrp = 0
    call hentbgrppamedarb(session("mid"))
    end if
    
    response.cookies("webblik")("ignorer_pg") = IgnrProjGrp


    %>

      
	
<div class="wrapper">
      <div class="content">
          <div class="container" style="width:90%;"> <!-- <%=containerWdt %> -->
              <div class="portlet">
                  <h3 class="portlet-title"><u>Igangv�rende Job</u></h3>
                  <div class="portlet-body">


	


	<!--<div id="joblisten" style="position:absolute; left:<%=dleft%>px; top:<%=dtop%>px; visibility:visible; display:;">-->
	
	<%if print <> "j" then %>
	
	<div class="well">
	<table width=100% cellpadding=0 cellspacing=0 border=0>
	<form method="post" id="joblist_filter" action="webblik_joblisten.asp?FM_usedatokri=1&nomenu=<%=nomenu %>">
        <!--<input type="text" id="test">-->
	<tr>
	<td valign=top style="width:500px; padding-right:40px;">



    <%'**** Kunder ****' %>

    <b>Kunder:</b>
    
    <%if print <> "j" then %>
    <br> <select name="FM_kunde" size="1" class="form-control input-small">
		<option value="0">Alle</option>
		<%
		end if
		
		ketypeKri = " ketype <> 'e'"
		strKundeKri = " AND kid <> 0 "
		vlgtKunde = " Alle "
		
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE "& ketypeKri &" "& strKundeKri &" ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(valgtKunde) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				vlgtKunde = oRec("Kkundenavn") & " ("& oRec("Kkundenr") &")"
				else
				isSelected = ""
				end if
				
				if print <> "j" then%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%> (<%=oRec("Kkundenr") %>)</option>
				<%end if
				oRec.movenext
				wend
				oRec.close
				
				if print <> "j" then%>
		</select>
		<%end if %>
		
		
    <%if print = "j" then %>
    &nbsp;<%=vlgtKunde %><br />
    <%end if %>
    
    
    
    
   
    
    
     
    
    <br />
    <b>Job / kundeansvarlig</b> vis kun job hvor valgte medarb. er: <br />
           
           <%if print = "j" then%>
           - 
           <%end if %>
           
       
        
            <%if print <> "j" then%>
            <select name="FM_medarb_jobans" id="FM_medarb_jobans" class="form-control input-small">
            <option value="0">Alle - ignorer jobansv.</option>
            <%end if
            
            mNavn = "Alle (job / kunde ansv. ignoreret)"
            
             strSQL = "SELECT mnavn, mnr, mid, init FROM medarbejdere WHERE mansat = 1 ORDER BY mnavn"
             oRec.open strSQL, oConn, 3
             while not oRec.EOF
                
                if cint(medarb_jobans) = oRec("mid") then
                selThis = "SELECTED"
                        
                        '** medarb navn til print **
                        mNavn = oRec("mnavn") & " ("& oRec("mnr") &") "
                       
                        if len(trim(oRec("init"))) <> 0 then 
                        mNavn = mNavn & " - "& oRec("init") 
                        end if 
                        
                else
                selThis = ""
                end if
                
                if print <> "j" then%>
                <option value="<%=oRec("mid")%>" <%=selThis%>><%=oRec("mnavn") %> (<%=oRec("mnr") %>)
                 <%if len(trim(oRec("init"))) <> 0 then %>
                  - <%=oRec("init") %>
                  <%end if %>
                  </option>
                 <%
                end if
             
             oRec.movenext
             wend
             oRec.close
             %>
             
         <%if print <> "j" then%>   
        </select>
        <input name="jobansv" id="job_kans" value="1" type="checkbox" <%=jobansvCHK%> /> Jobansv. (1-5) &nbsp;
        <input name="salgsansv" id="salgs_kans" value="1" type="checkbox" <%=salgsansvCHK%> /> Salgsansv. (1-5) &nbsp; 
        <input name="kansv" id="Radio1" value="1" type="checkbox" <%=kansvCHK%> /> Kundeansvarlig 
        
        <%else %>
     
       
            
            <%=mNavn %> er
            <%if cint(jobansv) = 1 then %>
            <br />Jobansvarlig
            <%end if %>

             <%if cint(salgsansv) = 1 then %>
            <br />Salgsansvarlig
            <%end if %>

            <%if cint(kansv) = 1 then %>
            <br />Kundeansvarlig
            <%end if %>
        
        <%end if 
        
      
        
        %>
    
  
    
      
    
         <% if print <> "j" then %>
          <br /><br /><div class="well well-white">
         <b>S�g p� jobnavn ell. jobnr.:<br /><span style="font-size:9px; font-weight:lighter; line-height:14px;">(% wildcard, <b>231, 269</b> specifik, <b>201--225</b> (dobbelt bindestreg) interval, eller <b>< ></b> - MAX 5000 job)</span></b>
         <input type="text" name="jobnr_sog" id="jobnr_sog" value="<%=jobnrnavn%>" class="form-control input-small" style="width:400px;">
		    <%
		
		    else %>
		    <b>jobnr./jobnavn:</b><br />
		    <%=jobnrnavn%>
		    <%end if %>
              </div>  


       
	    
	    
	<br />
	<span id="sp_avanceret_1" style="color:#5582d2;">[+] Avanceret & Pr�sentation</span>
    <div class="dv_avanceret_1" style="visibility:hidden; display:none;">
    <br /><b>Realiserede timer og faktureret bel�b:</b>	

           <%select case lto 
        case "epi", "epi_no", "epi_ab", "epi_sta", "intranet - local"
        %>
        <br /><span style="color:#999999;">Timer realiseret opgjort t.o.m ig�r.</span>
        <%end select %>
        <br /><br />

      <span>
	
	<% if print <> "j" then %>
        <input id="Radio2" name="realfakpertot" type="radio" value="0" <%=realfakpertot0CHK %> /> Vis <b>total</b> p� viste job uanset valgt periode
           
        
        <br />
        <input id="Radio3" name="realfakpertot" type="radio" value="1" <%=realfakpertot1CHK %> /> Vis kun timer og bel�b fra <b>den valgte periode</b>

          <%select case lto 
        case "epi", "epi_no", "epi_ab", "epi_sta", "intranet - local"
        %>
        (fulde m�neder)
        <%end select %>
            </span>


          <br />
        <input id="Checkbox4" name="timertilfak" type="checkbox" value="1" <%=timertilfakCHK %> /> Vis igangv�rende arb. (timer nyere en seneste faktura i valgte periode)
        

        <br />
        <input id="Radio6" name="historisk_wip" type="checkbox" value="1" <%=historisk_wipCHK %> /> Brug historisk WIP (henter g�ldende WIP i forhold til valgte datointerval)
        
     

       
	
	<%else %>
	- <%=realfakpertotTxt%><br />
    - <%=timertilfakTxt%><br />
    - <%=historisk_wipTxt%>

	<%end if%>
		<br /> &nbsp; 
	
	</div>
	</td>
	
	<td valign=top style="padding-left:20px;">

                <b>Periode fra - til:</b><br />

        <%if print <> "j" then %>
                <div class="row">
                    
                            <div class="col-lg-2">
                                <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="FM_start_dato" value="<%=strDag&"-"&strMrd&"-"&strAar%>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>

                            
                            <div class="col-lg-2">
                                <div class='input-group date' id='datepicker_stdato'>
                                <input type="text" class="form-control input-small" name="FM_slut_dato" value="<%=strDag_slut&"-"&strMrd_slut&"-"&strAar_slut%>" placeholder="dd-mm-yyyy" />
                                <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                                </div>
                            </div>

                    </div>

	
	    <%end if %>


         <div class="row">
              <div class="col-lg-8">
		
		<%if print = "j" then%>
		<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%><br /><br />
	    <%end if %>
	    
		
		

        
        Viser kun job der opfylder f�lgende kriterie:
    
    <%if print <> "j" then %>

        <select name="st_sl_dato" class="form-control input-small" style="width:400px;">
            <option value="0" <%=st_sl_DatoChk0 %>>Startdato</option>

            <option value="1" <%=st_sl_DatoChk1 %>>Startdato</option>
            <option value="3" <%=st_sl_DatoChk3 %>>Slutdato</option>
            <option value="6" <%=st_sl_DatoChk6 %>>Timer, fakturaer ell. materiale-forbrug</option>
            <option value="5" <%=st_sl_DatoChk5 %>>Fakturaer i valgte interval</option>
            <option value="4" <%=st_sl_DatoChk4 %>>Der var aktivt: alle job + lukkede job hvor Lukkedato > Slutdato i periode. (igangv�rende arb.)</option>
            <option value="2" <%=st_sl_DatoChk2 %>>Ignorer datointerval (vis alle job, maks. <%=ltoLimitTxt %>)</option>

        </select>

  
   <%else 
	    
	   
	    Response.Write printVal & " i forhold til datointerval."
	   
	    
	end if 
        
       %>
    </div>
    </div>

        
          <div class="row">
              <div class="col-lg-6">
        <br /><br />
		<b>Jobstatus:</b><br />
	
		
		
		
		
		<%
        if print <> "j" then%>
        <input type="checkbox" name="FM_status0" value="0" <%=stCHK0%>>&nbsp;Aktive&nbsp;
	    <input type="checkbox" name="FM_status1" value="1" <%=stCHK1%>>&nbsp;Passiv / Til fak.&nbsp;
	     <input type="checkbox" name="FM_status2" value="2" <%=stCHK2%>>&nbsp;Lukkede&nbsp;<br />
         <input type="checkbox" name="FM_status3" value="3" <%=stCHK3%>>&nbsp;Tilbud&nbsp;
         <input type="checkbox" name="FM_status4" value="4" <%=stCHK4%>>&nbsp;Gennemsyn
        <input type="CHECKBOX" name="FM_status5" value="5" <%=stCHK5%>/>&nbsp;Eval.<br />
	    <%else %>

               <%if stat0 = "1" then%>
	        - Aktive<br />
	        <%end if %>
	    
	        <%if stat1 = "1" then%>
	        - Passiv / Til fak.<br />
	        <%end if %>
	        
	         <%if stat2 = "1" then%>
	        - Lukkede<br />
	        <%end if %>

             <%if stat3 = "1" then%>
	        - Tilbud<br />
	        <%end if %>

             <%if stat4 = "1" then%>
	        - Gennemsyn<br />
	        <%end if %>

            <%if stat5 = "1" then%>
	        - Evaluering<br />
	        <%end if %>
	    
	    
	    <%end if %>

         </div>
        </div>



        <div class="dv_avanceret_1" style="visibility:hidden; display:none; padding-top:60px;"> 
		<b>Filter:</b>
   
  
    
    <%if print <> "j" then %>
       <br /> <input id="visskjulte" name="visskjulte" type="checkbox" <%=CHKskj %> value="1" /> Vis ogs� Interne job (prioitet < 0)
    <%else 
        if visskjulte = 1 then%>
        - Viser interne job 
        <%else %>
        - Viser ikke interne job 
        <%end if %>
    <%end if %>
	    
	    
		
		

     <%if print <> "j" then %>
        <br /><input id="Checkbox2" name="visKunFastpris" type="checkbox" <%=visKunFastprisCHK %> value="1" /> Vis kun Fastpris job
    <%else 
        if visKunFastpris = 1 then%>
        <br />- Viser kun fastpris job
        <%else %>
        <br />- Viser b�de fastpris og lnb. timer job 
        <%end if %>
    <%end if %>

   

      <%if print <> "j" then %>
        <br /><input id="Checkbox3" name="skjulnuljob" type="checkbox" <%=skjulnuljobCHK %> value="1" /> Skjul "nul" job (job uden timer i per.)
    <%else 
        if skjulnuljob = 1 then%>
        <br />- Skjuler "nul" job 
        <%else %>
        <br />- Viser "nul" job 
        <%end if %>
    <%end if %> 

    
    
		
		
		<%
		if print <> "j" then%>
		<br /><input type="checkbox" name="FM_ignorer_pg" id="FM_ignorer_pg" value="1" <%=chkThis%>> Ignorer tilknytning til job via dine projektgrupper.
		<%else 
		    
		    if cint(IgnrProjGrp) = 1 then%>
		    - Ignoreret (vis alle).
		    <%else %>
		    - Viser kun de job du er tilknyttet via dine projektgrupper.
		    <%end if %>
		<%end if %>
    
   
    
    <br />  <br />

   


    <br /><br /><b>Kolonne layout</b><br />
    <%if print <> "j" then %>


            <select name="visSimpel" class="form-control input-small" style="width:300px;">
                <option value="0" <%=visSimpelCHK0 %>>Simpel</option>
                <option value="1" <%=visSimpelCHK1 %>>Standard</option>
                <option value="2" <%=visSimpelCHK2 %>>Udviddet</option>
                <!--<option value="3" <%=visSimpelCHK3 %>>N�gletal �TD</option>-->

            </select>

            <!--

        <br /><input id="Checkbox3" name="visSimpel" type="radio" <%=visSimpelCHK0 %> value="0" /> Vis simpel 
        <br /><input id="Radio4" name="visSimpel" type="radio" <%=visSimpelCHK1 %> value="1" /> Vis standard (job.besk., wip og forecast) 
           <%
                 
                 if instr(lto, "xepi") OR lto = "xintranet - local" then%>
                  <span style="color:red;">Konsoliderede tal</span>
                 <%end if %>
           
         <br /><input id="Radio5" name="visSimpel" type="radio" <%=visSimpelCHK2 %> value="2" /> Vis udviddet (som standard + �konomi) 
        
        <%
                 if instr(lto, "xepi") OR lto = "xintranet - local" then%>
                  <span style="color:red;">Konsoliderede tal</span>
                 <%end if %>
          
         -->
          <br /><br />
          <input type="checkbox" value="1" name="visKol_jobTweet" <%=visKol_jobTweet_CHK %> /> Vis Job Tweet kolonne

          <br />
          <input type="checkbox" value="1" name="visKol_terminer" <%=visKol_terminer_CHK %> /> Vis Terminer (kun udviddet)

          <br />
          <input type="checkbox" value="1" name="visKol_rekvnr" <%=visKol_rekvnr_CHK %> /> Vis Rekvisitions nr.

        <br />
          <input type="checkbox" value="1" name="visKol_prioritet" <%=visKol_prioritet_CHK %> /> Vis Prioritet

          <br />
          <input type="checkbox" value="1" name="visKol_forecast" <%=visKol_forecast_CHK %> /> Vis Forecast

        <br />
          <input type="checkbox" value="1" name="visKol_planlagning" <%=visKol_planlagning_CHK %> /> Vis Planl�gning
      
     <%else 

        select case visSimpel
          case 0%>
        <br />- Simpel visning
        <%case 1 %>
        <br />- Standard visning
        <%case 2 %>
        <br />- Udviddet visning
        <%end select %>

      
    <%end if %> 

     
           <%if print <> "j" then %>
        <br /><br />
         <b>Eksporter direkte</b>   
         <input id="Checkbox1" name="visKunGt" type="checkbox" <%=visKunGTCHK %> value="1" /> Vis kun Grandtotal / forbered til csv. eksport (hurtig loadtid)<br /><span style="color:#999999;">
        (grandtotal vises ikke hvis der sorteres efter projektgrupper)</span>
    <%else 
        if visKunGT = 1 then%>
        <br />- Viser kun grandtotal 
        <%end if %>   


    <%end if %> 


            </div>
            
    </td>
    </tr>
    <tr>
        <td valign=bottom><br /><br />


            <%


  
		

        if cint(sorter) = 7 then
        dv_prg_dsp = ""
        dv_prg_vzb = "visible"
		else
         dv_prg_dsp = "none"
        dv_prg_vzb = "hidden"
        end if

		%>
	    
	    
	   
		<b>Sort�r efter:</b><br />
		
		<%if print <> "j" then %>
		
		<select name="FM_sorter" id="FM_sorter" class="form-control input-small" style="width:300px;" onchange="submit();">
		<option value="1" <%=prioCHK1%>>Priroitet (drag'n drop mode)</option>
		<option value="0" <%=prioCHK0%>>Slutdato</option>
		<option value="2" <%=prioCHK2%>>Startdato</option>
        <option value="3" <%=prioCHK3%>>Kunde</option>
        <option value="31" <%=prioCHK31%>>Kunde - Priroitet (drag'n drop mode)</option>
        <option value="4" <%=prioCHK4%>>Jobnavn</option>
		<option value="5" <%=prioCHK5%>>Jobnr (stigende)</option>
        <option value="6" <%=prioCHK6%>>Jobnr (faldende)</option>

            <%select case lto 
               case "epi", "epi_ab", "epi_sta", "epi_no"
                
                case else%>
        <option value="7" <%=prioCHK7%>>Projektgruppe - Priroitet (drag'n drop mode)</option>
              <!--<option value="8" <%=stCHK8%>>Projektgruppe (eksl. "Alle-gruppen")</option>-->
            <%end select %>
        
        </select>
         


        <br /><br /><div id="dv_projektgrp" style="position:relative; width:350px; left:20px; background-color:#F7F7F7; padding:20px; display:<%=dv_prg_dsp%>; visibility:<%=dv_prg_vzb%>;"><b>V�lg projektgrupper:</b> (og klik "vis joblisten >> ")<br />
              
            
            <input type="checkbox" id="progrp_all" /> V�lg Alle / Ingen <br /><br />
                
                <%strSQLprogrpfilt = "SELECT id, navn FROM projektgrupper WHERE id > 1 ORDER BY navn" 
                    
                    oRec.open strSQLprogrpfilt, oConn, 3
                    while not oRec.EOF 
                        if instr(sorter_visprogrp, "#"&oRec("id")&"#") <> 0 then
                        chkThis = "CHECKED"
                        else
                        chkThis = ""
                        end if%>
                    <input type="checkbox" value="#<%=oRec("id") %>#" id="FM_sorter_visprogrp_<%=oRec("id") %>" name="FM_sorter_visprogrp" class="FM_sorter_visprogrp" <%=chkThis %> /> <%=oRec("navn") %><br />
                    <%
                    oRec.movenext
                    wend 
                    oRec.close%>

              </div>

		
		<%else %>
		
		- <%=vlgPrioTxt %>
		
		<%end if
            %>

        </td>
	<td align="right">
        
	<%if print <> "j" then%>
     <br /><br />
	<input type="submit" class="btn btn-success" value="Vis joblisten >>">

       

	<%end if%>
	</td></tr>
	</form>
	</table>
	
	<!-- filter div -->
	</td></tr></table>    
  
    </div><!-- well -->


   <%end if 'print
       
     response.cookies("webblik").expires = date + 60   
       

        if cint(fromvmenu) = 1 then
       
                'Stop load
                Response.Write "Page stopped loading.<br>Please re-submit search criteria."
                Response.end  

        end if

       
    if cint(fromvmenu) <> 1 then 
     
    %>
    <br />
	<br />
	<%if sorter = "1" OR sorter = "31" OR sorter = "7" OR sorter = "8" then
	tbID = "incidentlist"


    'G�r det muligt at stoppe load
    response.flush

	else
	tbID = ""

    

	end if 
        
        
     %>

   

    
	<table class="table table-striped" id="<%=tbID%>" style="width:100%;">
	<form method="post" action="webblik_joblisten.asp?func=opdaterjobliste">
    <input type="hidden" id="FM_session_user" name="FM_session_user" value="<%=session("user")%>">
    <input type="hidden" id="FM_now" name="FM_now" value="<%=formatdatetime(now, 2)%>">
    
    <!--
    <input id="st_sl_dato_forventet" name="st_sl_dato_forventet" type="hidden" value="<%=st_sl_dato_forventet%>"/>
	-->
	
	<% 
	sqlDatostart = strAar&"/"&strMrd&"/"&strDag  'year(datointervalstart)&"/"& month(datointervalstart)&"/"&day(datointervalstart) 
	sqlDatoslut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut 'year(datointervalslut)&"/"& month(datointervalslut) &"/"& day(datointervalslut)
	
    sqlDatostartMtypTkons = strAar &"/"& strMrd & "/1" 
           

    startDatovlgt = strDag&"/"&strMrd&"/"&strAar
    slutDatovlgt = strDag_slut&"/"&strMrd_slut&"/"&strAar_slut
    startDatovlgt4md = dateAdd("m", -3, slutDatovlgt)

     'mtyp On
    if visSimpel = 2 then
        call bdgmtypon_fn()
    
        vlgtmtypgrp = 0
        call mtyperIGrp_fn(vlgtmtypgrp,1) 
        call fn_medarbtyper()

    end if
        
        
                select case visSimpel 
                case 0
                dtColSPan = 6 '10
                case 1
                dtColSPan = 7 '8 '14
                case 2
                dtColSPan = 10
                end select       

                if cint(visKol_forecast) = 1 then
                dtColSPan = dtColSPan + 1
                end if

                if cint(visKol_planlagning) = 1 then
                dtColSPan = dtColSPan + 1
                end if
                

                if cint(visKol_rekvnr) = 1 then
                dtColSPan = dtColSPan + 1
                end if

                if cint(visKol_prioritet) = 1 then
                dtColSPan = dtColSPan + 1
                end if

                

                if cint(visKol_jobTweet) = 1 then
                dtColSPan = dtColSPan + 1
                end if

                if cint(visKol_terminer) = 1 then
                dtColSPan = dtColSPan + 1
                end if

                if sorter = 3 OR sorter = 31 then
                dtColSPan = dtColSPan - 1
                end if

                


                if lto = "dencker" OR lto = "intranet - local" OR lto = "hidalgo" then 
                dtColSpan = dtColspan + 1
                end if
        
     %>
	
    <%if print <> "j" then %>
	<!--<tr>

	<td colspan="<%=dtColSPan %>" align=right style="padding:20px 10px 10px 0px;"><input type="submit" value="Opdater liste >>"></td>
    </tr>
        --><br /><br />
	<%end if %>
	
	<%call akttyper2009(2) %>
	
	<thead>
        <tr>

            <%select case visSimpel 
        case 0,1
                if sorter <> 3 AND sorter <> 31 then%>
                <th >Kontakt</th>
                <%end if %>
		<th >Job</th>
        <!--<td >Job. ansv. & ejer</td>-->

        <%
        case 2%>
		<th style="width:300px;">
                <%
                if sorter <> 3 AND sorter <> 31 then%>
                 Kontakt<br />
                <%end if %>
           
		Job<br />
		Jobbesk.<br />
		Job. ansv. & ejer<br />
        </th>
        <%end select %>


         <%if cint(visKol_rekvnr) = 1 then%>
        <th >Rekvitisions nr.<br />Pre. konditioner m�dt.</th>
        <%end if %>
		
        

        <%select case visSimpel 
        case 0
              %>
        <th style="width:250px;">Periode</th>
         <th>Status</th>

         <% if cint(visKol_prioritet) = 1 then %> 
         <th>Prioritet</><br />(-1 = skjul)</th>
        <%end if %>
		
        <%
        case 1
        %>
      
        <th>Status</th>

         <% if cint(visKol_prioritet) = 1 then %> 
         <th>Prioritet</><br />(-1 = skjul)</th>
        <%end if %>
		
        <%
        case 2
        %>
        <th style="width:250px;">Periode<br />Status</th>

         <% if cint(visKol_prioritet) = 1 then %> 
        <th >Prioritet<br />(-1 = skjul)</th>
        <%end if

        end select%>
		

          <!--<%if lto = "dencker" OR lto = "intranet - local" then %>
          <th>Ill.</th>
          <%end if%>
            -->

		<%if cint(visKol_forecast) = 1 then%>
        <th>Forecast</th>
        <%end if %>
        
     


        <th >
		<b>Realiseret</b>

             <%select case visSimpel 
            case 0
            case 1,2
                 
                 if instr(lto, "xepi") OR lto = "xintranet - local" then%>
                  <span style="background-color:#CCCCCC; color:red; display:block; padding:2px;">Konsoliderede tal<br />(se konsolidering)</span>
                 <%end if %>
            <%end select %>

        </th>
		
        
       

         <%select case visSimpel 
        case 0
        case 1,2%>
		
        
        <th >Stade</th>
       <th>&nbsp;</th>
         <%end select %>

        <%if visSimpel = 2 OR (lto <> "dencker" AND lto <> "hidalgo") then %>
        <th style="width:200px;">Forkalk. Budget</th>
        <%end if %>



          <%select case visSimpel 
        case 0, 1
              %>
     
        <%
        case 2%>
        <th>Realiseret</th>
		<th>WIP - (Igangv�rende arb.<br /> baseret p� restestimat)</th>
        <th>Faktisk faktureret</th>


         <%if cint(visKol_terminer) = 1 then %>
        <th>Betalingsplan<br /> terminer 
             <%if jobasnvigv = 1 then %>
              (stadeindm.)
             <%end if %>
            </th>
        <%end if%>

		<!--<th>Balance <%=basisValISO_f8 %></th>-->


        <%end select %>

        <%if cint(visKol_jobTweet) = 1 then %>
		<th>Job Tweet</th>
        <%else %>
        
        <%end if %>


             <%if lto = "dencker" OR lto = "dencker" OR lto = "hidalgo" then %>
        <th>
        <b>Delprocesser</b>
        </th>
        <%end if %>

            <%if cint(visKol_planlagning) = 1 then %>
          <th>Planl�gning</th>
          <%end if%>


	</tr>
    </thead>
	<%
	
	
    thisMid = session("mid")

    if cint(stat0) = 1 then
	statKri = "jobstatus = 1"
	else
	statKri = "jobstatus = -1" 'ingen
	end if

	if cint(stat1) = 1 then
	statKri = statKri & " OR jobstatus = 2 "
	else
	statKri = statKri 
	end if
	
	if cint(stat2) = 1 then
	statKri = statKri & " OR jobstatus = 0 "
	else
	statKri = statKri
	end if

    if cint(stat3) = 1 then
	statKri = statKri & " OR jobstatus = 3 "
	else
	statKri = statKri
	end if


    if cint(stat4) = 1 then
	statKri = statKri & " OR jobstatus = 4 "
	else
	statKri = statKri
	end if

    if cint(stat5) = 1 then
	statKri = statKri & " OR jobstatus = 5 "
	else
	statKri = statKri
	end if
	
	if cint(medarb_jobans) <> 0 then

        kansKri = ""
        jobansKri = ""
        salgsansKri = ""

         if cint(jobansv) = 0 AND cint(salgsansv) = 0 AND cint(kansv) = 0 then 'Slet ingen valgt

            jobansKri = " AND (j.jobans1 = -1) " 'viser ingen

         else

	        if cint(jobansv) = 1 then
	        jobansKri = " AND ((j.jobans1 = "& medarb_jobans &" OR j.jobans2 = "& medarb_jobans &" OR j.jobans3 = "& medarb_jobans &" OR j.jobans4 = "& medarb_jobans &" OR j.jobans5 = "& medarb_jobans &") "
	        else
            jobansKri = " AND ((j.jobans1 <> -1) "
            end if

            if cint(salgsansv) = 1 then

                if cint(jobansv) = 1 then
                andor = "OR"
                else
                andor = "AND"
                end if

	        salgsansKri = " "& andor &" (j.salgsans1 = "& medarb_jobans &" OR j.salgsans2 = "& medarb_jobans &" OR j.salgsans3 = "& medarb_jobans &" OR j.salgsans4 = "& medarb_jobans &" OR j.salgsans5 = "& medarb_jobans &") "
	        else
            salgsansKri = ""
            end if

            if cint(kansv) = 1 then

                if cint(jobansv) = 1 OR cint(salgsansv) = 1 then
                andor = "OR"
                else
                andor = "AND"
                end if

	        kansKri = " "& andor &" (kundeans1 = "& medarb_jobans &" OR kundeans2 = "& medarb_jobans &"))"
            else
            kansKri = ")"
	        end if

        end if

	else
	kansKri = ""
	jobansKri = ""
    salgsansKri = ""
	end if
	
	if cint(visskjulte) = 1 then
	visskjulteKri = ""
	else
	visskjulteKri = " AND risiko >= 0"
	end if
	
    if cint(visKunFastpris) = 1 then
    fspSQLkri = " AND fastpris = 1"
    else
    fspSQLkri = " AND fastpris <> -1"
    end if

	jids = 0



    '** Kun job med timer/fakturaer/materiale reg. i periode ****'
     
   
    '**** Timer *****
     if cint(usedatoKri) = 3 OR cint(usedatoKri) = 5 then
    
     strSQLt = "SELECT tjobnr, j.id AS jid FROM timer AS t "_
     &" LEFT JOIN job AS j ON (j.jobnr = t.tjobnr) " _
     &" WHERE tdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"' AND ("& statKri &") GROUP BY tjobnr"
	 
   
     strJobMTimer = " AND (j.id = 0 " 

     'if session("mid") = 1 then
     'Response.write "usedatoKri: "& usedatoKri &" SQL:<br>"&  strSQLt &"<br><br>"& strJobMTimer &"<br>"
     'Response.flush
     'end if

     oRec3.open strSQLt, oConn, 3
     while not oRec3.EOF

     if instr(jobisWrt, ",#"& oRec3("jid") &"#") = 0 then
     strJobMTimer = strJobMTimer & " OR j.id = " & oRec3("jid")
     jobisWrt = jobisWrt & ",#"& oRec3("jid") &"#" 
     end if

     oRec3.movenext
     wend
     oRec3.close


    end if



     if cint(usedatoKri) = 3 OR cint(usedatoKri) = 6 then
    '** Fakturaer ****'
   
  
     strSQLt = "SELECT jobid, faknr FROM fakturaer AS f "_
     &" WHERE ((fakdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"' AND brugfakdatolabel = 0) "_
     &" OR (brugfakdatolabel = 1 AND labeldato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"' )) AND "_
	 &" (faktype = 0 OR faktype = 1) AND shadowcopy <> 1 GROUP BY jobid"
     
   
     strJobMTimer = strJobMTimer 
     strJobMFak = " AND (j.id = 0 " 

    
     oRec3.open strSQLt, oConn, 3
     while not oRec3.EOF

      if instr(jobisWrt, ",#"& oRec3("jobid") &"#") = 0 then
     strJobMTimer = strJobMTimer & " OR j.id = " & oRec3("jobid")
     strJobMFak = strJobMFak & " OR j.id = " & oRec3("jobid")
     jobisWrt = jobisWrt & ",#"& oRec3("jobid") &"#" 
     end if

    
    
     oRec3.movenext
     wend
     oRec3.close

    end if


     
        
      if cint(usedatoKri) = 3 then

     '**** Materialer ****'
     strSQLt = "SELECT jobid FROM materiale_forbrug AS m "_
     &" WHERE (forbrugsdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"') GROUP BY jobid"
     
   
     strJobMTimer = strJobMTimer 

   
     oRec3.open strSQLt, oConn, 3
     while not oRec3.EOF

      if instr(jobisWrt, ",#"& oRec3("jobid") &"#") = 0 then
     strJobMTimer = strJobMTimer & " OR j.id = " & oRec3("jobid")
     jobisWrt = jobisWrt & ",#"& oRec3("jobid") &"#" 
     end if

     oRec3.movenext
     wend
     oRec3.close



 

     end if


         if cint(usedatoKri) = 3 OR cint(usedatoKri) = 5 then

            strJobMTimer = strJobMTimer & ")"

        end if

         if cint(usedatoKri) = 6 then

            strJobMFak = strJobMFak & ")"

        end if



     if lto = "dencker" OR lto = "intranet - local" OR lto = "hidalgo" then
        fot = 100
        dim fomr_aid, fomr_anavn, fomr_abudgettimer, fomr_aktpris, fomr_fomr, fomr_sumtimer, form_aktstatus, form_lasttregdato, form_firstregdato 
        Redim fomr_aid(fot), fomr_anavn(fot), fomr_abudgettimer(fot), fomr_aktpris(fot), fomr_fomr(fot), fomr_sumtimer(fot), form_aktstatus(fot), form_lasttregdato(fot), form_firstregdato(fot)
     end if


      dim prgpid, pgrpnavn
      redim prgpid(0), pgrpnavn(0) 

    if cint(sorter) = 7 then 'Alle projektgrupper 

        
        'if cint(sorter) = 8 then
        'whSQL = "(id > 1 AND id <> 10)"
        'else
        whSQL = "id > 1"
        'end if

         
        strSQLp = "SELECT p.id As prgid, p.navn AS projektgruppenavn"_
        &" FROM projektgrupper AS p WHERE "& whSQL &" ORDER BY p.navn"
        p = 0
        oRec3.open strSQLp, oConn, 3
        while not oRec3.EOF 

        if instr(sorter_visprogrp, "#"&oRec3("prgid")&"#") <> 0 then
        redim preserve prgpid(p), pgrpnavn(p) 

        prgpid(p) = oRec3("prgid")
        pgrpnavn(p) = oRec3("projektgruppenavn")

        p = p + 1

        end if

        oRec3.movenext
        wend 
        oRec3.close
           
    else 
        
        p = 0 
      

    end if


     'Response.write "strJobMTimer "& strJobMTimer & "<br><br>"
     'Response.flush
  
       groupBY =  "j.id"
   
    
     c = 0
     for p = 0 TO UBOUND(prgpid)


         if cint(sorter) = 7 OR cint(sorter) = 8 then

        %>
	    <tr bgcolor="#8caae6">
		    <td style="padding:10px 0px 5px 5px; border-bottom:1px #FFFFFF solid;" colspan="<%=dtColSPan %>"><h4 style="color:#FFFFFF;"><span style="font-size:9px;">Projekgruppe<br /></span><%=pgrpnavn(p) %></h4></td>
	    </tr>
	    <%
     

    end if

	strSQL = "SELECT j.id, jobnavn, jobnr, jobknr, kkundenavn, jobstatus, fastpris, "_
	&" kkundenr, kid, jobslutdato, jobstartdato, j.beskrivelse, jobans1, m.mnavn, jobans2, m2.mnavn AS mnavn2,"_
    &" jobans3, m3.mnavn AS mnavn3,"_
    &" jobans4, m4.mnavn AS mnavn4,"_
    &" jobans5, m5.mnavn AS mnavn5,"_
	&" ikkebudgettimer, budgettimer, COALESCE(sum(r.timer), 0) AS restimer, stade_tim_proc, "_
	&" risiko, udgifter, rekvnr, forventetslut, restestimat, jobstatus, j.kommentar, s.navn AS aftnavn, "_
    &" jo_dbproc, "_
    &" (jo_bruttooms * jo_valuta_kurs/100) AS jo_bruttooms, jo_bruttooms AS jo_bruttooms_injobvaluta,"_
    &" (jo_udgifter_intern * jo_valuta_kurs/100) AS jo_udgifter_intern, "_
    &" (jo_udgifter_ulev * jo_valuta_kurs/100) AS jo_udgifter_ulev, "_
    &" (jo_bruttofortj * jo_valuta_kurs/100) AS jo_bruttofortj, jo_gnsfaktor, "_
    &" (jobtpris * jo_valuta_kurs/100) AS jobtpris, "_
    &" jobans_proc_1, jobans_proc_2, jobans_proc_3, jobans_proc_4, jobans_proc_5, virksomheds_proc, lukkedato, preconditions_met, jo_valuta, jo_valuta_kurs, project_tier"
	
        

    strSQL = strSQL &" FROM job j LEFT JOIN kunder ON (kid = jobknr) "_
	&" LEFT JOIN medarbejdere m ON (m.mid = jobans1)"_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = jobans2)"_
    &" LEFT JOIN medarbejdere m3 ON (m3.mid = jobans3)"_
    &" LEFT JOIN medarbejdere m4 ON (m4.mid = jobans4)"_
    &" LEFT JOIN medarbejdere m5 ON (m5.mid = jobans5)"_
    &" LEFT JOIN ressourcer_md r ON (r.jobid = j.id)"_
	&" LEFT JOIN serviceaft s ON (s.id = serviceaft)"

     if cint(usedatoKri) = 4 then
    statKri = replace(statKri, "OR jobstatus = 0", "OR (jobstatus = 0 AND "& stsldatoSQLKri &" >= '"& sqlDatoslut &"')")     
     end if

	strSQL = strSQL &" WHERE fakturerbart = 1 "& visskjulteKri &" "& fspSQLkri &" AND ("& statKri &") "& jobKri &" "
	
	if cint(usedatoKri) <> 2 AND cint(usedatoKri) <> 3 AND cint(usedatoKri) <> 4 AND cint(usedatoKri) <> 5 AND cint(usedatoKri) <> 6 then
	strSQL = strSQL &" AND "& stsldatoSQLKri &" BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
	end if

    if cint(usedatoKri) = 3 OR cint(usedatoKri) = 5 then
    strSQL = strSQL & strJobMTimer
    end if

    if cint(usedatoKri) = 6 then
    strSQL = strSQL & strJobMFak
    end if

            

    'if cint(usedatoKri) = 4 then
    'strSQL = strSQL &" AND (("& stsldatoSQLKri &" >= '"& sqlDatoslut &"' AND jobstatus = 0) OR jobstatus <> 0) "
    'end if

    if cint(sorter) = 7 OR cint(sorter) = 8 then
      strSQL = strSQL & " AND (j.projektgruppe1 = "& prgpid(p) &" OR j.projektgruppe2 = "& prgpid(p) &" OR j.projektgruppe3 = "& prgpid(p) &" OR j.projektgruppe4 = "& prgpid(p) &" OR "_
        &" j.projektgruppe5 = "& prgpid(p) &" OR j.projektgruppe6 = "& prgpid(p) &" OR j.projektgruppe7 = "& prgpid(p) &" OR j.projektgruppe8 = "& prgpid(p) &" OR j.projektgruppe9 = "& prgpid(p) &" OR j.projektgruppe10 = "& prgpid(p) &")"
    end if
	
	strSQL = strSQL & jobansKri & salgsansKri & kansKri & strPgrpSQLkri &" AND "& sqlKundeKri &" "_
	&" GROUP BY "& groupBY &" ORDER BY "& orderBY &", kkundenavn LIMIT "& ltoLimit 
	
	
        

    'if lto = "epi" AND (session("mid") = 1) then 'OR session("mid") = 59
    'if session("mid") = 1 AND lto = "synergi1" then
    'Response.write "strSQL: "& strSQL
	'Response.write "Hej David, jeg er ved at teste. Det er kun os to der kan se denne kode: <br>"& strSQL & "<br><br>"
	'response.write statKri
    'Response.flush
	'end if

	
	budgetIalt = 0
	budgettimerIalt = 0
	restimerIalt = 0
	lastMonth = 0
	tilfakturering = 0
	udgifterTot = 0
	timerforbrugtIalt = 0
	lastsortid = 0
    tilfaktureringWIP = 0
	salgsOmkFaktiskTot = 0
    totTerminBelobGrand = 0

    lastKid = 0

	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	
	jo_valuta_kurs = oRec("jo_valuta_kurs")
    nettoomstimer = oRec("jobtpris")

	faktureret = 0
	timerTildelt = 0
	timerforbrugt = 0
	totalforbrugt = 0
	totalBalance = 0
	jobbudget = 0
	
	if cint(oRec("jobstatus")) = 2 then '(passiv)
	stname = "<font class=megetlillesort>(passivt / Til fak.)</font> "
	else
	stname = ""
	end if
	
	bg = right(c, 1)
	select case bg
	case 0,2,4,6,8
	bgthis = "#ffffff"
	case else
	bgthis = "#eff3ff"
	end select
	
	
    level = session("rettigheder")
	
	'** Rettigheder til at redigere **
	if level = 1 then
	editok = 1
	else
			if cdbl(session("mid")) = oRec("jobans1") OR cdbl(session("mid")) = oRec("jobans2") OR (cdbl(oRec("jobans1")) = 0 AND cdbl(oRec("jobans2")) = 0) then
			editok = 1
			end if
	end if
	%>
	
	
	
	<!--<input id="hd_jobid" value="<%=jobid%>" type="hidden" />-->
	
	
	<% 
	'if isDate(oRec("forventetslut")) = true then
	'forvSlutDato = oRec("forventetslut") 
	'else
	'forvSlutDato = oRec("jobslutdato")
	'end if
	
	
	btop = 1
	
	select case sorter 
	case 2
	datoFilterUse = oRec("jobstartdato")
	case 1
	datoFilterUse = oRec("jobslutdato")
	case else
	datoFilterUse = oRec("jobslutdato")
	end select

  
        
    

    '*** Faktureret ****
    lastFakbrugt = 0
    
    

    call stat_faktureret_job(oRec("id"), sqlDatostart, sqlDatoslut) '** cls_fak


    '**** Realiserede timer **************
    '** Tot
    '** I periode
    '** Siden sidste faktura
    if cint(timertilfak) = 1 then 'kun timer til fakturering i valgte periode (siden sidste faktura)
	
        'response.write oRec("jobnr") &": faktureretLastFakDato: "& faktureretLastFakDato & "// startDatovlgt : " & startDatovlgt & "<br>" 

        if cDate(faktureretLastFakDato) > cdate(startDatovlgt) then
        sqlDatostRealTimer = year(faktureretLastFakDato) &"/"& month(faktureretLastFakDato) &"/"& day(faktureretLastFakDato)
        'startDatovlgt = sqlDatostRealTimer 
        lastFakbrugt = 1
        else
        sqlDatostRealTimer = sqlDatostart
        lastFakbrugt = 0
        end if

    else
    sqlDatostRealTimer = sqlDatostart
    end if

    '*** Realiseret timer og omkostninger
    call timeRealOms(oRec("jobnr"), sqlDatostRealTimer, sqlDatoslut, nettoomstimer, oRec("fastpris"), budgettimerIalt, aty_sql_realhours, jo_valuta_kurs) '** cls_timer
    
    timerforbrugtIalt = timerforbrugtIalt + timerforbrugt


        %>
        <!--timerforbrugt: <%=timerforbrugt%> // faktureret: <%=faktureret %> salgsOmkFaktisk: <%=salgsOmkFaktisk %> // skjulnuljob <%=skjulnuljob %> <br />-->
        <%


    if (timerforbrugt <> 0 AND cint(skjulnuljob) = 1) OR cint(skjulnuljob) = 0 then '((timerforbrugt <> 0 OR faktureret <> 0 OR salgsOmkFaktisk <> 0) AND

    %>
    <input type="hidden" name="FM_jobid" id="FM_jobid" value="<%=oRec("id")%>">
    <%

	
    if visKunGT <> 1 then


    '** orderBY Sorter efter dato
	if cint(sorter) = 0 OR cint(sorter) = 2 then
	
	    if datepart("m", datoFilterUse, 2, 3) <> lastMonth OR datepart("yyyy", datoFilterUse, 2, 3) <> lastYear then
	        if c <> 0 then%>
	         <!--<tr>
		    <td height="20" colspan="15">&nbsp;</td>
	        </tr>-->
	        <%end if 
                
                
            %>
	    <tr><!-- 8caae6 -->
		    <td style="padding:20px 0px 5px 10px; border-bottom:1px #999999 solid; background-color:#8dd7e8;" colspan="<%=dtColSPan %>"><h4><%=ucase(monthname(datepart("m", datoFilterUse, 2, 3))) &" "& year(datoFilterUse)%></h4></td>
	    </tr>
	    <%
	    btop = 0
	    end if
	
	lastMonth = datepart("m", datoFilterUse, 2, 3) 
	lastYear = datepart("yyyy", datoFilterUse, 2, 3) 
	
	end if



   
    '** orderBY Sorter efter kunde
    if cint(sorter) = 31 OR cint(sorter) = 3 then
	
	    if lastKid <> oRec("kid") then
	       
            if c <> 0 then%>
	         <!--<tr>
		    <td height="20" colspan="15">&nbsp;</td>
	        </tr>-->
	        <%end if 
                
                
            %>
	    <tr><!-- 8caae6 -->
		    <td style="padding:20px 0px 5px 10px; border-bottom:1px #999999 solid; background-color:#8dd7e8;" colspan="<%=dtColSPan %>"><h4><%=oRec("kkundenavn") & " ("& oRec("kkundenr") &")"%></h4></td>
	    </tr>
	    <%
	    btop = 0
	    end if
	
	lastKid = oRec("kid") 
	 
	
	end if

    end if





    '*********************************************************************************
    '*** Job hoved variable ****
    '*********************************************************************************

    	timerTildelt = oRec("ikkebudgettimer") + oRec("budgettimer")
		if len(timerTildelt) <> 0 then
		timerTildelt = timerTildelt
		else
		timerTildelt = 0
		end if
		
        budgettimerIalt = budgettimerIalt + timerTildelt


         if len(oRec("udgifter")) <> 0 then 'ulev + internkost
		udgifterpajob = oRec("udgifter")
		else
		udgifterpajob = 0
		end if

        udgifterTot = udgifterTot + udgifterpajob


        if len(oRec("jo_bruttooms")) <> 0 then
		jobbudget = oRec("jo_bruttooms")
		else
		jobbudget = 0
		end if
            
        jo_valuta = oRec("jo_valuta")      

        if len(oRec("restestimat")) <> 0 then
		restestimat = oRec("restestimat")
		else
		restestimat = 0
		end if
		
		
		'*************** WIP / Stade *******************
         stade_tim_proc = oRec("stade_tim_proc")

        if cint(historisk_wip) = 1 then 'brug historisk WIP
                
            call wip_historik_fn(oRec("id"),sqlDatoslut)
          
        end if


        select case stade_tim_proc '0 = rest i timer, 1 = i proc
		case 0
		totalforbrugt = (timerforbrugt + restestimat)
		totalBalance = (timerTildelt - totalforbrugt)
		case 1
		    
		    if timerforbrugt <> 0 AND restestimat <> 0 then
		    totalforbrugt = timerforbrugt * (100/restestimat) 
		    else
		    totalforbrugt = 0
		    end if '- (restestimat/timerTildelt) * 100)
		
		totalBalance = timerTildelt - totalforbrugt '(timerforbrugt + (timerTildelt - (restestimat/timerTildelt) * 100))
		end select


        if cint(stade_tim_proc) = 0 then 'rest estimat angivet i timer
		
		      if totalforbrugt <> 0 then
		
		        if restestimat = 0 then
		        stade = 100
        		else
        		stade = (timerforbrugt/totalforbrugt) * 100
		        end if
		    
		    else
		    
		    stade = 0
		    
		    end if
		    
		    stadeTxt = "Stade: "& formatnumber(stade, 0) & "% afsluttet"
		    afsl_proc = stade
		else
	        
	        stadeTxt = "Stade: " & formatnumber(restestimat, 0) & "% afsluttet"
		    afsl_proc = restestimat
		end if

        
	 
     '**************************************************************************************************


    if visKunGT <> 1 then
	%>
	
	
	
	<tr bgcolor="<%=bgthis%>">
         
		<td valign=top style="padding:6px 10px 3px 3px; width:200px; border-right:1px #cccccc solid; border-top:<%=btop%>px #cccccc solid;">

             <input type="hidden" name="SortOrder" class="SortOrder" value="<%=oRec("risiko")%>" />
		    <input type="hidden" name="rowId" value="<%=oRec("id")%>" />
             


            <% if sorter = 1 OR sorter = 31 OR sorter = 7 then %>
            <img src="../ill/pile_drag.gif" alt="Tr�k og sorter job" border="0" class="drag" />&nbsp;&nbsp;
            <%end if %>


            <%
            '** Vis kundenavn i kolonne eller som overskrift hvis der sorteres efter kunde
            if sorter <> 31 AND sorter <> 3 then %>
		        
                <%=oRec("kkundenavn")%>
                &nbsp;(<%=oRec("kkundenr") %>) 

            <%
                select case visSimpel
                case 0,1
                     %></td><td valign=top style="padding:6px 10px 3px 3px; width:200px; border-right:1px #cccccc solid; border-top:<%=btop%>px #cccccc solid;"><%
                case 2
                     %><br /><%
                end select 

            
             end if

            
            
		   
		    if print <> "j" then%>
		        <%if editok = 1 then%>
		        <a href="../timereg/jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=1&rdir=webblik"><%=oRec("jobnavn")%>&nbsp;(<%=oRec("jobnr")%>)&nbsp;</a>
		        <a href="../timereg/jobs.asp?menu=job&func=slet&id=<%=oRec("id")%>&jobnr_sog=a&filt=aaben&fm_kunde_sog=0&rdir=webblik" style="color:red;">[x]</a>
		        <%else%>
		        <b><%=oRec("jobnavn")%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)</b>
		        <%end if%>
		    <%else %>
		    <%=oRec("jobnavn")%>&nbsp;&nbsp;(<%=oRec("jobnr")%>)
		    <%end if %>


		
		<%
		
        if oRec("fastpris") = 1 then
        fastPrisTxt = "(fixed price)"
        %>
        <span style="color:yellowgreen; font-size:10px;"><%=fastPrisTxt %></span>
        <%

        else
        fastPrisTxt = "" '"Lbn. timer"
        end if

       

        if len(trim(oRec("aftnavn"))) <> 0 then
	    Response.Write "<br>Aft.: "& oRec("aftnavn") 
	    end if
        

        if len(trim(oRec("project_tier"))) <> 0 then
	    Response.Write "<br>Proj. tier: "& oRec("project_tier")
	    end if

            


        '** job beskrivelse ***'
        select case visSimpel
            case 0
            case 1,2


		if len(trim(oRec("beskrivelse"))) <> 0 then
		htmlparseCSV(oRec("beskrivelse")) 
		strBesk = htmlparseCSVtxt 
		else
		strBesk = ""
		end if

        if len(trim(strBesk)) > 0 then
		%>
		<span style="font-size:9px;"><i>
		<%if len(strBesk) > 100 then %>    
		<br /><%=left(strBesk, 100) %>..
		<%else %>
		<br /><%=strBesk %>
		<%end if %>
		</i></span><br />
        <%end if %>
            <%end select 






        
         select case visSimpel
         case 10,11
                %>
           </td><td valign=top style="padding:6px 10px 3px 3px; width:200px; border-right:1px #cccccc solid; border-top:<%=btop%>px #cccccc solid;">
                <% 
        case 0,1
       
        case 2
        end select
        %>
		<span style="color:#999999;"><i> 
        
        	<%if len(trim(oRec("mnavn"))) <> 0 then %>
        <%=oRec("mnavn")%>
        <%end if %>

         <%if oRec("jobans_proc_1")  <> 0 then %>
         (<%=oRec("jobans_proc_1") %>%)
         <%end if %>

		<%if len(trim(oRec("mnavn2"))) <> 0 then %>
		<%=", "&oRec("mnavn2")%> 
             <%if oRec("jobans_proc_2") <> 0 then %>
            (<%=oRec("jobans_proc_2") %>%)
            <%end if %>
		<%end if %>

        <%if len(trim(oRec("mnavn3"))) <> 0 then %>
		<%=", "&oRec("mnavn3")%> 
                 <%if oRec("jobans_proc_3") <> 0 then %>
                (<%=oRec("jobans_proc_3") %>%)
                <%end if %>
		<%end if %>

        <%if len(trim(oRec("mnavn4"))) <> 0 then %>
		<%=", "&oRec("mnavn4")%> 
             <%if oRec("jobans_proc_4") <> 0 then %>
            (<%=oRec("jobans_proc_4") %>%)
            <%end if %>
		<%end if %>

        <%if len(trim(oRec("mnavn5"))) <> 0 then %>
		<%=", "&oRec("mnavn5")%> 
             <%if oRec("jobans_proc_5") <> 0 then %>
            (<%=oRec("jobans_proc_5") %>%)
            <%end if %>
		<%end if %>

        <%
        if level = 1 then
            if lto <> "epi" OR (lto = "epi" AND thisMid = 6 OR thisMid = 11 OR thisMid = 59 OR thisMid = 1720) OR (lto = "epi_no" AND thisMid = 2) OR (lto = "epi_ab" AND thisMid = 2) OR (lto = "epi_sta" AND thisMid = 2) then 
                if oRec("virksomheds_proc") <> 0 then%>
                <br /><b><%=lto %>: </b> (<%=oRec("virksomheds_proc") %>%)
                <%end if %>
            <%end if %>
        <%end if %>
		</i></span>
		
        
        <%'Slut jobnavn og jobans 
        select case visSimpel
        case 1
        case else
        %></td>
        <%end select %>



		<% '**Rekvisitivions nr, aftalenr. pre konditiobner
          
         if cint(visKol_rekvnr) = 1 then
            
        %>
            <td valign=top style="padding:6px 10px 3px 3px; width:200px; border-right:1px #cccccc solid; border-top:<%=btop%>px #cccccc solid;">
            

       <!--
        select case lto
        case "xdencker"

        'if len(trim(oRec("rekvnr"))) <> 0 then
        %>
        Rekvnr.: <input type='text' id="aa_job_rekvnr_=oRec("id") %>" style='width:125px; font-size:10px;' value="=oRec("rekvnr")%>" /> 
       
        <span class='aa_job_rekvnr' id="=oRec("id")%>" style="background-color:#9ACD32; border:0px #999999 solid; font-size:10px; padding:2px; cursor:pointer;">Gem >></span>
        <span id="rekvnrGemt_=oRec("id") %>" style="display:none;">Gemt!</span>
        
        'response.write "Rekvnr.:"
        'end if
        -->

       <%

                    
		if len(trim(oRec("rekvnr"))) <> 0 then
            response.write "Rekvnr.: " & oRec("rekvnr")
        end if
		
		
		
		
		

                select case oRec("preconditions_met")
                case 0
                preconditions_met_Txt = ""
                case 1
                preconditions_met_Txt = "<br><span style='color:#6CAE1C; font-size:10px; background:#DCF5BD; width:200px;'>Pre-konditioner opfyldt</span>"
                case 2
                preconditions_met_Txt = "<br><span style='color:red; font-size:10px; background:pink; width:200px;'>Pre-konditioner ikke opfyldt!</span>"
                case else
                preconditions_met_Txt = ""
                end select


                %>
                <%=preconditions_met_Txt %>

            
            </td>
             

        <%
        end if 'cint(visKol_rekvnr) = 1 
        
        
        

            
		
		
         jo_udgifter_intern = oRec("jo_udgifter_intern")
        jo_udgifter_ulev = oRec("jo_udgifter_ulev")

        if timerTildelt <> 0 then
        udg_internKostprTim = oRec("jo_udgifter_intern") / timerTildelt
        else
        udg_internKostprTim = oRec("jo_udgifter_intern") / 1
        end if 
        
        'interKostEsti = 0
        'interKostEsti = (udg_internKostprTim/1 * totalforbrugt/1)

        
        select case visSimpel
        case 1
        %><br />
        <%case else%>
		 <td valign=top style="border-top:<%=btop%>px #cccccc solid; border-right:1px #cccccc solid; white-space:nowrap;">
        <%end select %>
        
        <%if print <> "j" then%>
        <input type="hidden" value="<%=oRec("id") %>" id="jobid_<%=c%>"/>

         <div class="form-inline" style="font-size:60%; padding-top:5px;">Fra:
		 <select name="FM_start_dag" id="FM_start_dag_<%=c%>" class="s_jobdato form-control input-small" style="width:50px;">
										<option value="<%=datepart("d",oRec("jobstartdato"), 2, 2)%>"><%=datepart("d",oRec("jobstartdato"), 2, 2)%></option> 
									   	<option value="1">1</option>
									   	<option value="2">2</option>
									   	<option value="3">3</option>
									   	<option value="4">4</option>
									   	<option value="5">5</option>
									   	<option value="6">6</option>
									   	<option value="7">7</option>
									   	<option value="8">8</option>
									   	<option value="9">9</option>
									   	<option value="10">10</option>
									   	<option value="11">11</option>
									   	<option value="12">12</option>
									   	<option value="13">13</option>
									   	<option value="14">14</option>
									   	<option value="15">15</option>
									   	<option value="16">16</option>
									   	<option value="17">17</option>
									   	<option value="18">18</option>
									   	<option value="19">19</option>
									   	<option value="20">20</option>
									   	<option value="21">21</option>
									   	<option value="22">22</option>
									   	<option value="23">23</option>
									   	<option value="24">24</option>
									   	<option value="25">25</option>
									   	<option value="26">26</option>
									   	<option value="27">27</option>
									   	<option value="28">28</option>
									   	<option value="29">29</option>
									   	<option value="30">30</option>
										<option value="31">31</option></select>
					
					<select name="FM_start_mrd" id="FM_start_mrd_<%=c%>" class="s_jobdato form-control input-small" style="width:50px;">
					<option value="<%=datepart("m",oRec("jobstartdato"), 2, 2)%>"><%=left(monthname(datepart("m",oRec("jobstartdato"), 2, 2)), 3)%></option>
					<option value="1">jan</option>
				   	<option value="2">feb</option>
				   	<option value="3">mar</option>
				   	<option value="4">apr</option>
				   	<option value="5">maj</option>
				   	<option value="6">jun</option>
				   	<option value="7">jul</option>
				   	<option value="8">aug</option>
				   	<option value="9">sep</option>
				   	<option value="10">okt</option>
				   	<option value="11">nov</option>
				   	<option value="12">dec</option></select>
					
					
					<select name="FM_start_aar" id="FM_start_aar_<%=c%>" class="s_jobdato form-control input-small" style="width:50px;">
					<option value="<%=datepart("yyyy",oRec("jobstartdato"), 2, 2)%>"><%=right(datepart("yyyy",oRec("jobstartdato"), 2, 2), 2)%></option>
					<%for x = -5 to 10
		            useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		            <option value="<%=right(useY, 2)%>"><%=right(useY, 2)%></option>
		            <%next %>
					</select>
             <span id="Span1" style="color:green; font-size:12px; visibility:hidden;" ><i>V</i></span>
            </div>

         

		
							
							
		<%else %>
				
		<%=datepart("d",oRec("jobstartdato"), 2, 2)%>.<%=datepart("m",oRec("jobstartdato"), 2, 2)%>.<%=datepart("yyyy",oRec("jobstartdato"), 2, 2)%>
				
		<%end if %>


					
					<%if print <> "j" then%>
                    <div class="form-inline" style="padding-top:5px; font-size:60%;">Til:&nbsp;
					<select name="FM_slut_dag" id="FM_slutt_dag_<%=c%>" class="s_jobdato form-control input-small" style="width:50px;"  >
										<option value="<%=datepart("d",oRec("jobslutdato"), 2, 2)%>"><%=datepart("d",oRec("jobslutdato"), 2, 2)%></option> 
									   	<option value="1">1</option>
									   	<option value="2">2</option>
									   	<option value="3">3</option>
									   	<option value="4">4</option>
									   	<option value="5">5</option>
									   	<option value="6">6</option>
									   	<option value="7">7</option>
									   	<option value="8">8</option>
									   	<option value="9">9</option>
									   	<option value="10">10</option>
									   	<option value="11">11</option>
									   	<option value="12">12</option>
									   	<option value="13">13</option>
									   	<option value="14">14</option>
									   	<option value="15">15</option>
									   	<option value="16">16</option>
									   	<option value="17">17</option>
									   	<option value="18">18</option>
									   	<option value="19">19</option>
									   	<option value="20">20</option>
									   	<option value="21">21</option>
									   	<option value="22">22</option>
									   	<option value="23">23</option>
									   	<option value="24">24</option>
									   	<option value="25">25</option>
									   	<option value="26">26</option>
									   	<option value="27">27</option>
									   	<option value="28">28</option>
									   	<option value="29">29</option>
									   	<option value="30">30</option>
										<option value="31">31</option></select>
					
					<select name="FM_slut_mrd" id="FM_slutt_mrd_<%=c%>" class="s_jobdato form-control input-small" style="width:50px;">
					<option value="<%=datepart("m",oRec("jobslutdato"), 2, 2)%>"><%=left(monthname(datepart("m",oRec("jobslutdato"), 2, 2)), 3)%></option>
					<option value="1">jan</option>
				   	<option value="2">feb</option>
				   	<option value="3">mar</option>
				   	<option value="4">apr</option>
				   	<option value="5">maj</option>
				   	<option value="6">jun</option>
				   	<option value="7">jul</option>
				   	<option value="8">aug</option>
				   	<option value="9">sep</option>
				   	<option value="10">okt</option>
				   	<option value="11">nov</option>
				   	<option value="12">dec</option></select>
					
					
					<select name="FM_slut_aar" id="FM_slutt_aar_<%=c%>" class="s_jobdato form-control input-small"  style="width:50px;">
					<option value="<%=datepart("yyyy",oRec("jobslutdato"), 2, 2)%>"><%=right(datepart("yyyy",oRec("jobslutdato"), 2, 2), 2)%></option>
					<%for x = -5 to 10
		            useY = datepart("yyyy", dateadd("yyyy", x, date()))%>
		            <option value="<%=right(useY, 2)%>"><%=right(useY, 2)%></option>
		            <%next %></select>

                        <span id="sp_dtopd_<%=c%>" style="color:green; font-size:12px; visibility:hidden;" ><i>V</i></span>

                
					
						
                    </div>

				<%else %>
				
				- <%=datepart("d",oRec("jobslutdato"), 2, 2)%>.<%=datepart("m",oRec("jobslutdato"), 2, 2)%>.<%=datepart("yyyy",oRec("jobslutdato"), 2, 2)%>
				
				<%end if %>



           


		<%
		  select case visSimpel
            case 0
            %>
            </td><td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc solid; border-right:1px #cccccc solid; white-space:nowrap;" >
            <%
            case 1
             %>
            <td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc solid; border-right:1px #cccccc solid; white-space:nowrap;" >
            <%
            case 2, 3
                %>
                <br /> <!--<br />-->
                <%
            end select




		stCHK0 = ""
		stCHK1 = ""
		stCHK2 = ""
        stCHK3 = ""
        stCHK4 = ""
        stCHK5 = ""
        lkDatoThis = "" 
		
		
		select case oRec("jobstatus")
		case 1
		stCHK1 = "SELECTED"
		stName = "Aktiv"
		case 2
		stCHK2 = "SELECTED"
        stName = "Passiv / Til fak."
        case 3
		stCHK3 = "SELECTED"
		stName = "Tilbud"
        case 4
        stCHK4 = "SELECTED"
        stName = "Gennemsyn"
        case 5
        stCHK5 = "SELECTED"
        stName = "Evaluering"
		case 0
		stCHK0 = "SELECTED"
		stName = "Lukket"
                    if cdate(oRec("lukkedato")) <> "01-01-2002" then
                    lkDatoThis = " ("& formatdatetime(oRec("lukkedato"), 2) & ")"
                    end if
		end select

      
		
		if print <> "j" then
                    
        if visSimpel = 2 then%>
                Status:<br />
                <%end if %>

            <%select case lto
            case "mpt", "local - intranet"

                if level = 1 then
                wprotec = ""
                else
                
                    if stCHK0 = "CHECKED" then 'Hvis job er lukket m� kun admin �ndre status
                    wprotec = "readonly"
                    else
                    wprotec = ""
                    end if

                end if

            case else

                if level <= 2 OR level = 6 then
                wprotec = ""
                else

                    if stCHK0 = "CHECKED" then
                    wprotec = "readonly"
                    else
                    wprotec = ""
                    end if

                end if

         end select  %> 

		<select name="FM_jobstatus" id="FM_jobstatus_<%=c%>" class="s_jobstatus form-control input-small" style="width:80px;">
		<option value="0" <%=stCHK0%>>Lukket <%=lkDatoThis%></option>

        <%if wprotec <> "readonly" then %>
		<option value="1" <%=stCHK1%>>Aktiv</option>
		<option value="2" <%=stCHK2%>>Passiv / Til fak.</option>
        <option value="3" <%=stCHK3%>>Tilbud</option>
        <option value="4" <%=stCHK4%>>Gennemsyn</option>
        <option value="5" <%=stCHK5%>>Evaluering</option>
        <%end if %>
		</select>

                <span id="sp_stopd_<%=c%>" style="color:green; font-size:12px; visibility:hidden;" ><i>V</i></span>

		<%else %>
		
		<%=stName%>

                <%if oRec("jobstatus") = 0 then 'lukket %>
                <br /><%=lkDatoThis %>

                <%end if %>
		
		<%end if %>
		
        <%
		

        '** Billlede 
        if lto = "dencker" OR lto = "intranet - local" OR lto = "hidalgo" then %>
        <br />
            
            <%
                findesfil = 0
                filnavn = "blank.gif"
                strSQLf = "SELECT filnavn FROM filer WHERE filertxt = '"& oRec("id") &"_beskrivelsesdokument' ORDER BY id DESC"
                oRec9.open strSQLf, oConn, 3
                if not oRec9.EOF then
                    findesfil = 1
                    filnavn = oRec9("filnavn")
                end if
                oRec9.close    
                

                if cint(findesfil) = 1 then
                %>
                <img src="../inc/upload/<%=lto%>/<%=filnavn %>" width="100" height="100" />
                <%
                End if
            %>

        </td>
         <%end if %>


           
        <%
		  select case visSimpel
            case 0,1,2
            %>
            </td>
            <% if cint(visKol_prioritet) = 1 then %>
            <td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc solid; border-left:1px #cccccc solid; border-right:1px #cccccc solid; white-space:nowrap;" >
            <%end if %>
            <%
            case 21
            end select
		
		
		'if cdbl(lastsortid) >= oRec("risiko") AND oRec("risiko") >= 0 then
		'thissortID = lastsortid + 1
		'else
		thissortID = oRec("risiko")
		'end if
		
		'if oRec("risiko") >= 0 then
		'lastsortid = thissortID
		'end if
		
		if print <> "j" then%>
		
	
		 <%if sorter = 1 OR sorter = 31 OR sorter = 7 then
		 prioFMType = "hidden"

            if cint(visKol_prioritet) = 1 then
            %><%=thissortID%><%
            end if
         else
		 prioFMType = "text"
		 end if
             
             
         if cint(visKol_prioritet) = 0 then
             prioFMType = "hidden"
         end if
         %>
         <input name="FM_risiko" id="FM_risiko_<%=c %>" type="<%=prioFMType %>" value="<%=thissortID%>" class="s_prio form-control input-small form-inline" style="width:40px;" />
         
         
         <%if cint(visKol_prioritet) = 1 then %>
         <span id="sp_propd_<%=c %>" style="color:green; font-size:12px; visibility:hidden;" ><i>V</i></span><br />
      
                  <!-- <input name="FM_SortOrder" id="SortOrder" type="text" value="<%=thissortID%>"/>     --> 

        
         
         <%
         '** Gantt grafisk modtpunk ***
         startdate = dateadd("m", -4, oRec("jobslutdato"))
         start_dag = day(startdate)
         start_mrd = month(startdate)
         start_aar = year(startdate)

         slut_dag = day(oRec("jobslutdato"))
         slut_mrd = month(oRec("jobslutdato"))
         slut_aar = year(oRec("jobslutdato"))
          
             if sorter = 1 OR sorter = 31 OR sorter = 7 OR cint(visKol_prioritet) = 0 then

             else
             %><br /><%
             end if
             %>

                
        <!--
         <a href="webblik_joblisten21.asp?nomenu=1&FM_kunde=<%=oRec("jobknr") %>&jobnr_sog=<%=oRec("jobnr") %>&FM_start_dag=<%=start_dag%>&FM_start_mrd=<%=start_mrd%>&FM_start_aar=<%=start_aar%>&FM_slut_dag=<%=slut_dag%>&FM_slut_mrd=<%=slut_mrd%>&FM_slut_aar=<%=slut_aar%>&FM_usedatokri=1" target="_blank" class=rmenu>+ Gantt</a>
                -->
					
        
         
		<%else %>

                <%if cint(visKol_prioritet) = 1 then%>
		        <%=thissortID%>
                <%end if %>

        <%end if 'if cint(visKol_prioritet) = 1 then
            
        end if 'print%>


		</td>


       



        <!-- Forecast --->
        <%if cint(visKol_forecast) = 1 then%>
        <td valign=top style="padding:6px 3px 3px 3px; white-space:nowrap; border-top:<%=btop%>px #cccccc solid; border-right:1px #CCCCCC solid; border-left:1px #CCCCCC solid;">Forecast 
            
           
                
            <% 
             if aktBudgettjkOn_afgr = 1 OR aktBudgettjkOn_afgr = 2 then

             aktid = 0
             timerTastet = 0
             usemrn = 0
             useDateStSQL = strAar_slut &"-"& strMrd_slut &"-"& strDag_slut
                
           
               datoKrionly = 0
               call ressourcetimerTildelt(useDateStSQL, oRec("id"), 0, usemrn, datoKrionly)

               'call ressourcefc_tjk(ibudgetaar, ibudgetmd, aar, md, usemrn, aktid, timerTastet)

             %>
            <%if aktBudgettjkOn_afgr = 1 then %>
            <br /> FY <%=strAar_slut %>:
            <%end if %>

            <%if aktBudgettjkOn_afgr = 2 then %>
            <br />MD <%=strMrd_slut %>:
            <%end if %>

            <a href="ressource_belaeg_jbpla.asp" target="_blank"><%=formatnumber(restimerAll, 2)%> t.</a>
            <br /><span style="font-size:9px;">GT: (<%=formatnumber(oRec("restimer"), 2)%>)</span>
            <%else %>

            :<br />
            <a href="ressource_belaeg_jbpla.asp" target="_blank"><%=formatnumber(oRec("restimer"), 2)%> t.</a>
            <%end if %>

        </td>
		<%end if 'cint(visKol_forecast) = 1 then






         %>
         <!-- Timer realiseret -->

          <%'select case visSimpel 
           'case 0
           'bdrightreal = 0
            '  bdleftreal = 1
            '  case 1,2
           bdrightreal = 1
              bdleftreal = 0
            '  end select
              
              
              
        if timerTildelt <> 0 then
		projektcomplt = ((timerforbrugt/timerTildelt)*100)
              
              if timerTildelt > 1 then     
              forkalkWdt = 100
              else
              forkalkWdt = 1
              end if
	    else
	    projektcomplt = 100
              forkalkWdt = 0
	    end if
            	    

              forkalkWdt = 100

              if projektcomplt > 200 then
              forkalkWdt = 50
              end if

              if projektcomplt > 1000 then
              forkalkWdt = 25
              end if

              if projektcomplt > 10000 then
              forkalkWdt = 10
              end if


	    if projektcomplt >= 100 then
            
		    fcol = "crimson"
		    showprojektcomplt = projektcomplt
            fntCol = "#999999"

		    projektcomplt = 135

		else
            if projektcomplt >= 50 then
            fcol = "#6CAE1C"
		    showprojektcomplt = projektcomplt
            fntCol = "#999999"
            else
		    fcol = "#DCF5BD"
		    showprojektcomplt = projektcomplt
            fntCol = "#999999"
            end if
		end if %>



        <!-- Realiseret -->
        <td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc solid; border-left:<%=bdleftreal%>px #CCCCCC solid; border-right:<%=bdrightreal%>px #cccccc solid; white-space:nowrap;">
       
       



            <table border="0" cellspacing="1" cellpadding="0" width="100%">
            <tr>
                
                
            <td valign="top">
            Real.: <span style="float:right;"><%=formatnumber(timerforbrugt, 2)%> t.</span>


                <% if lto ="dencker" OR lto = "hidalgo" then

                  

                  strSQL2 = "SELECT sum(t.timer) AS ubemandet FROM timer t WHERE t.tjobnr = '"& oRec("jobnr") &"' AND tfaktim = 90"

                    if realfakpertot <> 0 then
		            strSQL2 = strSQL2 &" AND tdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"'"
		            end if


                 strSQL2 = strSQL2 & " GROUP BY tjobnr"

                 ubemandet = 0
                  oRec10.open strSQL2, oConn, 3
                  if not oRec10.EOF then
                    ubemandet = oRec10("ubemandet")
                  end if
                  oRec10.close

                    if ubemandet <> 0 AND isNull(ubemandet) <> true then
                    ubemandet = ubemandet
                    else
                    ubemandet = 0
                    end if

                    %><br />Ubem.: <span style="float:right;"><%=formatnumber(ubemandet, 2)%> t.</span>
                    <br />Ialt: <span style="float:right;"><u><%=formatnumber(timerforbrugt+ubemandet, 2)%></u> t.</span><br />

                    Fork.: <span style="float:right;"><%=formatnumber(timerTildelt, 2) %> t.</span>&nbsp;<br />

                <%end if
                    %>

                   <span style="font-size:9px; color:#999999;">
            <% if cint(realfakpertot) <> 0 then %> 
              
           <!-- <b>I periode <br />-->
            <%if cint(timertilfak) = 1 AND cint(lastFakbrugt) = 1 then %>
            <br /><span style="color:red;">Sidste Faktura sys.dato:</span> 
             <%end if %>

                 
                
           <br />
                       
         

                    <%if cint(lastFakbrugt) = 1 then%> 
                    <%=formatdatetime(faktureretLastFakDato, 2) &" - "& formatdatetime(slutDatovlgt, 2) %>
                    <%else %>
                    <%=formatdatetime(sqlDatostRealTimer, 2) &" - "& formatdatetime(slutDatovlgt, 2) %>
                    <%end if %>


            <%else %>
            <!--Total:-->
            <%end if %>
            </span>

            </td>




               


            </tr>

            <%if cdbl(showprojektcomplt) > 0 AND cdbl(timerforbrugt) > 0 then%>
            <tr><td style="padding-top:1px;">
           
		    Fremdrift:<br />
               
		
	        <div style="border-bottom:5px <%=fcol%> solid; font-size:12px; white-space:nowrap; color:<%=fntCol%>; width:<%=cint(left(projektcomplt, 3))%>; padding:0px 0px 0px 3px;">
            <%=formatpercent(showprojektcomplt/100, 0)%></div>	
            <div style="border-top:5px #999999 solid; width:<%=forkalkWdt%>px; padding:0px 0px 0px 0px;"></div>	
		
        
            </td>

            
            </tr>

            <%end if%>

            </table>



        <%
        if lto <> "dencker" AND lto <> "intranet - local" AND lto <> "hidalgo" then

        if timerforbrugt > 0 then


       '** Timeforbrufg real de sidste 4 md. 
            
            select case month(startDatovlgt4md)
            case 1,3,5,7,8,10,12
            dainmo = 31
            case 2
            dainmo = 28
            case else
            dainmo = 30
            end select
                  
       select case visSimpel
       case 0

       case 1

            

        %>
        <br />

             <span style="font-size:9px;"><b>Timefordeling pr. md<br /> 1. <%=left(monthname(month(startDatovlgt4md)), 3) &" "& year(startDatovlgt4md) & " - "& dainmo &". "& left(monthname(month(slutDatovlgt)), 3) &" "& year(slutDatovlgt)%>:</b>


            <%if cint(bdgmtypon_val) = 1 then  %>
            (realtid)
            <%end if %>

        </span>
        <%
       call timeforbrug4md_2015
    
       case 2
      
       '** Fordeling p� medarbejder / medarb.type    
        'if cint(bdgmtypon_val) <> 1 then 'ikke dme med budget p� medarb. typer da de er med i budget kolonne og derved vlil bleve dobbelt persformance
        %>
          <br /><%
        call timer_fordeling_medarb_typer(oRec("jobnr"),realfakpertot, timerforbrugt, sqlDatostRealTimer, slutDatovlgt, oRec("id")) 'startDatovlgt �ndret til startDatovlgt 12.3.2015
        'end if%>
              

        <br />
        <span style="font-size:9px;"><b>Timefordeling pr. md<br /> 1. <%=left(monthname(month(startDatovlgt4md)), 3) &" "& year(startDatovlgt4md) & " - "& dainmo &". "& left(monthname(month(slutDatovlgt)), 3) &" "& year(slutDatovlgt)%>:</b>

              <%if cint(bdgmtypon_val) = 1 then  %>
            (realtid)
            <%end if %>

        </span>
        <%
       call timeforbrug4md_2015

       end select  


      

		
		else
		
		Response.Write "&nbsp;"
		
		end if
		end if 'lto
		%>
		
		</td>
		

                      <%select case visSimpel 
                        case 0
                        case 1,2
                        %>

                    <!-- Stadeindmelding / Restestimat / Stade -->
		            <td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc solid;">
                    
		            <%
	
		
		
		
		                    if totalforbrugt <> 0 then
		
		                        if totalBalance >= 0 then
		                        bgst = "#DCF5BD"
                                ill = "dot_gron.gif"
		                        else
		                        bgst = "crimson"
		                        ill = "dot_rod.gif"
                                end if
		
		                    else
		
		                    bgst = "#cccccc"
                            
                            ill = "dot_graae.gif"
		
		                    end if %>
		
		
		            <%if print <> "j" AND cint(historisk_wip) <> 1 then%>
                            <div class="form-inline">

		                    <input type="text" name="FM_restestimat" id="FM_restestimat_<%=c%>"  value="<%=restestimat%>" style="width:60px;" class="form-control input-small">
		                    <%select case oRec("stade_tim_proc")
		                    case 0
                                if restestimat <> 0 then '**% forvalgt
		                        stade_timproc_SEL0 = "SELECTED"
		                        stade_timproc_SEL1 = ""
                                else
                                stade_timproc_SEL0 = ""
		                        stade_timproc_SEL1 = "SELECTED"
                                end if
		                    case 1
		                    stade_timproc_SEL0 = ""
		                    stade_timproc_SEL1 = "SELECTED"
		                    end select %>
		                    <select class="stade form-control input-small" name="FM_stade_tim_proc" id="FM_stade_ti_pr_<%=c%>" style="width:80px;">
		                        <option value=0 <%=stade_timproc_SEL0 %>>timer til rest.</option>
		                        <option value=1 <%=stade_timproc_SEL1 %>>% afsluttet</option>
		                    </select>

                            
                          
		
		            <%'else 
		
		                    'select case oRec("stade_tim_proc")
		                    'case 0
		                    'stade_timproc_txt = "timer til rest."
		                    'case 1
		                    'stade_timproc_txt = "% afsluttet"
		                    'end select
		                    %>
		                    <%'restestimat & " " & stade_timproc_txt%> 

                                <!--if cint(historisk_wip) = 1 then %>
                                    <br /><span style="color:#999999;">(Wip dato: <%=wipHistdato %>)</span>
                                <!--end if%>
                                    -->


		            <%end if %>
		
		
		              <br />
                            <span style="font-size:9px; white-space:nowrap;" id="wiphist_<%=c%>">
                              WIP historik (seneste 5):<br />
                                <%
                                strSelWIp = "SELECT w.dato, medid, init, stade_tim_proc, restestimat FROM wip_historik w "_
                                &" LEFT JOIN medarbejdere ON (mid = medid) "_
                                &" WHERE jobid = "& oRec("id") &" ORDER BY dato DESC limit 5"

                                oRec10.open strSelWip, oConn, 3
                                while not oRec10.EOF 

                                if oRec10("stade_tim_proc") = 1 then
                                restestimatTxt = oRec10("restestimat") & " %"
                                else
                                restestimatTxt = oRec10("restestimat") & " t. rest"
                                end if

                                %>
                                <%=left(oRec10("dato"), 16) & " <b>"& restestimatTxt  &"</b> - " & oRec10("init") %><br />
                                <%

                               oRec10.movenext
                               wend
                               oRec10.close%>
                                </span>
                              
                            </div>



                    <%if timerTildelt <> 0 and totalforbrugt <> 0 then
		
		                    if timerTildelt < totalforbrugt then
		                    prc = 100 - (timerTildelt/totalforbrugt) * 100
		                    else
		                    prc = 100 - (totalforbrugt/timerTildelt) * 100
		                    end if
		    
		 
		            end if
		            %>



       

            
       
                    </td>
		


                    <td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc solid; border-right:1px #cccccc solid;" align=right>
		            <div id="divindikator_<%=c%>" style="width:20px; background-color:<%=bgst%>; height:20px;">
                        <img src='../ill/<%=ill %>' width='20' height='20' alt='' border='0'>
                    </div>
                          
                        <%if print <> "j" AND cint(historisk_wip) <> 1 then%>
                        <div style="padding-top:5px;">
                        <div id="bn_restestimat_<%=c%>" class="rest" style="background-color:#CCCCCC; border:1px #999999 solid; padding:2px; float:right"> >> </div> <!-- btn btn-outline-success btn-sm -->
                        <span id="sp_esopd_<%=c %>" style="color:green; font-size:12px; visibility:hidden;" ><i>V</i></span>
                            </div>
                        <%end if %>

	                </td>
        
                    <%end select %>
		

         <input id="FM_timerreal_<%=c%>" type="hidden" value="<%=timerforbrugt%>" />
         <input id="FM_forkalk_<%=c%>" type="hidden" value="<%=timerTildelt%>" />





        	




        <%select case visSimpel 
        case 0,1
        %>

        <!-- budget -->
	    <%
            forkalkvalue = timerTildelt
            bruttooms = oRec("jo_bruttooms_injobvaluta") 'jobbudget


        if lto <> "dencker" AND lto <> "hidalgo" then
        %>
		<td valign="top" style="border-top:<%=btop%>px #cccccc solid;">
         

            <div class="form-inline">
            
                    <input id="FM_forkalk_<%=c %>" type="text" class="s_forkalk form-control input-small" value="<%=forkalkvalue %>" style="width:65px; text-align:right;"/>&nbsp;t.
                    <input id="FM_brutoms_<%=c %>" type="text" class="s_brutoms form-control input-small" value="<%=bruttooms %>" style="width:70px; text-align:right;"/>
           
                      <%
                                     felt = "FM_jo_valuta_"& c
                                     call valutaList(jo_valuta, felt)
                                     %>
       
                        <span id="sp_forkalk_<%=c %>" style="color:green; font-size:12px; visibility:hidden;" ><i>V</i></span>
         
                  


            </div>

		</td>
        <%end if %>
        
		
       	
       <!--  <td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc solid; border-left:1px #cccccc solid; white-space:nowrap;">
               <%=formatnumber(timerTildelt, 2) & " t." %>hej<br />
               <%=formatnumber(jobbudget, 2) &" "& basisValISO_f8%> 
        </td> -->
        <%
        case 2%>

        <td valign=top  style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc solid; border-right:1px #cccccc solid; white-space:nowrap;">
        Budget  <%if editok = 1 then%>
		        <a href="../timereg/jobs.asp?menu=job&func=red&id=<%=oRec("id")%>&int=1&rdir=webblik&showdiv=forkalk" class=rmenu target="_blank">(+ rediger)</a> 
                <%end if %>


        <table width=100% cellspacing=1 cellpadding=1 bgcolor="#cccccc">
        <tr>
        
        <td  style="background-color:#FFFFFF;">Timer:</td>
        <td  align=right style="background-color:#FFFFFF; white-space:nowrap;">
        <%="<b>"& formatnumber(timerTildelt, 2) & " t.</b>" %>
        </td></tr>

        <tr>
        <td  style="background-color:#F7f7f7; vertical-align:top;">Bruttooms:</td>
        <td  align=right style="background-color:#F7f7f7; white-space:nowrap;">
		<b><%=formatnumber(jobbudget, 2) &" "& basisValISO_f8%> </b>

            <%if cint(oRec("jo_valuta")) <> cint(basisValId) then 
                 call beregnValuta(jobbudget,100,oRec("jo_valuta_kurs"))
                jobbudget_opr_curr = valBelobBeregnet

                call valutakode_fn(oRec("jo_valuta"))
                 %>

            <br />(<%=formatnumber(jobbudget_opr_curr, 2) & " "& valutaKode_CCC_f8  %>)

            <%end if %>

		</td></tr>

        <tr>
        <td  style="background-color:#FFFFFF;">Nettooms:</td>
        <td  align=right style="background-color:#FFFFFF; white-space:nowrap;">
		<%=formatnumber(oRec("jobtpris"), 2) &" "& basisValISO_f8%>
		</td></tr>
        
        <%if cint(bdgmtypon_val) = 1 then 
            isGrpWritten = "#0#"
            for t = 1 to UBOUND(mtypgrpids) 

            'Response.write "mtypgrpnavn(t): "& mtypgrpnavn(t) & "<br>"
                         if len(trim(mtypgrpnavn(t))) <> 0 then
                        
                                if instr(isGrpWritten, "#"&mtypgrpid(t)&"#") = 0 then
                                thisMtypGrpBelIg(mtypgrpid(t)) = 0
                                end if

                                strSQLmtypbgt = "SELECT SUM(belob) AS belob FROM medarbejdertyper_timebudget WHERE jobid = "& oRec("id") &" AND ("& mtypgrpids(t) &") GROUP BY jobid"
            
                                'if session("mid") = 1 then
                                'Response.write strSQLmtypbgt & " timer: "& thisMtypGrpBel(mtypgrpid(t)) &"<br>"
                                'Response.flush
                                'end if
            
                                oRec6.open strSQLmtypbgt, oConn, 3
                                if not oRec6.EOF then
                                thisMtypGrpBelIg(mtypgrpid(t)) = thisMtypGrpBelIg(mtypgrpid(t)) + formatnumber(oRec6("belob"), 2)

                                end if
                                oRec6.close
            
                               

                        end if
                          
                        isGrpWritten = isGrpWritten & ",#"& mtypgrpid(t) &"#"
            
            next

                          

                      for mtgp = 1 to UBOUND(kunMtypgrp) 'mtypgrpids 
            
                                      if len(trim(kunMtypgrpNavn(mtgp))) <> 0 then  %>

                                      <tr>
                                    <td  style="background-color:#FFFFFF; color:#999999;"><%=kunMtypgrpNavn(mtgp) %>:</td>
                                    <td  align=right style="background-color:#FFFFFF; white-space:nowrap; color:#999999;">
             
                                       

		                            <%=formatnumber(thisMtypGrpBelIg(kunMtypgrp(mtgp)), 2) &" "& basisValISO_f8%>
		                            </td></tr> 
                                   <%
                                        end if
                        next%>
        <%
        end if



        if timerTildelt <> 0 then
        budget_gns_timepris = (oRec("jobtpris")/timerTildelt)
        else
        budget_gns_timepris =  0
        end if %>
         <tr>
         
         <td  style="background-color:#F7F7F7;">Timepris:</td>
         <td  align=right style="background-color:#F7F7F7; white-space:nowrap;">
		 <%=formatnumber(budget_gns_timepris, 2)&" "&basisValISO_f8 %>/t.
		</td></tr>
		
	
         <tr>
         <td  style="background-color:#FFFFFF;">Salgsomk.:</td>
         <td  align=right style="background-color:#FFFFFF; white-space:nowrap;">
		 <%=formatnumber(jo_udgifter_ulev, 2) &" "& basisValISO_f8%> 
         </td>
         </tr>

        <tr>
         <td  style="background-color:#F7F7F7;">Intern omk.:</td>
        <td  align=right style="background-color:#F7F7F7; white-space:nowrap;">
        <%=formatnumber(jo_udgifter_intern, 2) &" "& basisValISO_f8%> 
        </td></tr>

        <%
        if timerTildelt <> 0 then
        budget_gns_kostpris = (jo_udgifter_intern/timerTildelt)
        else
        budget_gns_kostpris =  0
        end if %>

         <tr>
         <td  style="background-color:#FFFFFF; white-space:nowrap;">Kost. pr. time.:</td>
        <td  align=right style="background-color:#FFFFFF; white-space:nowrap;">
        <%=formatnumber(budget_gns_kostpris, 2) &" "& basisValISO_f8 %> /t.
        </td></tr>


         <%
        
        'joDbbelob = (jobbudget - (jo_udgifter_ulev + jo_udgifter_intern)) 
        
        joDbbelob = oRec("jo_bruttofortj")%>

          <tr><td  style="background-color:#FFFFFF;">DB bel�b: </td>
          <td  align=right style="background-color:#FFFFFF;">
         <b><%=formatnumber(joDbbelob,0) & " "& basisValISO_f8 %></b> 
      </td></tr>

        <tr>
         <td style="background-color:#ffdfdf;">DB: </td>
        <td  align=right style="background-color:#ffdfdf; white-space:nowrap;">
        <b><%=formatnumber(oRec("jo_dbproc"),0) %> %</b>
         </td></tr>
		



         <%
        '**** Balance ***'         
        bal = (faktureret - jobbudget)
		
		if bal < 0 then
		fcol = "red"
		else
		fcol = "green"
		end if
		 %>
       
             <tr>
         <td>Budget / Fak.: </td>
        <td align=right style="white-space:nowrap;">
        <b><%="<span style='color:"& fcol &";'>" & formatnumber(bal, 2)%></span> <%=basisValISO_f8 %></b>
         </td></tr>
		
		
        </table>

        </td>
	
		

        <!-- Realiseret --> 
		<td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc solid; white-space:nowrap; border-right:1px #cccccc solid;" >
        <%=tsa_txt_017 %>
        <table width=100% cellspacing=1 cellpadding=1 bgcolor="#cccccc">
        <tr><td  align=right style="background-color:#FFFFFF; white-space:nowrap;">
		<b><%=formatnumber(timerforbrugt, 2)%> t.</b>
        </td>
        </tr>
                

        <%bruttoOmsReal = (OmsReal + matSalgsprisReal) %>

        <tr><td  align=right style="background-color:#F7f7f7; white-space:nowrap;">
        <b><%=formatnumber(bruttoOmsReal, 2) &" "& basisValISO_f8%> </b>
            <%if cint(oRec("jo_valuta")) <> cint(basisValId) then %>
                <br>&nbsp;
            <%end if%>

        </td>
        </tr>

         <tr><td  align=right style="background-color:#FFFFFF; white-space:nowrap;">
		<%=formatnumber(OmsReal, 2) &" "& basisValISO_f8%> 
		</td></tr>

        <%if cint(bdgmtypon_val) = 1 then 

         


            'strAktpajobKri = replace(strAktpajobKri, "taktivitetid = 0 OR", "")
             isGrpWritten = "#0#"
            for t = 1 to UBOUND(mtypgrpids)

        
            thisMtypTimerBelob(mtypgrpid(t)) = 0
            if len(trim(mtypgrpnavn(t))) <> 0 AND strAktpajobKri <> "taktivitetid = 0" then
           


           'Response.write "strAktpajobKri: " & strAktpajobKri & "<br>"

                                if cint(realfakpertot) <> 0 then 'timer o peridoe elle total == konsolideret

                                if instr(isGrpWritten, "#"&mtypgrpid(t)&"#") = 0 then
                                thisMtypTimerBelob(mtypgrpid(t)) = 0
                                end if

           
                                '** Timer konsolideret ligger altid p� den 1 p� en md. Derfor skal startdato altid v�re den 1. Slut dato vil atil v�re >= 1 S� det er ikke noget problem.
                               strSQLmtypbgt = "SELECT SUM(belob) AS mtyptimerbelob FROM timer_konsolideret_tot WHERE jobid = "& oRec("id") & " AND mtgid = "& mtypgrpid(t) & " AND dato BETWEEN '"& sqlDatostartMtypTkons &"' AND '"& sqlDatoslut &"' GROUP BY jobid"
                               else

                               strSQLmtypbgt = "SELECT SUM(belob) AS mtyptimerbelob FROM timer_konsolideret_tot WHERE jobid = "& oRec("id") & " AND mtgid = "& mtypgrpid(t) & " GROUP BY jobid"
                               end if


                                'response.write strSQLmtypbgt
                                'response.flush

                                oRec6.open strSQLmtypbgt, oConn, 3
                                if not oRec6.EOF then
                                
                                
                                if isNull(oRec6("mtyptimerbelob")) <> true then
                                thisMtypTimerBelob(mtypgrpid(t)) = thisMtypTimerBelob(mtypgrpid(t)) + formatnumber(oRec6("mtyptimerbelob"), 2)
                                else
                                thisMtypTimerBelob(mtypgrpid(t)) = thisMtypTimerBelob(mtypgrpid(t))
                                end if

                                end if
                                oRec6.close

                                 
                        isGrpWritten = isGrpWritten & ",#"& mtypgrpid(t) &"#"

            end if
            
         
             next


            for mtgp = 1 to UBOUND(kunMtypgrp) 'mtypgrpids 
            
                                      if len(trim(kunMtypgrpNavn(mtgp))) <> 0 then  %>

                                      <tr>
                                  
                                    <td  align=right style="background-color:#FFFFFF; white-space:nowrap; color:#999999;">
             
                                       

		                      
                                        <%=formatnumber(thisMtypTimerBelob(kunMtypgrp(mtgp)), 2) &" "& basisValISO_f8%>
		                            </td></tr> 
                                   <%
                                        end if
                        next

           
          
        end if

        
        %>
            <tr><td  align=right style="background-color:#F7F7F7; white-space:nowrap;">
		<%=formatnumber(tp, 2)&" "& basisValISO_f8%>/t.
		</td></tr>

        <tr><td  align=right style="background-color:#FFFFFF; white-space:nowrap;"><%=formatnumber(salgsOmkFaktisk,2) &" "& basisValISO_f8%></td></tr>

        
        

         <tr><td  align=right style="background-color:#F7F7f7; white-space:nowrap;">
		 <%=formatnumber(kostpris, 2) &" "& basisValISO_f8%> 
          
         <%if timerforbrugt <> 0 then 
         gnskostprtime = kostpris/timerforbrugt
         else
         gnskostprtime = 0
         end if%>
        </td></tr>
           <tr><td  align=right style="background-color:#FFFFFF; white-space:nowrap;">
         <%=formatnumber(gnskostprtime, 2) &" "& basisValISO_f8%>/t.
           </td>
        </tr>


         <%

         RealOmk = bruttoOmsReal - (salgsOmkFaktisk + kostpris)
         'dbreal = (totRealbel - realtimerkost) + (salgsomkostSalg - salgsomkostKost)
         realDbbelob = RealOmk
        
         call fn_dbproc(bruttoOmsReal,realDbbelob)
        realDB = dbProc


     
         %>

          <tr><td  align=right style="background-color:#FFFFFF;"><b><%=formatnumber(realDbbelob,0) & " "& basisValISO_f8 %></b>
      </td></tr>

         <tr><td  align=right style="background-color:#ffdfdf;">
         <%=left(tsa_txt_168, 4) %> DB: <b><%=formatnumber(realDB,0) %></b> %
       
      </td></tr>
      
      </table>

		</td>

        <%
		
        
		
		
        gnstimepris = 0
		if timerforbrugt <> 0 then
		gnstimepris = faktureretTimerEnhStk/timerforbrugt
		else
		gnstimepris = 0
		end if
		
		
		'faktottim = faktottim + (timerFak)


     
        
        udgifterFaktisk = salgsOmkFaktisk + kostpris
        db2bel = (faktureret-(udgifterFaktisk))

        call fn_dbproc(faktureret,db2bel)
        db2 = dbProc
       
        
        
        
        OmsWIP = (afsl_proc/100) * jobbudget
       
        
        salgsOmkWIP = salgsOmkFaktisk '(afsl_proc/100) * jo_udgifter_ulev
        nettoWIP = ((afsl_proc/100) * (jobbudget)) - salgsOmkWIP 'oRec("jobtpris") 
        

        wipGnskostprtime = gnskostprtime
        interKostWip = kostpris 'jo_udgifter_intern '(afsl_proc/100) *

       

        if timerforbrugt <> 0 AND nettoWIP <> 0 then
        wipTp = (nettoWIP/timerforbrugt)
        else
        wipTp = 0
        end if

        timerRealKvo = 1

        WipOmkIalt = interKostWip + salgsOmkWIP
        forvDbbel = (OmsWIP - (WipOmkIalt))

        call fn_dbproc(OmsWIP,forvDbbel)
        forvDb = dbProc

         %>

		
		
		<!-- WIP igangv�rende arbejde -->

		<td valign=top style="padding:6px 3px 3px 3px; white-space:nowrap; border-top:<%=btop%>px #cccccc solid; border-right:1px #cccccc solid;" >

        WIP (<%=formatnumber(afsl_proc,0) %>% af budget)
        <table width=100% cellspacing=1 cellpadding=1 bgcolor="#cccccc">
        <tr><td  align=right style="background-color:#FFFFFF; white-space:nowrap;"><%if instr(lto, "epi") = 0 then %>
        (forv.: <%=formatnumber(totalforbrugt, 0) %> t.) 
        <%end if %>
        <%=formatnumber(timerforbrugt, 0) %> t.</td></tr>
        <tr><td  align=right style="background-color:#F7f7f7; white-space:nowrap;">
		<b><%=formatnumber(OmsWIP, 2) &" "&basisValISO_f8%></b>

             <%if cint(oRec("jo_valuta")) <> cint(basisValId) then %>
                <br>&nbsp;
            <%end if%>

        </td></tr>
        <tr><td  align=right style="background-color:#FFFFFF;"><%=formatnumber(nettoWIP, 2) &" "& basisValISO_f8%></td></tr>



             <%if cint(bdgmtypon_val) = 1 then 'EPI ved budget og realtimer pr. fordeling p� medarbejdertyperGrupper 
            
                for mtgp = 1 to UBOUND(kunMtypgrp) 'mtypgrpids 
            
                                      if len(trim(kunMtypgrpNavn(mtgp))) <> 0 then  %>

                                    <tr>
                                        <td  style="background-color:#FFFFFF; color:#999999; height:17px;">&nbsp;</td>
                                    </tr> 
                                   <%
                                        end if
                next

            
            end if%>
         
         <tr><td  align=right style="background-color:#F7f7f7; white-space:nowrap;"><%=formatnumber(wipTp, 2) &" "& basisValISO_f8%>/t.</td></tr>
         <tr><td  align=right style="background-color:#FFFFFF; white-space:nowrap;"><%=formatnumber(salgsOmkWIP, 2) &" "& basisValISO_f8%></td></tr>

           <tr><td  align=right style="background-color:#F7f7f7; white-space:nowrap;">
         <%=formatnumber(interKostWip, 2) &" "& basisValISO_f8%>
         </td></tr>
         

         <tr><td  align=right style="background-color:#FFFFFF;"> <%=formatnumber(wipGnskostprtime, 2)&" "& basisValISO_f8%>/t.</td></tr>

            <tr><td  align=right style="background-color:#FFFFFF;"><b><%=formatnumber(forvDbbel,0) & " "& basisValISO_f8 %></b> 
      </td></tr>

          <tr><td  align=right style="background-color:#ffdfdf; white-space:nowrap;">WIP DB: 
          <%if lto = "xxx" then %>
          (Kvo: <%=formatnumber(timerRealKvo,2)&" x "&formatnumber(oRec("jo_dbproc"),0) %>)
          <%end if %> 
          <b><%=formatnumber(forvDb, 0) %></b> %
         </td></tr>



            <%balWIP = (faktureret - OmsWIP) %>
         <tr>
            
            <td align=right style="white-space:nowrap;">WIP / Fak.:
            <b><%=formatnumber(balWIP, 2)&" "& basisValISO_f8%></b></td>
        </tr>



        </table>

		</td>







		
        <!--- Faktureret --->

        <td valign=top style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc solid; white-space:nowrap; border-right:0px #cccccc solid;">

        <%=tsa_txt_155 %>
         <table width=100% cellspacing=1 cellpadding=1 bgcolor="#cccccc">
                <tr><td  align=right style="background-color:#FFFFFF;">&nbsp;</td></tr>
                <tr><td  align=right style="background-color:#F7f7f7; white-space:nowrap;">
		        <b><%=formatnumber(faktureret)&" "& basisValISO_f8%></b>

                     <%if cint(oRec("jo_valuta")) <> cint(basisValId) then %>
                        <br>&nbsp;
                    <%end if%>


                    </td></tr>
                    <tr><td  align=right style="background-color:#FFFFFF; color:#000000; white-space:nowrap;">(<%=formatnumber(faktureretTimerEnhStk,2) &" "& basisValISO_f8%>) *</td></tr>

              <%if cint(bdgmtypon_val) = 1 then 
            


                  for mtgp = 1 to UBOUND(kunMtypgrp) 'mtypgrpids 
            
                                      if len(trim(kunMtypgrpNavn(mtgp))) <> 0 then  %>

                                
                                    <tr>
                                          <td style="background-color: #FFFFFF; color: #999999; height:17px;">&nbsp;</td>
                                      </tr>
                                   <%
                                        end if
                        next

                end if%>


              <tr><td align=right style="background-color:#F7f7f7; white-space:nowrap;">
           <%=formatnumber(gnstimepris) &" "& basisValISO_f8%>/t.  </td></tr>

         <tr><td  align=right style="background-color:#FFFFFF; white-space:nowrap;"><%=formatnumber(salgsOmkFaktisk,2) &" "& basisValISO_f8%></td></tr>
         <tr><td  align=right style="background-color:#F7f7f7; white-space:nowrap;"><%=formatnumber(kostpris,2) &" "& basisValISO_f8%></td></tr>
           <tr><td  align=right style="background-color:#FFFFFF; white-space:nowrap;"><%=formatnumber(gnskostprtime, 2)&" "& basisValISO_f8%>/t.</td></tr>
           <tr><td  align=right style="background-color:#FFFFFF;"><b><%=formatnumber(db2bel,0) & " "& basisValISO_f8 %></b></td></tr>
         <tr><td  align=right style="background-color:#ffdfdf; white-space:nowrap;">
         Faktisk DB: <b><%=formatnumber(db2, 0)%></b> %
       </td></tr>


              <%
		
       '*** bal WIP ****'
        
     
        balTermin = (faktureret - totTerminBelobJob)
        
		tilfakturering = tilfakturering + (bal)
        tilfaktureringWIP = tilfaktureringWIP + (balWIP)
		%>


        <tr>
          
            <td align=right style="white-space:nowrap;">
            Termin / Fak.: <b><%=formatnumber(balTermin, 2)&" "& basisValISO_f8%></b></td>
        </tr>


       </table>
         
      
    
       
        </td>



         <%'** Terminer kolonne ***' %>
         <%if cint(visKol_terminer) = 1 then %>
         <td  style="border-top:<%=btop%>px #cccccc solid; white-space:nowrap; border-right:1px #cccccc solid; padding:6px 3px 3px 3px;" align="right" valign="top">
        

            <table cellpadding=2 cellspacing=0 border=0 width=100%>
            
            <%call terminbelob(oRec("id"), 1) %>
            
            <tr><td colspan=3 align=right>ialt: <b><%=formatnumber(totTerminBelobJob) %></b></td></tr>
            </table>

             <a href="javascript:popUp('milepale.asp?menu=job&func=opr&jid=<%=oRec("id")%>&rdir=wip&type=1','650','500','250','120');" target="_self" class=rmenu>+ Opret ny termin</a>
         </td>

        <%end if %>



       

       
      <%end select%>
        

      
        
        
        
        
        <%'**** job Tweet ***%>
        <%if cint(visKol_jobTweet) = 1 then %>

		<td valign=top colspan=2  style="padding:6px 3px 3px 3px; border-top:<%=btop%>px #cccccc solid; border-left:1px #cccccc solid; width:220px;">
		<%if print <> "j" then %>
        <input type="text" style="width:160px; color:#999999; font-style:italic; font-size:10px;" maxlength="60" value="Job tweet..(�ben for alle)" class="FM_job_komm form-control input-small" name="FM_job_komm_<%=oRec("id")%>" id="FM_job_komm_<%=c%>">
        <span class="aa_job_komm" id="aa_job_komm_<%=c%>" style="background-color:#9ACD32; border:0px #999999 solid; font-size:10px; padding:2px;">Gem >></span>
           
		
        <br />

        <%if len(trim(oRec("kommentar"))) <> 0 then
        showRyd = 1
         kommHgt = 100
            else
        showRyd = 0
             kommHgt = 20
            end if      
         %>

        <div id="dv_job_komm_<%=c%>" style="width:160px; font-size:9px; font-family:Arial; color:#000000; font-style:italic; overflow:auto; height:<%=kommHgtpx%>;"><%=oRec("kommentar")%></div>
        <%if cint(showRyd) = 1 AND (level = 1 OR cdbl(oRec("jobans1")) = cdbl(session("mid")) OR cdbl(oRec("jobans2")) = cdbl(session("mid"))) then %>
            <span class="aa_job_komm_ryd" id='FM_job_komm_ryd_<%=c%>' style="color:#FF0000; background-color:#FfC0CB;">Ryd >></span>   
            <%end if %> 
            <!--
            <input type="Text" id="FM_kommentar" name="FM_kommentar" value="<%=oRec("kommentar")%>">	
            <input id="FM_kommentar" name="FM_kommentar" value="#" type="hidden" />	
            -->
            <%else %>
            <%=oRec("kommentar") %>
            <%end if %>		

          

		</td>
        <%else %>
        
        <%end if %>




        <%
            if lto = "dencker" OR lto = "intranet - local" OR lto = "hidalgo" then %>
                <td valign="top" style="padding-left:10px; border-left:1px #CCCCCC solid;">
                    <%'Delprocesser %>
                    <table width="80%" class="table table-stribed" style="font-size:80%;">
                        <tr>
                            <th style="padding-left:0px;">Forretningsomr.</th>
                            <th style="padding-left:10px;">Forkalk.</th>
                            <th style="padding-left:10px;">Real.</th>
                           
                            <th style="padding-left:10px;" colspan="2">Stade</th>
                            <th style="padding-left:10px;">Aktiv</th>
                            <th style="padding-left:10px; white-space:nowrap;">Datoer trg.</th>
                        </tr>

                    <%
                      '*** Henter de 3 planl�gnings aktivieter 

                        'dim fomr_anavn, fomr_abudgettimer, fomr_aktpris, fomr_fomr, fomr_sumtimer
                        
                        strSQldelproc = "SELECT a.id AS aid, a.navn, a.aktstatus, aktbudgetsum, aktpris, budgettimer, fr1.for_fomr, fr2.for_aktid FROM aktiviteter a "_
                        &" LEFT JOIN fomr_rel fr1 ON (fr1.for_aktid = a.id) "_
                        &" LEFT JOIN fomr_rel fr2 ON (fr2.for_fomr = fr1.for_fomr AND fr2.for_jobid = a.job)  "_
                        &" WHERE fase = 'planl�g' AND fakturerbar = 91 AND a.job = "& oRec("id") &" GROUP BY fr2.for_aktid ORDER BY a.sortorder, fr1.for_fomr"
                        
                        'if session("mid") = 21 then
                        'Response.Write "strSQldelproc: " & strSQldelproc
                        'Response.flush
                        'end if

                        '3-5 stk.

                        budgetTimerTotpl = 0
                        sumTimerTotpl = 0
                        fot = 0
                        lastFomr = 0
                        statusStr = ""
                        cntf = 0
                        oRec9.open strSQldelproc, oConn, 3 
                        while not oRec9.EOF

                        if lastFomr <> oRec9("for_fomr") AND cntf > 0 then 
                        fot = fot + 1
                        end if

                        fomr_fomr(tot) = oRec9("for_fomr")
                        fomr_aid(fot) = oRec9("aid")
                        fomr_anavn(fot) = oRec9("navn")

                        

                        '*** Hvis bare en er aktiv er den gr�n. 
                        'statusStr = statusStr & ",#"& oRec9("aktstatus") &"#"
                        form_aktstatus(fot) = form_aktstatus(fot) & ",#"& oRec9("aktstatus") &"#"


                        strSQLfomraktbudget = "SELECT budgettimer, aktpris FROM aktiviteter WHERE id =  "& oRec9("for_aktid") 
                        oRec10.open strSQLfomraktbudget, oConn, 3 
                        if not oRec10.EOF then

                        fomr_abudgettimer(fot) = fomr_abudgettimer(fot)*1 + formatnumber(oRec10("budgettimer"), 2)*1
                        fomr_aktpris(fot) = fomr_aktpris(fot)*1 + formatnumber(oRec10("aktpris"), 2)*1

                        end if
                        oRec10.close

                        if realfakpertot <> 0 then 'total eller periode
                        peride_totalSQLKRI = " tdato BETWEEN '"& sqlDatostart &"' AND '"& sqlDatoslut &"' AND"
                        else
                        peride_totalSQLKRI = ""
                        end if

                        strSQLfomrTimer = "SELECT SUM(timer) AS fomr_timer, tdato FROM timer WHERE "& peride_totalSQLKRI &" taktivitetid =  "& oRec9("for_aktid") & " GROUP BY taktivitetid ORDER BY tdato DESC"
                        
                        'if session("mid") = 21 then
                        'Response.write strSQLfomrTimer
                        'Response.flush
                        'end if
                        
                        oRec10.open strSQLfomrTimer, oConn, 3 
                        if not oRec10.EOF then

                        fomr_sumtimer(fot) = (fomr_sumtimer(fot)*1) + (oRec10("fomr_timer")*1)
                        form_lasttregdato(fot) = oRec10("tdato")

                        end if
                        oRec10.close

                        strSQLfomrTimer = "SELECT tdato FROM timer WHERE taktivitetid =  "& oRec9("for_aktid") & " ORDER BY tdato"
                        oRec10.open strSQLfomrTimer, oConn, 3 
                        if not oRec10.EOF then

                        form_firstregdato(fot) = oRec10("tdato")

                        end if
                        oRec10.close

                        'fomr_fomr(fot)
                        
                        %>

                        <!--
                        <tr>
                            <td><%=oRec9("navn") %></td>
                            <td><%=formatnumber(oRec9("budgettimer"), 2) %> t.</td>
                            <td><%=formatnumber(oRec9("aktpris"), 2) %> DKK.</td>

                            <td></td>

                        </tr>
                        -->
                        <%

                       
                        cntf = cntf + 1
                        lastFomr = oRec9("for_fomr")
                        oRec9.movenext
                        wend
                        oRec9.close
                            
                            
                        if fot > 0 then


                        for fot = 0 TO UBOUND(fomr_aid)
                            
                            if len(trim(fomr_aid(fot))) <> 0 then
                             %>

                      
                        <tr>
                            <td style="padding-left:0px; white-space:nowrap;"><%=left(fomr_anavn(fot), 20) %></td>


                            <%if fomr_anavn(fot) <> "Ubemandet" then %>
                            <td style="padding-left:10px white-space:nowrap;" align="right"><%=formatnumber(fomr_abudgettimer(fot), 2) %> t.</td>
                            <td style="padding-left:10px; white-space:nowrap;" align="right"><%=formatnumber(fomr_sumtimer(fot), 2) %> t.</td>
                            <%else%>
                            <td style="padding-left:10px white-space:nowrap;" align="right">(<%=formatnumber(fomr_abudgettimer(fot), 2) %> t.)</td>
                            <td style="padding-left:10px; white-space:nowrap;" align="right">(<%=formatnumber(fomr_sumtimer(fot), 2) %> t.)</td>
                            <%end if

                                if fomr_anavn(fot) <> "Ubemandet" then
                                budgetTimerTotpl = budgetTimerTotpl + fomr_abudgettimer(fot)
                                sumTimerTotpl = sumTimerTotpl + fomr_sumtimer(fot)
                                end if

                                if fomr_anavn(fot) = "Ubemandet" then
                                budgetTimerTotplinklube = budgetTimerTotpl + fomr_abudgettimer(fot)
                                sumTimerTotplinklube = sumTimerTotpl + fomr_sumtimer(fot)
                                end if

                                fomr_procent = 0
                                if fomr_abudgettimer(fot) <> 0 then
                                fomr_procent = 100 * (fomr_sumtimer(fot)/fomr_abudgettimer(fot)) 
                                else
                                fomr_procent = 100
                                end if
                                
                                if fomr_sumtimer(fot) <> 0 then

                                    

                                if fomr_procent >= 100 then
                                form_bgCol = "red"
                                ill = "dot_rod.gif"
                                else
                                form_bgCol = "yellowgreen"
                                ill = "dot_gron.gif"
                                end if

                                else
                                form_bgCol = "#999999"
                                fomr_procent = 0
                                ill = "dot_graae.gif"
                                end if

                                '*** Nulstiller **'
                                fomr_abudgettimer(fot) = 0
                                fomr_sumtimer(fot) = 0
                                %>

                            <td style="padding-left:10px;">
                                 <div style="background-color:<%=form_bgCol%>; width:15px; height:15px;">
                                 <img src='../ill/<%=ill %>' width='15' height='15' alt='' border='0'></div>
                            </td>
                            <td style="padding-left:10px; white-space:nowrap;"><%=formatnumber(fomr_procent, 0) %> %</td>

                            <td style="padding-left:10px;">
                                <%
                                 form_StatusbgCol = ""  
                                    
                                 if instr(form_aktstatus(fot), ",#1#") <> 0 then 'Der er aktive
                                 form_StatusbgCol = "yellowgreen"
                                 ill = "dot_gron.gif"
                                    else
                                        if fomr_sumtimer(fot) <> 0 then 'Der INGEN aktive og ingen timer = HVID eller R�D
                                        form_StatusbgCol = "red"
                                        ill = "dot_rod.gif"
                                        else
                                        form_StatusbgCol = ""
                                        ill = "blank.gif"
                                        end if
                                 end if

                                 %>
                               
                                <div style="background-color:<%=form_StatusbgCol%>; width:15px; height:15px;">
                                    <img src='../ill/<%=ill %>' width='15' height='15' alt='' border='0'>
                                </div>
                            </td>
                            <td style="padding-left:10px; white-space:nowrap;">
                                <span style="font-size:8px;"><%=formatdatetime(form_firstregdato(fot), 2) %></span>
                                <span style="font-size:8px; color:red;"><%=formatdatetime(form_lasttregdato(fot), 2) %></span></td>

                        </tr>
                       
                        <% end if

                        next

                        end if
                            %>
                        <tr><td>Total:<br />
                            Inkl. ubem.:
                            </td>
                            <td style="white-space:nowrap;" align="right"><%=formatnumber(budgetTimerTotpl, 2)%> t.<br />
                                <%=formatnumber(budgetTimerTotplinklube, 2)%> t.
                            </td>
                            <td style="white-space:nowrap;" align="right"><%=formatnumber(sumTimerTotpl, 2)%> t.<br />
                                <%=formatnumber(sumTimerTotplinklube, 2)%> t.
                            </td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                        </tr>
                       

                    </table>


                    <%if print <> "j" then %>
                    <div class="form-inline">
                    Budget:
                    <input id="FM_forkalk_<%=c %>" type="text" class="s_forkalk form-control input-small" value="<%=forkalkvalue %>" style="width:65px; text-align:right;"/>&nbsp;t.
                    <input id="FM_brutoms_<%=c %>" type="text" class="s_brutoms form-control input-small" value="<%=bruttooms %>" style="width:70px; text-align:right;"/>
           
                      <%
                                     felt = "FM_jo_valuta_"& c
                                     call valutaList(jo_valuta, felt)
                                     %>
       
                        <span id="sp_forkalk_<%=c %>" style="color:green; font-size:12px; visibility:hidden;" ><i>V</i></span>
                       
                    </div>
                    <%else 
                        
                        call valutakode_fn(jo_valuta)
                        %>
                    Forkalk.: <%=formatnumber(forkalkvalue, 2) & " t." %><br />
                    Budget: <%=formatnumber(bruttooms, 2) &" "& valutaKode_CCC %>

                    <%end if %>
                </td>
                <%end if %>



















         <%if cint(visKol_planlagning) = 1 then 
             
          ddplan = strAar &"-"& strMrd & "-"& strDag%>
          <input type="hidden" id="ddplan_0" value="<%=ddplan %>" />
        
         <td style="border-left:1px #CCCCCC solid;">
             <%'Planl�gning %>
             <div style="overflow-x:scroll; width:650px;">
             <table class="table table-stribed" style="font-size:80%;">
                  
                 
                 <tr>
                     <th><%=monthname(datepart("m", ddplan, 2,2)) %></th>
                     <%for w = 0 to 28 
                         
                         wrtPlanDato = dateadd("d", w, ddplan)
                         
                         thisWeekday = datepart("w", wrtPlanDato, 2,2)
                         select case thisWeekday
                         case 6,7
                         bgcolorWD = "#999999"
                         wrtDato = ""
                         case 3
                         bgcolorWD = "lightblue"
                         call thisWeekNo53_fn(wrtPlanDato)
                         thisWeekNo = thisWeekNo53 'datepart("ww", wrtPlanDato, 2,2)
                         wrtDato = thisWeekNo 

                         if wrtDato < 10 then
                         wrtDato = "0"& wrtDato
                         end if

                         case else
                         bgcolorWD = ""
                         wrtDato = left(formatdatetime(wrtPlanDato, 2), 2)
                         end select
                        
                         %>
                        <th style="background-color:<%=bgcolorWD%>; border-right:1px #cccccc solid;"><%=wrtDato %></th>
                     <%next %>
                     <th><!----></th>
                 </tr>


                 <%
                 apllopp = 0     
                 strSQLakt = "SELECT id As aid, navn as aktnavn, "_
                 &" projektgruppe1, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, "_
                 &" projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter WHERE job = "& oRec("id") & " AND fakturerbar = 91 ORDER BY sortorder"
                 oRec4.open strSQLakt, oConn, 3
                 while not oRec4.EOF 


                                     aktbookingdates = ""
                                     strSQLaktplanlagt = "SELECT ab_startdate FROM akt_bookings WHERE ab_jobid = "& oRec("id") & " AND ab_aktid = "& oRec4("aid") &""
                                     oRec5.open strSQLaktplanlagt, oConn, 3
                                     while not oRec5.EOF 

                                     aktbookingdates = aktbookingdates & ",#"& formatdatetime(oRec5("ab_startdate"), 2) &"#"

                                     oRec5.movenext
                                     wend
                                     oRec5.close

                        
                                    
                               

                     %>
                 <input type="hidden" id="planlaeg_jobid_<%=c%>" value="<%=oRec("id") %>" />
                 <input type="hidden" id="planlaeg_aktid_<%=oRec("id") %>_<%=apllopp%>" value="<%=oRec4("aid") %>" />
                     <tr>
                     <td style="white-space:nowrap">
                         <a href="http://timeout.cloud/timeout_xp/wwwroot/ver4_22/ressource_planner/ressplan_2017.aspx?FM_job=<%=oRec("id")%>&FM_projectgroup=<%=oRec4("projektgruppe1")%>&FM_lto=<%=lto%>" target="_blank"><%=left(oRec4("aktnavn"), 20) %></a>

                         <!--<br />planlaeg_jobid_<=apllopp%>-->
                        
                     </td>
                     <%for w = 0 to 28 
                         
                         

                         select case lcase(left(oRec4("aktnavn"), 4))
                         case "kons"
                         bgcolor = "#003399"
                         aktclass = "a0"
                         case "frem"
                         bgcolor = "yellowgreen"
                         aktclass = "a1"
                         case "st�b"
                         bgcolor = "yellow"
                         aktclass = "a2"
                         case "elek"
                         bgcolor = "lightblue"
                         aktclass = "a3"
                         case else
                         bgcolor = "lightpink"
                         aktclass = "a4"
                         end select



                         wrtPlanDatoThis = dateadd("d", w, ddplan)
                         if instr(aktbookingdates, "#"& cDate(wrtPlanDatoThis) &"#") <> 0 then
                         chk = "CHECKED"
                         clsThis = "chk"
                         else
                         chk = ""
                         clsThis = ""
                         end if

                         if chk = "CHECKED" then

                         

                         else

                         'bgcolor = "#FFFFFF"
                             thisWeekday = datepart("w", wrtPlanDatoThis, 2,2)
                             select case thisWeekday
                             case 6,7
                             bgcolor = "#999999"
                             case else
                             bgcolor = "#FFFFFF"
                             end select

                         end if

                       
                        %>

                        <td class="<%=aktclass%> planlaeg <%=clsThis%>" id="pl_<%=apllopp%>_<%=w%>_<%=c%>" style="border-right:1px #cccccc solid; background-color:<%=bgcolor%>;">
                            <%if print = "j" AND clsThis = "chk" then %>
                            X
                            <%else %>
                            &nbsp;
                            <%end if %>
                             




                        </td>
                        
                     <%next %>
                     <td>&nbsp;</td>
                 </tr>

                     <%

                 apllopp = apllopp + 1
                 oRec4.movenext
                 wend
                 oRec4.close%>

                
                 

                 </table>
                 </div>

         </td>
          <%end if%>


       
		</tr>
	
	<%

    else

    call terminbelob(oRec("id"), 0)

    OmsWIP = (afsl_proc/100) * jobbudget
      
   

    end if 'visKunGT = 1
    








    OmsRealTot = OmsRealTot + OmsReal
	
	call valutaKurs_fakhist(oRec("jo_valuta"))
    call beregnValuta(oRec("jo_bruttooms"),dblKurs_fakhist,100)
    budgetIalt = (budgetIalt*1) + (valBelobBeregnet) '/100 'oRec("jo_bruttooms") 'oRec("jobtpris")

	'budgettimerIalt = budgettimerIalt + (oRec("budgettimer") + oRec("ikkebudgettimer"))
	
     salgsOmkFaktiskTot = salgsOmkFaktiskTot + salgsOmkFaktisk/1

	if len(oRec("restimer")) <> 0 then
	restimerIalt = restimerIalt + oRec("restimer")
	else
	restimerIalt = restimerIalt
	end if

    OmsWIPtot = OmsWIPtot + OmsWIP  

    faktotbel = faktotbel + (faktureret)

    

	
	'** Til Export fil ***'
	jids = jids & "," & oRec("id")
    c = c + 1
    end if 'Realtimer (nulfilter)
	
	Response.flush
	
	
	oRec.movenext
	wend
	oRec.close 


        next 'p
	
	%>
	
	 <input id="FM_kommentar" name="FM_kommentar" value="xc" type="hidden" />	
	 <input id="Hidden5" name="FM_kommentar_ny" value="xc" type="hidden" />	
	
    <% if cint(sorter) = 7 OR cint(sorter) = 8 then 
        
     else%>

	<tr bgcolor="#ffdfdf">
	<td style="border-top:1px #cccccc solid;">Grandtotal</td>
    <%if sorter <> 3 AND sorter <> 31 then %>
     <td align=right style="border-top:1px #cccccc solid; border-right:0px #cccccc solid;">&nbsp;</td>
    <%end if %>

        <td align=right  style="border-top:1px #cccccc solid; border-right:0px #cccccc solid;">&nbsp;</td>

        <%
            
         


                if cint(visKol_rekvnr) = 1 then
                %>
                <td align=right  style="border-top:1px #cccccc solid; border-right:1px #cccccc solid;">&nbsp;</td>
                <%
                end if

            
            
         select case visSimpel
         case 0  %>
        <td align=right style="border-top:1px #cccccc solid; border-right:0px #cccccc solid;">&nbsp;</td>

        <% if cint(visKol_prioritet) = 1 then %>
        <td align=right style="border-top:1px #cccccc solid; border-right:0px #cccccc solid;">&nbsp;</td>
        <%end if 
            
        case 1  %>
       

        <% if cint(visKol_prioritet) = 1 then %>
        <td align=right style="border-top:1px #cccccc solid; border-right:0px #cccccc solid;">&nbsp;</td>
        <%end if %>
           
        
       <%case 2 %>
          <!--<td align=right style="border-top:1px #cccccc solid; border-right:0px #cccccc solid;">&nbsp;</td> -->
          <!--<td align=right colspan="3"  style="border-top:1px #cccccc solid;">&nbsp;</td>-->
        <%end select %>

      

        <% if cint(visKol_forecast) = 1 then %>
	    <td valign=top align=right style="border-top:1px #999999 solid; border-right:1px #cccccc solid; padding:6px 3px 3px 3px;" nowrap>Forecast:<br /><b><%=formatnumber(restimerIalt, 2)%> t.</b> </td>
        <%end if %>

    
           <%select case visSimpel
         case 0,1
        
            case 2 %>
        <td colspan="3">&nbsp;</td>
    <td align=right  valign=top style="border-top:1px #cccccc solid; white-space:nowrap; border-right:1px #cccccc solid; padding:6px 3px 3px 3px;">Budget:<br />
    <b><%=formatnumber(budgettimerIalt, 2)%> t.</b><br />
    <b><%=formatnumber(budgetIalt, 2)&" "& basisValISO_f8%> </b><br />
	<br />Salgsomk.:<br /> <%=formatnumber(udgifterTot, 2) &" "& basisValISO_f8%></td>
	<%end select %>


	<td valign=top align=right  style="border-top:1px #cccccc solid; border-right:1px #cccccc solid; white-space:nowrap; padding:6px 3px 3px 3px;">Real:<br /><b><%=formatnumber(timerforbrugtIalt, 2) %> t.</b><br />
    <b><%=formatnumber(OmsRealTot, 2) &" "& basisValISO_f8%></b>
    
          <%select case visSimpel
         case 0,1
            case 2 %>    
    <br />
    <br />Salgsomk.:<br /> <%=formatnumber(salgsOmkFaktiskTot,2) &" "& basisValISO_f8%>
        <%end select %>
	</td>

   
          


          <%select case visSimpel
         case 0
         case 1%>
         <td align=right colspan="2"  style="border-top:1px #cccccc solid; border-right:1px #cccccc solid;">&nbsp;</td>
         <%case 2%>
         <td align=right  valign=top style="border-top:1px #cccccc solid; white-space:nowrap; border-right:1px #cccccc solid; padding:6px 3px 3px 3px;">
         WIP:<br /><b><%=formatnumber(OmsWIPtot, 2) &" "& basisValISO_f8%></b>
         </td>
         <td align=right valign=top  style="border-top:1px #cccccc solid; white-space:nowrap; padding:6px 3px 3px 3px; border-right:1px #cccccc solid;">
             Faktureret:<br />
             <b><%=formatnumber(faktotbel, 2)&" "& basisValISO_f8 %></b>


            <%'if visKunGT <> 1 then %>
            <br /><br />Igangv�rende arb.:<br />
	        <b><%=formatnumber(tilfakturering, 2) %></b><br />
            <%=formatnumber(tilfaktureringWIP, 2) %>
            
             


         </td>
   

            <%if cint(visKol_terminer) = 1 then %>
            <td align=right valign=top  style="border-top:1px #cccccc solid; white-space:nowrap; padding:6px 3px 3px 3px;">
	        <b><%=formatnumber(totTerminBelobGrand) &" "& basisValISO_f8%></b> 
	        </td>
            <%end if %>

	        
	
        <%end select %>


          <%select case visSimpel
         case 0,1
         %>
        <td align=right  valign=top style="border-top:1px #cccccc solid; white-space:nowrap; border-right:1px #cccccc solid; padding:6px 3px 3px 3px;">Budget:<br />
        <b><%=formatnumber(budgettimerIalt, 2)%> t.</b><br />
        <b><%=formatnumber(budgetIalt, 2)&" "& basisValISO_f8%>  </b>
        </td><%
        case 2
	    end select %>

   
          <%if cint(visKol_jobTweet) = 1 then %>
	    <td align=right style="border-top:1px #cccccc solid;padding:6px 3px 3px 3px;" colspan=2>&nbsp;</td>
        <%end if %>
        
        <%if (lto = "dencker" OR lto = "intranet - local" OR lto = "hidalgo") AND visSimpel = 2 then %>
	    <td>&nbsp;</td>
        <%end if %>

        <%if cint(visKol_planlagning) = 1 then %>
        <td>&nbsp;</td>
        <%end if %>
	</tr>
	<%end if %>


	
	<%if print <> "j" then %>
        <!--
	<tr><td colspan="<%=dtColSPan %>" align=right style="padding:20px 10px 0px 0px;"><input type="submit" value="Opdater liste >>"></td></tr>
        -->
	<%end if %>
	
	
	  </form>
	</table>
  
    Der er <b><%=c %></b> job p� listen i denne s�gning.
	
	<!--</div>-->
                      <!-- section -->
	
	
	
	
	<%if print <> "j" then
	
	itop = 100
	ileft = 0
	iwdt = 400
	
	call sideinfo(itop,ileft,iwdt)
	
	%>
	<!--
	<b>Skjulte job</b><br />
	s�t prioitet = -1
    <br /><br />
    <b>Sortering:</b><br>Tr�k i et job for at prioitere r�kkef�lgen. <br>Prioitets-angivelsen skifter f�rst ved gen-indl�s.<br /><br />

    <%if lto <> "epi" AND lto <> "epi_no" AND lto <> "epi_sta" AND lto <> "epi_ab" AND lto <> "intranet - local" then %>
    <b>Kvotient</b><br />
    Kvotient beregnes udfra restestimat * (timer budgetteret / forventet timeforbrug)<br /><br />
    <%end if %>

    *)  Faktureret ekskl. materialeforbrug (salgsomkostninger) og km.<br /> 
     -->
	</td>
	</tr>
	</table>
	</div>
	







                  <%'*** Eksport og print%>
                  <br /><br /><br /><br />
                    <div class="well">
                      <table>


	           
                <%

                 'ptop = 0 '57
                'pleft = 935
                'pwdt = 180

                'call eksportogprint(ptop,pleft, pwdt)



                if visSimpel = 2 then
                eksDataNrlCHK = "CHECKED"

                    select case lto 
                    case "epi", "epi2017", "epi_2017"
                    eksDataNr3CHK = "CHECKED"
                    case else
                    eksDataNr3CHK = ""
                    end select

                else
                eksDataNrlCHK = ""
                eksDataNr3CHK = ""
                end if
                %>
                
                
                    <form action="../timereg/job_eksport.asp" method="post" target="_blank">
                    <input id="jids" name="jids" value="<%=jids%>" type="hidden" />
                    <input id="Hidden1" name="realfakpertot" value="<%=realfakpertot%>" type="hidden" />
                    <input id="Hidden7" name="historisk_wip" value="<%=historisk_wip%>" type="hidden" /> 
                    <input id="Hidden9" name="timertilfak" value="<%=timertilfak%>" type="hidden" />     
                           
                    <input id="Hidden3" name="sqlDatostart" value="<%=sqlDatostart%>" type="hidden" />
                    <input id="Hidden4" name="sqlDatoslut" value="<%=sqlDatoslut%>" type="hidden" />

                        
                        
                    <input id="Hidden6" name="visSimpel" value="<%=visSimpel%>" type="hidden" />

                     
                        <input name="visKol_rekvnr" value="<%=visKol_rekvnr%>" type="hidden" />
                        <input name="visKol_forecast" value="<%=visKol_forecast%>" type="hidden" />
                        <input name="visKol_planlagning" value="<%=visKol_planlagning%>" type="hidden" />
                        <input name="visKol_jobTweet" value="<%=visKol_jobTweet%>" type="hidden" />
                         <input name="visKol_prioritet" value="<%=visKol_prioritet%>" type="hidden" />

                        

                        <input id="Hidden6" name="jobansv" value="<%=jobansv%>" type="hidden" />
                        <input id="Hidden6" name="salgsansv" value="<%=salgsansv%>" type="hidden" />
                        <input id="Hidden6" name="kansv" value="<%=kansv%>" type="hidden" />
                        
                        
                    

                     <tr>
                        <td align=center valign="top">&nbsp;</td>
                        <td>
                            
                            <!--<input id="Submit1" type="submit" value=".csv fil eksport" style="font-size:9px; width:90px;"/><br />-->

                         
                        
                           <input type="checkbox" value="1" name="xeksDataStd" checked disabled /> Stamdata<br />
                          <input type="checkbox" value="1" name="eksDataNrl" <%=eksDataNrlCHK %> /> N�gletal (Salgsomk mm. hvis udviddet valgt)<br />
                            <input type="checkbox" value="1" name="eksDataNr3" <%=eksDataNr3CHK %> /> WIP (Kun hvis n�gletal og udviddet valgt)<br />
                           <input type="checkbox" value="1" name="eksDataNr2" /> Forr.omr., Projektgrupper<br />
                        <input type="checkbox" value="1" name="eksDataMile" /> Milep�le/Terminer 
                        
                        <%
                        call stadeOn()
                        if jobasnvigv = 1 then %>
                        (stade)
                        <%end if %>
                        
                        <br />
                        <input type="checkbox" value="1" name="eksDataJsv" />  <%
                        call salgsans_fn()    
                        if cint(showSalgsAnv) = 1 then  %>
                        Job- og salgs -ansvarlige
                        <%else %>
                        Jobansvarlige
                        <%end if %>

                              <br /><br /><input id="Submitcsv" type="submit" value=".csv fil eksport" class="btn btn-small btn-secondary" />

                         </td>
                   </tr>
                   </form>
                    


                    <form action="webblik_joblisten.asp?print=j" method="post" target="_blank">
                    <input id="Hidden2" name="realfakpertot" value="<%=realfakpertot%>" type="hidden" />
                    <input id="Hidden8" name="historisk_wip" value="<%=historisk_wip%>" type="hidden" />
                    <input id="Hidden10" name="timertilfak" value="<%=historisk_wip%>" type="hidden" />
                     <input id="Hidden6" name="jobansv" value="<%=jobansv%>" type="hidden" />
                        <input id="Hidden6" name="salgsansv" value="<%=salgsansv%>" type="hidden" />
                        <input id="Hidden6" name="kansv" value="<%=kansv%>" type="hidden" />
                        
                        
                    <tr>
                        <td align=center>&nbsp;</td>
                        <td><br /><br /><br /><input id="Submit2" type="submit" value="Printvenlig" class="btn btn-small btn-info" /></td>
                   </tr>



                  <!--      	
	              <tr><td colspan="2"><br /><br /><%
                  LnkUrl = "jobs.asp?menu=job&func=opret&id=0&int=1&rdir=webblik" 
                  lnkTxt = "Opret nyt job"
                  lnkAlt = lnkTxt
                  lnkWdt = 100
                  lnkTgt = "_top"
                  call opretLink_2013(lnkUrl, lnkTxt, lnkAlt, lnkTgt, lnkWdt) %>
                    
                    </td></tr>
                        -->
	            
	         

                   </form>
                    </table>
                      </div>
                   
                   
           
                <%else%>

                <% 
                Response.Write("<script language=""JavaScript"">window.print();</script>")
                %>
                <%end if%>
	
	
	
	    <br /><br />
            &nbsp;
	
    	    </div>























            
                 </div><!-- Portlet body--> 
                  </div>
              </div><!-- wrapper -->
          </div><!-- content --->
        </div>




        <%else %>
        <%end if 'fromvmenu %>
	
	
	<%end if%>


<!--#include file="../inc/regular/footer_inc.asp"-->

