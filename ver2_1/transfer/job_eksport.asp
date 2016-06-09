<%Response.buffer = true%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="inc/convertDate.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<%
if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
	
	
	filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
	
	if len(request("optiprint")) <> 0 then
	optiPrint = 1	
	else
	optiPrint = 0
	end if
	
	
	if optiPrint = 0 then
	ext = "csv"
	else
	ext = "txt"
	end if

    if len(trim(request("realfakpertot"))) <> 0 then
    realfakpertot = realfakpertot
    else
    realfakpertot = 0
    end if

    call TimeOutVersion()

    call akttyper2009(2)
	
	Set objFSO = server.createobject("Scripting.FileSystemObject")
	
	if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\job_eksport.asp" then
							
		Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&"", True, False)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&"", 8)
	
	else
		
		Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&"", True, false)
		Set objNewFile = nothing
		Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\log\data\jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&"", 8, -1)
		
	end if
	
	file = "jobexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&"."&ext&""
	
	
	
	if optiPrint = 0 then
	strTxtExport = strTxtExport & "Jobnavn;Jobnr.;Status;Kontakt;Kontakt id;"_
	& "Startdato;Slutdato;Fastpris:1/Lbn Timer:0;Forkalk. timer;Brutto Oms.;Salgsomk.; Intern kost.;DB 1 % (Budget);Timer Realiseret;~Timepris;Real. Omsætning;"_
    &"Forv. DB %;Faktureret;Salgsomk. (faktisk);Intern Kost. (faktisk);DB 2 % (faktisk);"_
    &"Timepris (faktisk);Stade %;Oms. WIP; Bal. Faktureret-Oms. WIP;"_
	&"Jobansvarlig 1;Init;Jobejer 2;Init;Jobmedansv. 1;Init;Jobmedansv. 2;Init;Jobmedansv. 3;Init;Aftale;Rekvnr.;Forretningsomr.;Beskrivelse;Kommentar" & vbcrlf 
	end if
	
	
	jids = request("jids")
	jobIder = split(jids, ",") 
	for i = 0 to Ubound(jobIder)
	
	
	strSQL = "SELECT id, jobnavn, jobnr, kkundenavn, kid, jobknr, jobTpris, jobstatus, jobstartdato, "_
	&" jobslutdato, budgettimer, fakturerbart, Kkundenr, ikkebudgettimer, jobans1, jobans2, jobans3, jobans4, jobans5, j.beskrivelse, fastpris, "_
	&" m1.mnavn AS m1navn, m1.init AS m1init, m1.mnr AS m1nr, "_
	&" m2.mnavn AS m2navn, m2.init AS m2init, m2.mnr AS m2nr, "_
    &" m3.mnavn AS m3navn, m3.init AS m3init, m3.mnr AS m3nr, "_
    &" m4.mnavn AS m4navn, m4.init AS m4init, m4.mnr AS m4nr, "_
    &" m5.mnavn AS m5navn, m5.init AS m5init, m5.mnr AS m5nr, "_
    &" rekvnr, serviceaft, kommentar, jo_dbproc, udgifter, jo_bruttooms, jo_udgifter_intern, jo_udgifter_ulev, restestimat, stade_tim_proc "_
	&" FROM job j, kunder "_
	&" LEFT JOIN medarbejdere m1 ON (m1.mid = jobans1) "_
	&" LEFT JOIN medarbejdere m2 ON (m2.mid = jobans2) "_
    &" LEFT JOIN medarbejdere m3 ON (m3.mid = jobans3) "_
    &" LEFT JOIN medarbejdere m4 ON (m4.mid = jobans4) "_
    &" LEFT JOIN medarbejdere m5 ON (m5.mid = jobans5) "_
	&" WHERE id = "& jobIder(i) &" AND kunder.Kid = jobknr ORDER BY jobnavn"
	
	'Response.write strSQL & "<hr>"
	'Response.flush
	
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
	
	
	strJobnavn = oRec("jobnavn")
	strJobnr = oRec("jobnr")
	intJobstatus = oRec("jobstatus")
	strKunde = oRec("kkundenavn")
	strKnr = oRec("kkundenr")
	dtStdato = oRec("jobstartdato")
	dtSldato = oRec("jobslutdato")
	intFakbart = oRec("fakturerbart")
	intFakbaretimer = oRec("budgettimer")
	intIkkeFakbaretimer = oRec("ikkebudgettimer")
	timerialt = (intFakbaretimer + intIkkeFakbaretimer)
    budgettimerIalt = timerialt
	strJobans1 = oRec("m1navn")
	strJobans1nr = oRec("m1nr")
	strJobans1init = oRec("m1init")
	
    strJobans2 = oRec("m2navn")
	strJobans2nr = oRec("m2nr")
	strJobans2init = oRec("m2init")

    strJobans3 = oRec("m3navn")
	strJobans3nr = oRec("m3nr")
	strJobans3init = oRec("m3init")

    strJobans4 = oRec("m4navn")
	strJobans4nr = oRec("m4nr")
	strJobans4init = oRec("m4init")

    strJobans5 = oRec("m5navn")
	strJobans5nr = oRec("m5nr")
	strJobans5init = oRec("m5init")

	strJobBesk = oRec("beskrivelse")
	dblBudget = oRec("jo_bruttooms") 'oRec("jobTpris")
	intFaspris = oRec("fastpris")
	rekvnr = oRec("rekvnr")
	serviceaft = oRec("serviceaft")
	komm = oRec("kommentar")
	jo_dbproc = oRec("jo_dbproc")
	
    'udgifter = oRec("udgifter") 'salgsomk + internkost
	salgsomk = oRec("jo_udgifter_ulev")
    jo_udgifter_ulev = oRec("jo_udgifter_ulev")
    internKost = oRec("jo_udgifter_intern")

		'*** Export ****
		'objF.WriteLine("Id "&  chr(009) &" Kundenavn "&  chr(009) &"Kundenr "& chr(009) &"Jobnavn "& chr(009)&"Aktivitet "& chr(009) &"Kunde "& chr(009) &"Timer / Enheder"& chr(009) &"Heraf fakturerbare timer"& chr(009) &"Reg. Timepris (uanset fastpris/lbn.tim.)"& chr(009) &"Gns.fak. Timepris" & chr(009) & "Medarbejder "& chr(009) &"Kommentar ")
		
		select case intJobstatus
		case 0
		strStatus = "Lukket"
		case 1
		strStatus = "Aktiv"
		case 2
		strStatus = "Passiv"
		case 3
		strStatus = "Tilbud"
		end select 
		
		
		if optiPrint = 0 then

        if dblBudget <> 0 then
        jobbudget = dblBudget
        else
        jobbudget = 0
        end if


        call timeRealOms '** cls_timer

       
        call stat_faktureret_job(oRec("id"), sqlDatostart, sqlDatoslut) '** cls_fak


       
		
		
		if timerforbrugt <> 0 then
		gnstimepris = faktureretTimerEnhStk/timerforbrugt
		else
		gnstimepris = 0
		end if


       
       
        udgifterFaktisk = salgsOmkFaktisk + kostpris

        if faktureret <> 0 then
            if faktureret > udgifterFaktisk then 
            db2 = 100 - ((udgifterFaktisk/faktureret) * 100)
            else
                if udgifterFaktisk <> 0 then
                db2 = -((faktureret/udgifterFaktisk) * 100)'100
                else
                db2 = 100
                end if
            end if
        else
        db2 = 0
        end if
        
        timerTildelt = timerialt

      
        if len(oRec("restestimat")) <> 0 then
		restestimat = oRec("restestimat")
		else
		restestimat = 0
		end if
		
		
		select case oRec("stade_tim_proc") '0 = rest i timer, 1 = i proc
		case 0
		totalforbrugt = (timerforbrugt + restestimat)
		case 1
		    
		    if timerforbrugt <> 0 AND restestimat <> 0 then
		    totalforbrugt = timerforbrugt * (100/restestimat) 
		    else
		    totalforbrugt = 0
		    end if '- (restestimat/timerTildelt) * 100)
		
		end select
		
        jobbudget = dblBudget

        
        if timerTildelt <> 0 then
        udg_internKostprTim = oRec("jo_udgifter_intern") / timerTildelt
        else
        udg_internKostprTim = oRec("jo_udgifter_intern") / 1
        end if 
        
        interKostEsti = 0
        interKostEsti = (udg_internKostprTim/1 * totalforbrugt/1)
        
       
        if jobbudget <> 0 AND oRec("restestimat") <> 0 then
            if (interKostEsti + jo_udgifter_ulev) <= jobbudget then 
            forvDb = 100 - ((interKostEsti + jo_udgifter_ulev)/jobbudget * 100)
            else
            forvDb = -((interKostEsti + jo_udgifter_ulev)/jobbudget * 100)
            end if
        else
        forvDb = 0
        end if

       
        '*** bal WIP ****'

        if cint(oRec("stade_tim_proc")) = 0 then 'rest estimat angivet i timer
		
		      if totalforbrugt <> 0 then
		
		        if restestimat = 0 then
		        stade = 100
        		else
        		stade = (timerforbrugt/totalforbrugt) * 100
		        end if
		    
		    else
		    
		    stade = 0
		    
		    end if
		    
		     afsl_proc = stade
		else
	        
	       afsl_proc = restestimat
		end if
        
        OmsWIP =  (afsl_proc/100) * jobbudget 
        balWIP = (faktureret - OmsWIP)

		
		strTxtExport = strTxtExport &strJobnavn&";"&strJobnr&";"&strStatus&";"&strKunde&"; "&strKnr&";"&dtStdato&";"_
		& dtSldato&";"&intFaspris&";"& formatnumber(timerialt, 2) &";"& formatnumber(dblBudget,2) &";"& formatnumber(salgsomk,2) &";"& formatnumber(internKost,2) &";"& formatnumber(jo_dbproc,0) &";"_
        & formatnumber(timerforbrugt,2) &";"& formatnumber(tp,2) &";"& formatnumber(OmsReal,2) &";"_
        & formatnumber(forvDb,0) &";"_
        & formatnumber(faktureret,2) &";"_
        & formatnumber(salgsOmkFaktisk,2) &";"& formatnumber(kostpris,2) &";"_
        & formatnumber(db2,0) &";"& formatnumber(gnstimepris,2) &";"& formatnumber(afsl_proc, 0) &";"& formatnumber(OmsWIP, 2) &";"& formatnumber(balWIP, 2) &";"_
		& strJobans1 &" ("& strJobans1nr &");"& strJobans1init &";"_
		& strJobans2 &" ("& strJobans2nr &");"& strJobans2init &";"_
        & strJobans3 &" ("& strJobans3nr &");"& strJobans3init &";"_
        & strJobans4 &" ("& strJobans4nr &");"& strJobans4init &";"_
        & strJobans5 &" ("& strJobans5nr &");"& strJobans5init &";"_
		& serviceaft&";"& rekvnr &";"

        '**** Fomr ****'
        strTxtFomr = ""
        strSQLf = "SELECT for_fomr, fomr.navn FROM fomr_rel "_
        &" LEFT JOIN fomr ON (fomr.id = for_fomr) WHERE for_jobid = "& jobIder(i) & " GROUP BY for_fomr"
        fo = 0
        oRec3.open strSQLf, oConn, 3
        while not oRec3.EOF 

        strTxtFomr = strTxtFomr & oRec3("navn") & ", "

        fo = fo + 1
        oRec3.movenext
        wend
        oRec3.close

        if fo <> 0 then
        len_strTxtFomr = len(strTxtFomr)
        left_strTxtFomr = left(strTxtFomr, len_strTxtFomr-2)
        strTxtFomr = left_strTxtFomr
        end if



        strTxtExport = strTxtExport & strTxtFomr & ";"


        'htmlparseCSV(strJobBesk)
        'strJobBesk = htmlparseCSVtxt 

        strJobBesk = "Format err."

        'htmlparseCSV(komm)
        'komm = htmlparseCSVtxt 

        'strJobBesk = EncodeUTF8(strJobBesk)
        'strJobBesk = replace(strJobBesk, "  ", " ")
        'komm = EncodeUTF8(komm)

		strTxtExport = strTxtExport & strJobBesk &";"& komm &";"& vbcrlf
		
		else
		
		strTxtExport = strTxtExport &strJobnavn&" ("&strJobnr&")"  
		strTxtExport = strTxtExport &" // Kontakt: "&strKunde&" ("&strKnr&")" & vbcrlf
		strTxtExport = strTxtExport &"Status: "& strStatus  & vbcrlf
		strTxtExport = strTxtExport &"Brutto Oms.: "& formatnumber(dblBudget, 2) & vbcrlf 
		strTxtExport = strTxtExport &"Udgifter: "& formatnumber(udgifter, 2) & vbcrlf 
		strTxtExport = strTxtExport &"DB %: "& formatnumber(jo_dbproc,0) & vbcrlf 
		strTxtExport = strTxtExport &"Timer forkalk.: "& formatnumber(timerialt, 2) & vbcrlf
		strTxtExport = strTxtExport &"Periode: "& dtStdato &" til " & dtSldato & vbcrlf
		strTxtExport = strTxtExport &"jobansvarlig1: "& strJobans1 & "("& strJobans1nr &") - "& strJobans1init & vbcrlf
		strTxtExport = strTxtExport &"jobansvarlig2: "& strJobans2 & "("& strJobans2nr &") - "& strJobans2init & vbcrlf
		strTxtExport = strTxtExport &"Beskrivelse:" & vbcrlf
		strTxtExport = strTxtExport &strJobBesk & vbcrlf & vbcrlf & vbcrlf 
		strTxtExport = strTxtExport &"Kommentar:" & vbcrlf & ""& komm & vbcrlf 
		strTxtExport = strTxtExport & vbcrlf & vbcrlf & vbcrlf
		strTxtExport = strTxtExport &"----------------------------------------------------------------------------------------------------" & vbcrlf & vbcrlf & vbcrlf
		
		
		end if
		
		
	end if
	oRec.close
	
	next
	
	
	objF.WriteLine(strTxtExport)

	'Response.write strTxtExport
	
	objF.close

	Response.redirect "../inc/log/data/"& file &""	
	
end if
%>
<!--#include file="../inc/regular/footer_inc.asp"-->
		
