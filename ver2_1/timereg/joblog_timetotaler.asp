<%response.buffer = true
timeA = now%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->
<!--#include file="inc/isint_func.asp"-->
<!--#include file="inc/joblog_timetotaler_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->



<%''GIT 20160811 - SK
func = request("func")
thisfile = "joblog_timetotaler"
    

    
   call bdgmtypon_fn()


	    
	'*** Altid = 1 (Brug datointerval)
	fmudato = 1
	
	'**************************************************
	'***** Faste Filter kriterier *********************
	'**************************************************


     if len(trim(request("vis_fakbare_res"))) <> 0 then '** Er der s�gt via FORM 
            
            if len(request("vis_aktnavn")) <> 0 then 
            vis_aktnavn = 1
            vis_aktnavnCHK = "CHECKED"
            else
            vis_aktnavn = 0
            vis_aktnavnCHK = ""
            end if
    
    else

      if request.cookies("cc_vis_fakbare_res")("aktn") = "1" then
                
		    vis_aktnavn = 1
            vis_aktnavnCHK = "CHECKED"

        else
		    vis_aktnavn = 0
            vis_aktnavnCHK = ""
		end if


	end if

   
    response.cookies("cc_vis_fakbare_res")("aktn") = vis_aktnavn


   

    if len(trim(request("csv_pivot"))) <> 0 then
    csv_pivot = request("csv_pivot")
    response.cookies("tsa")("csv_pivot") = csv_pivot

        else
        
           if request.cookies("tsa")("csv_pivot") <> "" then
            csv_pivot = request.cookies("tsa")("csv_pivot")
           else
            csv_pivot = 0
           end if

    end if
      
    if csv_pivot <> "0" then
    csv_pivotSEL0 = ""
    csv_pivotSEL1 = "CHECKED"
    else
    csv_pivotSEL0 = "CHECKED"
    csv_pivotSEL1 = ""
    end if

	
	'*** Job og Kundeans ***
	call kundeogjobans()
	

	
	'*** Medarbejdere / projektgrupper
	'selmedarb = session("mid")
	'call medarbogprogrp("timtot")
	medarbSQlKri = ""
	kundeAnsSQLKri = ""
	jobAns2SQLkri = ""
	jobAnsSQLkri = ""
	ugeAflsMidKri = ""
	fakmedspecSQLkri = ""
	

    if len(trim(request("nomenu"))) <> 0 then
    nomenu = request("nomenu")
    else
    nomenu = 0
    end if
	
    if len(trim(request("FM_progrp"))) <> 0 then
	progrp = request("FM_progrp")
	else
    progrp = 0
	end if
	
	'*** Rettigheder p� den der er logget ind **'
	medarbid = session("mid")
    media = request("media") 'bruges ikke da printversion kaldes


     
	if len(request("FM_medarb")) <> 0 OR func = "export" then
	
	    if left(request("FM_medarb"), 1) = "0" then 'ikke l�ngere mulig efer jq v�lg alle automatisk
	        
            if media <> "print" then
	        thisMiduse = request("FM_medarb_hidden")
    	    else
    	    thisMiduse = request("FM_medarb")
    	    end if
    	
    	intMids = split(thisMiduse, ", ")
	    else
	    thisMiduse = request("FM_medarb")
	    intMids = split(thisMiduse, ", ")
	    end if
	
	else
	    
	    thisMiduse = session("mid") 
	    intMids = split(thisMiduse, ", ")
	   
	end if
	
    'fomrLNK = ""
    strFomr_rel = "0"
	strFomr_reljobids = "0"
    strFomr_relaktids = "0"
    
    
    
    if len(trim(request("FM_fomr"))) <> 0 then
    fomr = request("FM_fomr")
    'fomrLNK = fomr


    if right(fomr, 1) = "," then 
    fomr_len = len(fomr)
    fomr_left = left(fomr, fomr_len - 1)
    fomr = fomr_left
    end if

    'fomr = ""
    'response.write "<br><br><br><br><br><br>fomr: " & fomr

    fomrArr = split(fomr, ", ")

        for f = 0 TO UBOUND(fomrArr)
         
             strFomr_rel = strFomr_rel & ", #"& fomrArr(f) & "#"  

            '*** Forrretningsomr�der Kriteerie strFomr_rel            
            strSQLfomrrel = "SELECT for_fomr, for_jobid, for_aktid FROM fomr_rel WHERE for_fomr = " & fomrArr(f) & " GROUP BY for_jobid, for_aktid"
            oRec3.open strSQLfomrrel, oConn, 3
            while not oRec3.EOF 

            strFomr_reljobids = strFomr_reljobids & ", #"& oRec3("for_jobid") & "#" 
            strFomr_relaktids = strFomr_relaktids & ", #"& oRec3("for_aktid") & "#"
        
            oRec3.movenext
            wend 
            oRec3.close

        next

    else
    fomrArr = 0
    fomr = 0
    end if

   

   

	'Response.Write "<br><br><br><br><br><br>fomr " & fomr
	'response.write "strFomr_reljobids: |" & strFomr_reljobids & "|"


	'**** Kundekri ***
	if len(request("FM_kunde")) <> 0 then
	kundeid = request("FM_kunde")
	else
	kundeid = 0
	end if

    

    if len(trim(request("FM_visMedarbNullinier"))) <> 0 AND request("FM_visMedarbNullinier") <> 0 then
        visMedarbNullinier = 1
        vis_visMedarbNullinierChk = "CHECKED"
    else
        
        if len(request("FM_kunde")) <> 0 then
            visMedarbNullinier = 0
            vis_visMedarbNullinierChk = ""
        else
            
            if request.cookies("stat")("visMedarbNullinier") = "1" then 
            visMedarbNullinier = 1
            vis_visMedarbNullinierChk = "CHECKED"
            else
            visMedarbNullinier = 0
            vis_visMedarbNullinierChk = ""
            end if

        end if
    end if

    response.cookies("stat")("visMedarbNullinier") = visMedarbNullinier
	
	'*** Kundeans ***
	strKnrSQLkri = ""
	
	
    antalm = 0
	'*** finder udfra valgte projektgrupper og medarbejdere
	'medarbSQlKri 
	'kundeAnsSQLKri
	
			    for m = 0 to UBOUND(intMids)
			    
			     if m = 0 then
			     medarbSQlKri = "(m.mid = " & intMids(m)
			     kundeAnsSQLKri = "kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = "jobans1 = "& intMids(m)  
			     jobAns2SQLkri = "jobans2 = "& intMids(m)
                 jobAns3SQLkri = "jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)
			     else
			     medarbSQlKri = medarbSQlKri & " OR m.mid = " & intMids(m)
			     kundeAnsSQLKri = kundeAnsSQLKri & " OR kundeans1 = "& intMids(m) &" OR kundeans2 = "& intMids(m) &""
			     jobAnsSQLkri = jobAnsSQLkri & " OR jobans1 = "& intMids(m)  
			     jobAns2SQLkri = jobAns2SQLkri & " OR jobans2 = "& intMids(m)
                 jobAns3SQLkri = jobAns3SQLkri & " OR jobans3 = "& intMids(m) & " OR jobans4 = "& intMids(m) & " OR jobans5 = "& intMids(m)
			     end if
			    
                antalm = antalm + 1  
			    next
			    
			    medarbSQlKri = medarbSQlKri & ")"
			    
			    
			jobAnsSQLkri =  " ("& jobAnsSQLkri &")"
			jobAns2SQLkri =  "xx (" & jobAns2SQLkri &")"
			jobAns3SQLkri =  "xx (" & jobAns3SQLkri &")"
			
	




	'** Er key acc og kundeansvarlig valgt **?
	if cint(kundeans) = 1 then
	kundeAnsSQLKri = kundeAnsSQLKri
	else
	kundeAnsSQLKri = " Kundeans1 <> -1 AND kundeans2 <> -1 "
	end if
	
	'*** forretningsomr�de ****
	if len(request("FM_fomr")) <> 0 then
	fomrid = request("FM_fomr")
	else
	fomrid = 0
	end if
	

    
	
	'***** Valgt job eller s�gt p� Job ****
	'** hvis Sog = Ja
	call jobsog()
	
	
	'*** Aftale ****
	if len(request("FM_aftaler")) <> 0 then 'AND jobid = 0 AND len(jobSogVal) = 0 then
	aftaleid = request("FM_aftaler")
	else
	aftaleid = 0
	end if
	
	
	
    
   

	response.cookies("stat").expires = date + 40
	
	
	
	'**** Alle SQL kri starter med NUL records ****
	jobidFakSQLkri = " OR jobid = -1 "
	jobnrSQLkri = " OR tjobnr = '-1' "
	jidSQLkri = " OR id = -1 "
	seridFakSQLkri = " OR aftaleid = -1 "
		
	'************ slut faste filter var **************	
	
	
	
    '***** sideskiftlinier *****
    if len(trim(request("FM_sideskiftlinier"))) <> 0 then
    sideskiftlinier = request("FM_sideskiftlinier")
    
    sideskiftlinier = replace(sideskiftlinier, ",", "")
                
                call illChar(sideskiftlinier)
                sideskiftlinier = vTxt

                call erDetInt(sideskiftlinier)

                if isInt <> 0 then
                sideskiftlinier = 0
                end if
                
    
    else
    sideskiftlinier = 15
    end if


   
	

	'*** Hvad skal vises ***
	'0 timer
	'1 fakbare timer og oms
	'2 ressource timer BRUGES IKKE 

    md_split0SEL = ""
    md_split3SEL = ""
    md_split12SEL = ""

    if len(trim(request("FM_md_split"))) <> 0 then
            
            select case request("FM_md_split")
            case 3
            md_split = 1
            md_splitVal = 3
	        md_split3SEL = "SELECTED"
            case 12
            md_split = 2
	        md_splitVal = 12
            md_split12SEL = "SELECTED"
            case else
            md_split = 0
            md_splitVal = 0
	        md_split0SEL = "SELECTED"
            end select 
           
    else
    md_split = 0
    md_splitVal = 0
	md_split0SEL = "SELECTED"
    end if


    
    if len(trim(request("FM_visnuljob"))) <> 0 AND request("FM_visnuljob") <> 0 then
    visnuljob = 1
	visnuljobCHK = "CHECKED"
    else
    visnuljob = 0
	visnuljobCHK = ""
    end if


    if len(trim(request("FM_visPrevSaldo"))) <> 0 AND request("FM_visPrevSaldo") <> 0 then
    visPrevSaldo = 1
	visPrevSaldoCHK = "CHECKED"
    else
    visPrevSaldo = 0
	visPrevSaldoCHK = ""
    end if

    if len(trim(request("FM_udspec"))) <> 0 AND request("FM_udspec") <> 0 then
    upSpec = 1
	upSpecCHK = "CHECKED"
    else
    upSpec = 0
	upSpecCHK = ""
    end if


    
    if len(trim(request("FM_directexp"))) <> 0 AND request("FM_directexp") <> 0 then
    directexp = 1
	directexpCHK = "CHECKED"
    else
    directexp = 0
	directexpCHK = ""
    end if
    
   
    

    if len(trim(request("FM_udspec_all"))) <> 0 AND request("FM_udspec_all") <> 0 then
    upSpec_all = 1
	upSpec_allCHK = "CHECKED"
    else
    upSpec_all = 0
	upSpec_allCHK = ""
    end if


   

  
	if level <= 2 OR level = 6 then
	
    '**** Vis alle timer / kun fakturerbare ****
	if len(request("vis_fakbare_res")) <> 0 then
	visfakbare_res = request("vis_fakbare_res")
	response.cookies("cc_vis_fakbare_res")("vis") = visfakbare_res
	else
		if len(request.cookies("cc_vis_fakbare_res")("vis")) <> 0 then
		visfakbare_res = request.cookies("cc_vis_fakbare_res")("vis")
		else
		visfakbare_res = 0
		end if
	end if


   
    '**** Vis kun godkenbdte ****
    if len(request("vis_fakbare_res")) <> 0 then '** Er der s�gt via FORM

        if len(request("vis_godkendte")) <> 0 then 
	    vis_godkendte = 1
	    vis_godkendteCHK = "CHECKED"
    
	    else
		
		    vis_godkendte = 0
            vis_godkendteCHK = ""
	    
        end if	
     
    else

        if len(request.cookies("cc_vis_fakbare_res")("gk")) <> 0 AND request.cookies("cc_vis_fakbare_res")("gk") <> "0" then
                
		        vis_godkendte = 1 'request.cookies("cc_vis_fakbare_res")("gk")
                vis_godkendteCHK = "CHECKED"

        else
		vis_godkendte = 0
        vis_godkendteCHK = ""
		end if


	end if

   
    response.cookies("cc_vis_fakbare_res")("gk") = vis_godkendte


     '**** Vis kun fakturerbare ****
    if len(request("vis_fakbare_res")) <> 0 then '** Er der s�gt via FORM

        if len(request("vis_fakturerbare")) <> 0 AND request("vis_fakturerbare") <> 0 then 
	    vis_fakturerbare = 1
	    vis_fakturerbareCHK = "CHECKED"
    
	    else
		
		    vis_fakturerbare = 0
            vis_fakturerbareCHK = ""
	    
        end if	
     
    else

        if len(request.cookies("cc_vis_fakbare_res")("fk")) <> 0 AND request.cookies("cc_vis_fakbare_res")("fk") <> "0" then
                
		        vis_fakturerbare = 1 'request.cookies("cc_vis_fakbare_res")("gk")
                vis_fakturerbareCHK = "CHECKED"

        else
		vis_fakturerbare = 0
        vis_fakturerbareCHK = ""
		end if


	end if

   
    response.cookies("cc_vis_fakbare_res")("fk") = vis_fakturerbare


   
     '*** Vis kontaktpersoner ***
	if len(trim(request("vis_kpers"))) <> 0 OR len(request("vis_fakbare_res")) <> 0 then '** Er der s�gt via FORM then
                    if request("vis_kpers") <> 0 then
                    vis_kpers = request("vis_kpers")
                    else
                    vis_kpers = 0
                    end if
        
	else
	   if len(request.cookies("cc_vis_fakbare_res")("kpers")) <> 0 then
	   vis_kpers = request.cookies("cc_vis_fakbare_res")("kpers")
	   else
	   vis_kpers = 0
	   end if
	end if
	
	if cint(vis_kpers) = 1 then
	kpersChk = "CHECKED"
	else
	kpersChk = ""
	end if
	
	response.cookies("cc_vis_fakbare_res")("kpers") = vis_kpers
	

     '*** Vis jobbesk ***
	if len(trim(request("vis_jobbesk"))) <> 0 OR len(request("vis_fakbare_res")) <> 0 then '** Er der s�gt via FORM then

                    if request("vis_jobbesk") <> 0 then
                    vis_jobbesk = request("vis_jobbesk")
                    else
                    vis_jobbesk = 0
                    end if
        
	else
	   if len(request.cookies("cc_vis_fakbare_res")("jobbesk")) <> 0 then
	   vis_jobbesk = request.cookies("cc_vis_fakbare_res")("jobbesk")
	   else
	   vis_jobbesk = 0
	   end if
	end if
	
	if cint(vis_jobbesk) = 1 then
	jobbeskChk = "CHECKED"
	else
	jobbeskChk = ""
	end if
	
	response.cookies("cc_vis_fakbare_res")("jobbesk") = vis_jobbesk
	
	




    

            '*** Vis Enheder ***
	        if len(trim(request("vis_enheder"))) <> 0 then
            
            vis_enheder = request("vis_enheder")
                    
	        else
	           if len(request.cookies("cc_vis_fakbare_res")("enh")) <> 0 then
	           vis_enheder = request.cookies("cc_vis_fakbare_res")("enh")
	           else
	           vis_enheder = 0
	           end if
	        end if
	
	        if cint(vis_enheder) = 1 then
	        enhChk1 = "CHECKED"
	        enhChk0 = ""
	        else
	        enhChk0 = "CHECKED"
	        enhChk1 = ""
	        end if
	
	        response.cookies("cc_vis_fakbare_res")("enh") = vis_enheder
	
	
	else
	
	visfakbare_res = 0
	vis_enheder = 0
    vis_kpers = 0
	vis_jobbesk = 0
	
	end if
	
    

   
     '*** Vis Ressourcetimer ***
	if len(trim(request("vis_restimer"))) <> 0 then
    vis_restimer = request("vis_restimer")
	else
	   if len(request.cookies("cc_vis_fakbare_res")("res")) <> 0 then
	   vis_restimer = request.cookies("cc_vis_fakbare_res")("res")
	   else
	   vis_restimer = 0
	   end if
	end if
	
	if cint(vis_restimer) = 1 then
	resChk1 = "CHECKED"
	resChk0 = ""
	else
	resChk0 = "CHECKED"
	resChk1 = ""
	end if

    response.cookies("cc_vis_fakbare_res")("res") = vis_restimer

   

     '*** Vis Normtimer ***
	if len(trim(request("vis_normtimer"))) <> 0 then
    vis_normtimer = request("vis_normtimer")
	else
	   if len(request.cookies("cc_vis_fakbare_norm")("norm")) <> 0 then
	   vis_normtimer = request.cookies("cc_vis_fakbare_norm")("norm")
	   else
	   vis_normtimer = 0
	   end if
	end if
	
	if cint(vis_normtimer) = 1 then
	normChk1 = "CHECKED"
	normChk0 = ""
	else
	normChk0 = "CHECKED"
	normChk1 = ""
	end if

    response.cookies("cc_vis_fakbare_norm")("norm") = vis_normtimer



    '*** Udspec. p� medarbejdertyper ***' 
    if len(trim(request("vis_fakbare_res"))) <> 0 then 'Er der s�gt via FORM
        if len(trim(request("FM_vis_medarbejdertyper"))) <> 0 AND request("FM_vis_medarbejdertyper") <> 0 then
        vis_medarbejdertyper = 1
        else
        vis_medarbejdertyper = 0
        end if
	else
	   if len(request.cookies("cc_vis_fakbare_res")("mtyp")) <> 0 then
	   vis_medarbejdertyper = request.cookies("cc_vis_fakbare_res")("mtyp")
	   else
	   vis_medarbejdertyper = 0
	   end if
	end if
	
    

	if cint(vis_medarbejdertyper) = 1 then
	vis_medarbejdertyperChk = "CHECKED"
	else
	vis_medarbejdertyperChk = ""
	end if

    response.cookies("cc_vis_fakbare_res")("mtyp") = vis_medarbejdertyper


     '*** Udspec. p� medarbejdertyperGrp ***' 
    if len(trim(request("vis_fakbare_res"))) <> 0 then 'Er der s�gt via FORM
        if len(trim(request("FM_vis_medarbejdertyper_grp"))) <> 0 AND request("FM_vis_medarbejdertyper_grp") <> 0 then
        vis_medarbejdertyper_grp = 1
        else
        vis_medarbejdertyper_grp = 0
        end if
	else
	   if len(request.cookies("cc_vis_fakbare_res")("mtypgrp")) <> 0 then
	   vis_medarbejdertyper_grp = request.cookies("cc_vis_fakbare_res")("mtypgrp")
	   else
	   vis_medarbejdertyper_grp = 0
	   end if
	end if
	
    

	if cint(vis_medarbejdertyper_grp) = 1 then
	vis_medarbejdertyper_grpChk = "CHECKED"
	else
	vis_medarbejdertyper_grpChk = ""
	end if

    response.cookies("cc_vis_fakbare_res")("mtypgrp") = vis_medarbejdertyper_grp
    
    response.cookies("cc_vis_fakbare_res").expires = date + 10


     '*** Redakt�r *****'
    if len(trim(request("vis_fakbare_res"))) <> 0 then '**Er der s�gt via form

   	    if len(trim(request("vis_redaktor"))) <> 0 AND request("vis_redaktor") <> 0 then
	    vis_redaktor = 1
	    vis_redaktorCHK = "CHECKED"
        response.cookies("cc_vis_redaktor")("vis") = "1"
	    else
        vis_redaktor = 0
        vis_redaktorCHK = ""
        response.cookies("cc_vis_redaktor")("vis") = "0"
        end if

     else

		    if request.cookies("cc_vis_redaktor")("vis") = "1" then
		    vis_redaktor = 1
	        vis_redaktorCHK = "CHECKED"
		    else
		    vis_redaktor = 0
            vis_redaktorCHK = ""
		    end if
	   

    end if


    


    
    '** hvis uspec p� �akt IKKE er valgt eller udspec p� medarbejdertyper ER valgt kan der ikke vises redakt�r **'
    if cint(vis_medarbejdertyper) = 1 OR visfakbare_res <> 1 then
    vis_redaktor = 0
    vis_redaktorCHK = ""
        if visfakbare_res <> 1 then
        no_redaktorTxt = ""
        else
        no_redaktorTxt = "Der kan ikke v�lges <b>Vis som redakt�r</b>, n�r <b>Medarbejdertyper</b> er sl�et til"
        end if
    else
    no_redaktorTxt = ""
    end if


    if len(trim(request("FM_segment"))) <> 0 AND request("FM_segment") <> 0 then
    segmentLnk = request("FM_segment") 
    else
    segmentLnk = 0
    end if





    strLink = "joblog_timetotaler.asp?FM_medarb="&thisMiduse&"&FM_medarb_hidden="&thisMiduse&""_
	&"&FM_job="&jobid&"&FM_start_dag="&strDag&""_
	&"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&"&FM_slut_dag="&strDag_slut&""_
	&"&FM_slut_mrd="&strMrd_slut&"&FM_slut_aar="&strAar_slut&""_
	&"&FM_jobsog="&jobSogValPrint&""_
	&"&FM_aftaler="&aftaleid&""_
	&"&FM_radio_projgrp_medarb="&progrp_medarb&""_
	&"&FM_progrp="&progrp&""_
	&"&FM_kunde="&kundeid&""_
	&"&FM_kundeans="&kundeans&""_
	&"&FM_jobans="&jobans&""_
	&"&FM_kundejobans_ell_alle="&visKundejobans&""_
	&"&viskunabnejob0="&viskunabnejob0&"&viskunabnejob1="&viskunabnejob1&""_
    &"&viskunabnejob2="&viskunabnejob2&"&FM_segment="&segmentLnk&"&nomenu="&nomenu&""_
    &"&FM_udspec_all="&upSpec_all&"&FM_udspec="&upSpec&""_
    &"&FM_visPrevSaldo="&visPrevSaldo&""_
    &"&FM_visnuljob="&visnuljob&""_
    &"&FM_md_split="&md_splitVal&""_
    &"&FM_fomr="&fomrid&""_
    &"&FM_visMedarbNullinier="&visMedarbNullinier&""_
    &"&FM_vis_medarbejdertyper="&vis_medarbejdertyper&""_
    &"&vis_godkendte="&vis_godkendte&""_
    &"&vis_fakturerbare="&vis_fakturerbare&""_
    &"&vis_redaktor="&vis_redaktor&"&FM_fomr="&fomr

    'response.write strLink
    'response.flush
    
    


    '************************************************************************'
    '************************************************************************'
    '*** Opdater timepriser *************************************************'
    '************************************************************************'
    '************************************************************************'


    if func = "dbupdatelist" then

     'Response.write strLink & "<br><br><br>"
     'Response.end
    

    '*** opdarter tp p� job ****'
    if len(request("FM_opdater_tp_job")) <> 0 then
    opdater_tp_job = 1
    else
    opdater_tp_job = 0
    end if




    '***** Opdaterer eksisterende timerpriser ******'
    jobids = split(request("FM_jobid_t"), ", ")
    jobnrs = split(request("FM_jobnr_t"), ", ")
    medids = split(request("FM_medid_t"), ", ")
    aktids = split(request("FM_aktid_t"), ", ")
    months = split(request("FM_mth_t"), ", ")
    
    tps = request("FM_tp_t")
    tps_len = len(tps)

    if tps_len > 3 then
    tps_left = left(tps, tps_len - 3)
    tpids = split(tps_left, ", #, ")

  

        '*** opdaterer timepriser ****
        for x = 0 to UBOUND(aktids)

        'thisId = 0
        tpids(x) = replace(tpids(x), ".","")
        tpids(x) = replace(tpids(x), ",",".")
                
                    
                    
                    '***** Opdaterer timepris p� job ***'
                    if opdater_tp_job = 1 then
                     
                     
                     
                     'strSQL = "SELECT id FROM timepriser WHERE jobid = "& jobids(x) &" AND aktid = "& aktids(x) &" AND medarbid = "& medids(x)
                    'Response.write strSQL & "<br><br>"
                    'oRec.open strSQL, oConn, 3
                    'if not oRec.EOF then 
            
                        'thisId = oRec("id")
            
                    'end if
                    'oRec.close


                            if trim(tpids(x)) <> "n" then 'nedarver

                             'n: s�tter akt til at nedarve, derfor ok at slette **
                             if aktids(x) = 0 then 'job g�lder alle aktiviteter UNDT KM
                             


                             '**IKKE KM akt
                             aktidKM = " AND aktid <> -1"
                             strSQLKMakt = "SELECT id FROM aktiviteter WHERE fakturerbar = 5 AND job = "& jobids(x) 
                             oRec5.open strSQLKMakt, oConn, 3
                             while not oRec5.EOF 

                             aktidKM = aktidKM & " AND aktid <> " & oRec5("id")  

                             oRec5.movenext
                             wend
                             oRec5.close
                             
                             aktSQLKri = aktidKM



                             else
                             aktSQLKri = "AND aktid = " & aktids(x)
                             end if
             
                             strSQldeltp = "DELETE FROM timepriser WHERE jobid = "& jobids(x) &" "& aktSQLKri &" AND medarbid = "& medids(x) 
                             oConn.execute(strSQldeltp)



                            'if thisId <> 0 then 'update / insert
                            'strSQLui = "UPDATE timepriser SET timeprisalt = 6, 6timepris = "& tpids(x) &" WHERE jobid = "& jobids(x) &" AND aktid = "& aktids(x) &" AND medarbid = "& medids(x) 
                            'else
                            '*** Ikke stamaktiviteter ***'
                            if jobids(x) <> 0 then
                            
                            strSQLui = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris) VALUES "_
		                    &" ("& jobids(x) &", "& aktids(x) &", "& medids(x) &", "_
		                    &" 6, "& tpids(x) &")"
                        


                            'Response.write strSQLui & "<br>"
                            'Response.flush
                            oConn.execute(strSQLui)

                            end if

                            end if' nedarver

                    end if 'opdater tp
                
                 
       

        '*** opaterer Real timer ***
        'if fra = 1 then
        'if len(trim(months(x))) <> 0 then

        fraDtSQL = year(months(x)) & "/"& month(months(x)) &"/1"
        tilDtSQL = year(months(x)) & "/"& month(months(x)) &"/31"

        strSQLuptWHdt = "tdato BETWEEN '"& fraDtSQL &"' AND '"& tilDtSQL &"'"
        'else
        'strSQLuptWHdt = "tdato >= '2000/1/1'"
        'end if

        if aktids(x) <> 0 then
        strSQLjobaktKri = " AND taktivitetid = "& aktids(x)
        else
        strSQLjobaktKri = " AND tjobnr = '"& jobnrs(x) & "'"
        end if

        strSQLupt = "UPDATE timer SET timepris = "& tpids(x) & " WHERE "& strSQLuptWHdt & " AND tmnr = "& medids(x) & strSQLjobaktKri 
        'Response.write strSQLupt & "<hr>"
        'Response.flush
        
        oConn.execute(strSQLupt)

        'end if

        next


    end if 'len timepris


    '*** timepris op� job ****'

  
    if opdater_tp_job = 1 then

    '*** Opdaterer timepriser fremadrettet ****'
    jobids = split(request("FM_jobid_j"), ", ")
    jobnrs = split(request("FM_jobnr_j"), ", ")
    medids = split(request("FM_medid_j"), ", ")
    aktids = split(request("FM_aktid_j"), ", ")
    'months = split(request("FM_mth_j"), ", ")
    tps = request("FM_tp_j")
   
        tps_len = len(tps)

        if tps_len > 3 then
        tps_left = left(tps, tps_len - 3)
        tpids = split(tps_left, ", #, ")

  

            '*** opdaterer timepriser ****
            for x = 0 to UBOUND(aktids)

            'thisId = 0
            tpids(x) = replace(tpids(x), ".","")
            tpids(x) = replace(tpids(x), ",",".")
        
            
            
           
            'n: s�tter akt til at nedarve, derfor ok at slette **
             if aktids(x) = 0 then 'job g�lder alle aktiviteter UNDT km akt
                    

                             '**IKKE KM akt
                             aktidKM = " AND aktid <> -1"
                             strSQLKMakt = "SELECT id FROM aktiviteter WHERE fakturerbar = 5 AND job = "& jobids(x) 
                             oRec5.open strSQLKMakt, oConn, 3
                             while not oRec5.EOF 

                             aktidKM = aktidKM & " AND aktid <> " & oRec5("id")  

                             oRec5.movenext
                             wend
                             oRec5.close
                             
                             aktSQLKri = aktidKM

             else
             aktSQLKri = "AND aktid = " & aktids(x)
             end if
             
             strSQldeltp = "DELETE FROM timepriser WHERE jobid = "& jobids(x) &" "& aktSQLKri &" AND medarbid = "& medids(x) 
             oConn.execute(strSQldeltp)
             
            

             if trim(tpids(x)) <> "n" AND len(trim(trim(tpids(x)))) <> 0 then 'nedarver
            
            'if thisId <> 0 then 'update / insert
            'strSQLui = "UPDATE timepriser SET timeprisalt = 6, 6timepris = "& tpids(x) &" WHERE jobid = "& jobids(x) &" AND aktid = "& aktids(x) &" AND medarbid = "& medids(x) 
            'else
            strSQLui = "INSERT INTO timepriser (jobid, aktid, medarbid, timeprisalt, 6timepris) VALUES "_
		    &" ("& jobids(x) &", "& aktids(x) &", "& medids(x) &", "_
		    &" 6, "& tpids(x) &")"
            'end if


           'Response.write strSQLui & "<br>"
           'Response.flush
           oConn.execute(strSQLui)
           end if 'nedarver

       
            next




        end if 'len tp

    end if 'opdater_tp_job = 1



    'Response.end

    %>
	<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
	<%

    %>

    <div style="position:absolute; left:100px; top:40px;"><table><tr><td>Timepriser opdateret. Vend tilbage til Grandtotal ved at klikke her <br />

        <form action="<%=strLink%>" method="post"><br /><input type="submit" value=" Videre >> " /></form>

        </td></tr></table>
    </div>
    <%
   
    Response.end
    'Response.redirect strLink '"joblog_timetotaler.asp"
    end if









	
	    '*** txt orientering ***
        'if print <> "j" then
		'tdstyleTimOms = "valign=top style='padding:2px; border-left:1px #CCCCCC solid; border-top:1px #CCCCCC solid; writing-mode:tb-rl;'"
		'tdstyleTimOms1 = "valign=top style='padding:2px; border-left:1px #CCCCCC solid; border-top:1px #CCCCCC solid; writing-mode:tb-rl;'"
        'else
        tdstyleTimOms = "valign=bottom style='padding:2px; border-left:1px #CCCCCC solid; border-top:1px #CCCCCC solid;'"
		tdstyleTimOms1 = "valign=bottom style='padding:2px; border-left:1px #CCCCCC solid; border-top:1px #CCCCCC solid;'"
        'end if

		oriChk1 = "CHECKED"
		oriChk = ""
		tdh = "70"
		
		tdstyleTimOms2 = "valign=bottom style='padding:2px; border-left:1px #CCCCCC solid; border-top:1px #CCCCCC solid;'"
		tdstyleTimOms20 = "valign=bottom style='padding:2px; border-left:1px #CCCCCC solid; border-right:1px #CCCCCC solid; border-top:1px #CCCCCC solid;'"

        tdstyleTimOms3 = "valign=bottom style='padding:2px; border-left:1px #CCCCCC solid; border-top:1px #CCCCCC solid; white-space:nowrap;'"

        tdstyleTimOms4 = "valign=bottom style='padding:2px; border-left:1px #999999 solid; border-top:1px #CCCCCC solid;'"
		
	    tdstyleTimOms10 = "valign=bottom style='padding:2px; border-left:1px #CCCCCC solid; border-right:1px #CCCCCC solid; border-top:1px #CCCCCC solid;'"

	
	if len(session("user")) = 0 then
	%>
	<!--#include file="../inc/regular/header_inc.asp"-->
	<%
	errortype = 5
	call showError(errortype)
	else
			
	
	'*** S�tter lokal dato/kr format. Skal inds�ttes efter kalender.
	Session.LCID = 1030
	%>

    
	
	

	
	
	<%

   

    if nomenu <> "1" AND media <> "print" then
    %>
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <script src="inc/joblog_timetotaler_jav.js"></script>
    <!--include file="../inc/regular/topmenu_inc.asp"-->
   

         <%

             call menu_2014()
    
           
	        

            pleft = 90
	        ptop = 102 '202

             pr_Top = 0
             hlp_top = -16
             ov_Top = 70

            'case else

            'pleft = 20
	        'ptop = 172 '202

             'pr_Top = 0
             'hlp_top = -16
             'ov_Top = 70

            'end select
    
    else
    
    if media <> "print" then
    %>
    <!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
    <%else %>
    <!--#include file="../inc/regular/header_hvd_inc.asp"-->
      


<!-- WC20160902 Udkommenter og Rotate virker
     
  <script type="text/javascript">
    function changeZoom2(oSel) {
    //alert("her")
  	newZoom= oSel.options[oSel.selectedIndex].innerText
	printarea.style.zoom=newZoom+'%';
	//oCode.innerText='zoom: '+newZoom+'';	
}

/***************************************Lei Rotate function 30-04-2013**************************/
function LeiRotate() {

   

    var div = document.getElementById("printarea");
   


    var cls = div.className;
    if (cls.indexOf("rotate2") !== -1) {    //rotate three times to the counterclockwise direction
        div.className = "rotate3";
        var widthHalf = div.clientWidth / 2;
        var heightHalf = div.clientHeight / 2;
        div.setAttribute("style", "left: 780px; top: 0px; visibility: visible; position: absolute; zoom: 100%;margin-top:" + (widthHalf - heightHalf) + "px;margin-left:-" + (widthHalf - heightHalf + 100)+"px");
    }
    else if (cls.indexOf("rotate3") !== -1) {   //rotate four times to the counterclockwise direction
        div.className = "";
        var widthHalf = div.clientWidth / 2;
        var heightHalf = div.clientHeight / 2;
        div.setAttribute("style", "left: 0px; top: -20px; visibility: visible; position: absolute; zoom: 100%;");
    }
    else if (cls.indexOf("rotate") !== -1) {    //rotate twice to the counterclockwise direction
        div.className = "rotate2";
        var widthHalf = div.clientWidth / 2;
        var heightHalf = div.clientHeight / 2;
        div.setAttribute("style", "left: 1120px; top: 620px; visibility: visible; position: absolute; zoom: 100%;");
    }
    else {      //rotate once to the counterclockwise direction
        div.className = "rotate";
        var widthHalf = div.clientWidth / 2;
        var heightHalf = div.clientHeight / 2;
        div.setAttribute("style", "left: 20px; top: 1120px; visibility: visible; position: absolute; zoom: 100%; margin-top:" + (widthHalf - heightHalf) + "px;margin-left:-" + (widthHalf - heightHalf+20)+"px");
    }    
}
</SCRIPT>

-->


<!--include file="../inc/regular/header_hvd_css3_html5_inc_Lei.asp"-->


<!--		
    <style type="text/css">
  
   
    .rotate
    {
        -webkit-transform:rotate(-90deg);
        -moz-transform:rotate(-90deg);
        -o-transform:rotate(-90deg);
        -ms-transform:rotate(-90deg);
    }
    .rotate2
    {
        -webkit-transform:rotate(-180deg);
        -moz-transform:rotate(-180deg);
        -o-transform:rotate(-180deg);
        -ms-transform:rotate(-180deg);
    }
    .rotate3
    {
        -webkit-transform:rotate(-270deg);
        -moz-transform:rotate(-270deg);
        -o-transform:rotate(-270deg);
        -ms-transform:rotate(-270deg);
    }
    </style> 

    -->

<%end if %>



    <script src="inc/joblog_timetotaler_jav.js"></script>
    <%

    pleft = 20
	ptop =  20 '112

    pr_Top = 0
    hlp_top = -16
    ov_Top = 20
    
    end if%>
	

  

     <div id="loadbar" class="load" style="position:absolute; display:; visibility:visible; top:260px; left:200px; width:300px; background-color:#ffffff; border:10px #9ACD32 solid; padding:10px; z-index:100000000;">

	<table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	<br />
	Forventet loadtid:
	<%
	
	exp_loadtid = 0
	'exp_loadtid = (((len(akttype_sel) / 3) * (len(antalvlgM) / 3)) / 50)  %> 
	ca.: <b>3-45 sek.</b><br />
    <br />&nbsp;
	</td><td align=right style="padding-right:40px;">
	<img src="../inc/jquery/images/ajax-loader.gif" />
    
  
    <br />
	<br />
    </td></tr>
    
    
        </table>

	</div>
   

	
	
	
	
	
	
	
	<%
	
	oimg = "ikon_grandtotal_48.png"
	oleft = 20
	otop = ov_Top
	owdt = 400
	oskrift = "Grandtotal, timer realiseret"
	
	'call sideoverskrift(oleft, otop, owdt, oimg, oskrift)
	
	if media = "xxxprint" then
    %>


      <div style="position:absolute; zoom: 100%; left:1200px; top:15px; z-index:1000000; visibility:visible;" bgcolor="#ffffff">
          <form>
  <p>
            <input id="Button1" type="button" value="Rotate Clockwise/CCW 90' >>" onclick="LeiRotate();" /><br />
      <span style="font-size:10px; color:#999999;">
            1-4 Portrait/Lanscape Normal - Upside/Down<br />
          
          </span>
      </form>
     </p>
    </div>
    <%end if %>



    

    <div id="sindhold" style="position:absolute; left:<%=pleft%>px; top:<%=ptop%>px; visibility:visible;">
   


    <%if media <> "print" then

    call filterheader_2013(0,0,800,oskrift)
	
	%>
	
	<table cellspacing=0 cellpadding=0 border=0 width=100% bgcolor="#FFFFFF">
    <form action="joblog_timetotaler.asp?FM_usedatointerval=1&nomenu=<%=nomenu %>" method="post">
	
	<%call medkunderjob %>
	
    </td>
    </tr>
   
        <tr><td colspan="5" style="padding-top:20px;"><span id="sp_ava" style="color:#5582d2;">[+] Avanceret</span></td></tr>
         

        <tr id="tr_ava" style="display:none; visibility:hidden;">
            <td colspan="5">
	   
		
		<table cellpadding=2 cellspacing=2 border=0 width=80%>
			
			
			<%select case cint(visfakbare_res) 
			case 1
				
				fakChk = ""
				fakChk1 = "CHECKED"
				fakChk2 = ""
				
			case 2
				
				fakChk = ""
				fakChk1 = ""
				fakChk2 = "CHECKED"
				
			case else
				
				fakChk = "CHECKED"
				fakChk1 = ""
				fakChk2 = ""
			
			end select%>

            <tr><td colspan="3" style="padding-left:40px;"><br /><b>Udspecificering p� aktiviteter</b>

               
                <br />

              <input type="checkbox" name="FM_udspec" id="FM_udspec" value="1" <%=upSpecCHK %> />Udspecificer de(t) valgte/s�gte job p� aktiviteter.
              <br />
               <input type="checkbox" name="FM_udspec_all" id="FM_udspec_all" value="1" <%=upSpec_allCHK %> disabled />Vis alle job + udspec. af de(t) s�gte job<br />
            
               <input type="checkbox" value="1" name="vis_aktnavn" id="vis_aktnavn" <%=vis_aktnavnCHK%> /> Vis kun job (og aktiviteter) hvor <input type="text" name="FM_aktnavnsog" id="FM_aktnavnsog" value="<%=aktNavnSogVal%>" style="width:200px;"> indg�r i <b>aktivitetsnavnet</b><br />
		 

                </td></tr>

            <tr>
				<td colspan="3"  style="padding-left:40px;"><br /><b>Timer:</b><br>

            </tr>
			
			<%if level <= 2 OR level = 6 then %>
			
			<tr>
				<td colspan="3" style="padding:0px 40px 20px 40px;" valign=top>
			    <input type="radio" name="vis_fakbare_res" id="vis_fakbare_res0" value="0" <%=fakChk%>> Vis <b>timer</b> (p� akt.typer der er med i dagligt timeregnskab)<br />
                <input type="radio" name="vis_fakbare_res" id="vis_fakbare_res2" value="2" <%=fakChk2%>> Vis <b>timer og kostpriser</b><br />
                <input type="radio" name="vis_fakbare_res" id="vis_fakbare_res1" value="1" <%=fakChk1%>> Vis <b>timer og oms�tning</b>
                

                
                <%
                if cint(vis_medarbejdertyper) <> 1 then%>
               
               &nbsp;<input type="checkbox" value="1" name="vis_redaktor" id="vis_redaktor" <%=vis_redaktorCHK%> /> Rediger timepriser
                
                <%
                else
                %>
               &nbsp;<input type="checkbox" value="1" name="vis_redaktor" id="Checkbox1" DISABLED /> <span style="color:#999999;">Rediger timepriser</span>
                <%
                end if
                %>
                
                <br /><br />
               
                  

                
                <input type="checkbox" value="1" name="vis_fakturerbare" id="vis_fakturerbare" <%=vis_fakturerbareCHK%> /> Vis kun <b>fakturerbare timer</b><br />
                <span style="background-color:#FFFFe1;"> <input type="checkbox" value="1" name="vis_godkendte" id="vis_godkendte" <%=vis_godkendteCHK%> /> Vis kun <b>godkendte timer</b> </span><br />
                
               </td>
			</tr>
		
          
            <tr>
             <td colspan="3" valign=top style="padding:0px 0px 0px 40px;" bgcolor="#ffffff">
		        <b>Vis enheder</b> <font class=lille>(timer * faktor) <input type="radio" name="vis_enheder" id="vis_enheder" value="1" <%=enhChk1%>>Ja &nbsp;&nbsp;<input type="radio" name="vis_enheder" id="vis_enheder" value="0" <%=enhChk0%>>Nej</td>
			</tr>
			<%else %>
        <input id="Hidden1" type="hidden" name="vis_fakbare_res" value="0" />
			
			<%end if %>

              <tr>
				<td colspan=3 valign=top style="padding:0px 0px 0px 40px;" bgcolor="#ffffff">
		        <b>Vis forecasttimer </b><font class=lille>(timebudget) <input type="radio" name="vis_restimer" id="Radio1" value="1" <%=resChk1%>>Ja &nbsp;&nbsp;<input type="radio" name="vis_restimer" id="Radio2" value="0" <%=resChk0%>>Nej &nbsp;<span style="color:red; font-size:9px;">Ekstra loadtid!</span></td>
			</tr>

             <tr>
				<td colspan=3 valign=top style="padding:0px 0px 0px 40px;" bgcolor="#ffffff">
		        <b>Vis normtid </b><input type="radio" name="vis_normtimer" id="Radio1" value="1" <%=normChk1%>>Ja &nbsp;&nbsp;<input type="radio" name="vis_normtimer" id="Radio2" value="0" <%=normChk0%>>Nej</td>
			</tr>

			
			</table>
		
		
		</td>
	</tr>

	
	<tr>
	<td valign=top colspan="2" style="padding-top:20px;">
	<!-- Brug altid datointerval, FM_usedatokri = 1 -->
			<input type="hidden" name="FM_usedatokri" id="FM_usedatokri" value="1">
			<b>Periode:</b><br>
			<!--#include file="inc/weekselector_s.asp"--> <!-- b -->
            
       
	</td>
   
    </tr>

           <tr><td colspan="5" style="padding-top:20px;"><span id="sp_pre" style="color:#5582d2;">[+] Pr�sentation</span></td></tr>
         

        <tr id="tr_pre" style="display:none; visibility:hidden;">

	<td colspan="2" style="padding:10px 40px 40px 40px;">
  
    
     <b>Pr�sentation, indstillinger og kolonner:</b><br />
        
        
                 <input type="checkbox" name="FM_directexp" id="directexp" value="1" <%=directexpCHK %>/>Vis direkte i excel (.csv, hurtig visning af store datam�ngder)<br />

                <br />
                <b>.csv layout:</b><br />
                <input type="radio" name="csv_pivot" id="csv_pivot0" value="0" <%=csv_pivotSEL0 %>> Layout optimeret (som p� sk�rm)<br />
                <input type="radio" name="csv_pivot" id="csv_pivot1" value="1" <%=csv_pivotSEL1 %>> Pivot optimeret (kun real. timer p� medarb.)

        
        <br /><br />  <br />
     
     Vis opdelt p� m�neder, total, 3 ell. 12 md.<br />Vis periode slutdato - 
      <select name="FM_md_split" style="width:200px;" onchange="submit();" />
      <option value="0" <%=md_split0SEL %>>Ingen (vis total)</option>
      <option value="3" <%=md_split3SEL %>>3 m�neder</option>
      <option value="12" <%=md_split12SEL %>>12 m�neder (ekstra loadtid!)</option>
      </select>
      <br /><br />
      <b>Sideskift</b> (overskrift) efter hver <input id="FM_sideskiftlinier" type="text" name="FM_sideskiftlinier" value="<%=sideskiftlinier %>" style="width:40px;"/>  linie. A4 ca. 15 A3 ca.30.
      <br /><br />
      <!--<input id="Checkbox2" type="checkbox" name="FM_visnuljob" value="1" <%=visnuljobCHK %> />Vis job uden real. timer i valgte periode.<br />-->
      <input id="Checkbox3" type="checkbox" name="FM_visPrevSaldo" value="1" <%=visPrevSaldoCHK %> />Vis <b>saldo</b> (for tidligere periode) og <b>total</b> timeforbrug (alle medarbejdere uanset valgte)<br />
      <input type="checkbox" value="1" name="vis_kpers" id="vis_kpers" <%=kpersCHK%> /> Vis kontaktpersoner<br />
    <input type="checkbox" value="1" name="vis_jobbesk" id="vis_jobbesk" <%=jobbeskCHK%> /> Vis jobbeskrivelse<br />
      </td>
       </tr>
       <tr>
           <td align="right" colspan="2">


	<input type="submit" value=" K�r >> "></td>
	</tr>
	</form>
	</table>



	<!-- filterDiv -->
	</td></tr>
	</table>
    </div>

    <%else 
        
    dontshowDD = 1%>
    <!--#include file="inc/weekselector_s.asp"-->
    <%end if 'mdia = print %>
	

	   <!--pagehelp-->

    <%if print <> "j" AND media <> "print" then

                itop = hlp_top 'ptop - 147 '56
                ileft = 635
                iwdt = 120
                ihgt = 0
                ibtop = 40 
                ibleft = 150
                ibwdt = 600
                ibhgt = 410
                iId = "pagehelp"
                ibId = "pagehelp_bread"
                call sideinfoId(itop,ileft,iwdt,ihgt,iId,phDsp,phVzb,ibtop,ibleft,ibwdt,ibhgt,ibId)
                %>
                       
			                <b>Visning af de-aktiverede medarbejdere: </b><br> de-aktiverede medarbejdere medtages ved s�gning p� "alle".<br>
			                de-aktiverede medarbejdere kan ikke v�lges fra dropdown menu.<br><br>
			                <b>Jobtyper</b><br /> Fastpris eller Lbn. timer. (v�gtet medarbejder timepriser)<br>
			                Ved fastpris job, hvor aktiviteterne danner grundlag for timepris, er oms�tning og timepriser (p� jobvisning) 
			                angivet med et ~, da bel�bet er beregnet udfra en tiln�rmet timepris. (gns. p� akt.). Ved udspecificering p� aktiviteter er det den re-elle timepris der vises.<br />
			                <br />
			                Ved lbn. timer er det altid den realiserede medarbejder timepris der vises.<br />
			                <br />
			                <b>Timer</b><br />
			                Kun realiserede timer p� typer der t�ller med i det daglige timeregnskab vises. 
			                <br />Registreringer p� ferie, frokost, sygdom og afspadsering er ikke med i Grand-total statistikken.<br />
                            <br />
                            <b>Key-account</b><br />
                            Ved brug af "key account" bliver der vist timer for alle medarbejder der er med (tilknyttet via projektgrupper) p� de job hvor den valgte Key-account er job ansvarlig / kunde ansvarlig. 
                            <br /><br />
                            <b>Udspecificering p� aktiviteter</b><br />
                            Hvis der v�lges et enkelt job bliver timeforbruget udspecificeret p� aktiviteter.
                            <br /><br />
                            <b>Bel�b</b><br />
                            Alle timepriser og bel�b er vist i <%=basisValISO %>.&nbsp;
                		
                		
		                </td>
	                </tr></table></div>
	            
	
	
	<%end if
	

        'directexp = 1

	
	
	'***** Periode *****
	datoStart = strDag&"/"&strMrd&"/"&strAar
	datoSlut = strDag_slut&"/"&strMrd_slut&"/"&strAar_slut

    select case md_split
    case 1 '3 md

    'datoStart = dateadd("m", -2, datoSlut)
    datoMiddle = dateadd("m", -1, datoSlut)
    sqlDatoStart = year(datoStart)&"/"&month(datoStart)&"/"&day(datoStart)
    sqlDatoSlut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	md_split_cspan = 3 'datediff("m", datoStart, datoSlut, 2,2)

    mdStart = month(datoStart)
    mdslut = month(datoSlut)

    mdThis1 = left(monthname(month(datoStart)), 3) &"<br>"& year(datoStart)
    mdThis2 = left(monthname(month(datoMiddle)), 3) &"<br>"& year(datoMiddle)
    mdThis3 = left(monthname(month(datoSlut)), 3) &"<br>"& year(datoSlut)

    case 2 '12 md

    'datoStart = dateadd("m", -2, datoSlut)
   
    sqlDatoStart = year(datoStart)&"/"&month(datoStart)&"/"&day(datoStart)
    sqlDatoSlut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	md_split_cspan = 12 'datediff("m", datoStart, datoSlut, 2,2)

    mdStart = month(datoStart)
    mdslut = month(datoSlut)

  
          
    case else '*** Ingen opdeling
	md_split_cspan = 1
    sqlDatoStart = strAar&"/"&strMrd&"/"&strDag
    sqlDatoSlut = strAar_slut&"/"&strMrd_slut&"/"&strDag_slut
	end select



    'response.write "datoStart: " &  datoStart
    'response.Flush

    perinterval = datediff("d", datoStart, datoSlut) 


    'Response.write "jobid " & jobid
    'Response.write 
	if media = "print" then
    call segment_kunder
    end if
	

    '*** valgte job ***
	call valgtejob()
		
		
    '**** Valgte aftaler *****
	call valgteaftaler()


    '*** Begr�nsninger **'
    antalJobs = split(jobnrSQLkri, "OR")
    antJob = 0
    for i = 0 TO UBOUND (antalJobs)
    antJob = antJob + 1 
    next

    'Response.write antJob
    'Response.flush

    if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
    '** AB server
    antJobkri = 150
    maksAntalM = 1300
    antJobkri12m = 20
    
    perHigh = 730
    perAarHigh = 2

    perMid = 184
    perAarMid = 6 'md

    perLow = 62
    perAarLow = 2 'md
            
    else

    '** Rackhosting server
    antJobkri = 600
    maksAntalM = 2000
    antJobkri12m = 100

    perHigh = 730
    perAarHigh = 2 '�r

    perMid = 368
    perAarMid = 24 'md

    perLow = 184
    perAarLow = 6 'md
    end if

    'Response.write datoStart & "#" & datoSlut

  

    

	if (datediff("d", datoStart, datoSlut) > perHigh AND cint(antJob) > antJobkri AND (antalm > 1 AND antalm <= 49)) _
    OR (datediff("d", datoStart, datoSlut) > perMid AND cint(antJob) > antJobkri AND (antalm > 49 AND antalm <= 149)) _
    OR (datediff("d", datoStart, datoSlut) > perLow AND cint(antJob) > antJobkri AND (antalm > 149 AND antalm <= 1499)) _
    OR (cint(antJob) > antJobkri12m AND antalm > 50 AND md_split = 2) _
    OR (antalm > maksAntalM) _
    OR (datediff("d", datoStart, datoSlut) > 1855) then
	%>
	
	
	
	<%
	
	itop = 200
	ileft = 140
	iwdt = 500
	iId = "jlogtot_err"
	idsp = ""
	ivzb = "visible"
	call sidemsgId(itop,ileft,iwdt,iId,idsp,ivzb) 
	
	%>	
			<br><b>Antallet af job, medarbejdere, kontakter og periode overstiger det tilladte antal.</b>
			<br>
			
			<b>Hvis der er valgt:</b><br /><br />
			
			<b>A)</b> Der er valgt mere end <b><%=antJobkri %> job</b> og: <br />
			<br> - Mellem <b>2 og 50</b> medarbejdere og en periode p� mere end <b><%=perAarHigh %> �r.</b>
			<br> - Mere end <b>50</b> medarbejdere og en periode p� mere end <b><%=perAarMid %> md.</b>
            <br> - Mere end <b>150</b> medarbejdere og en periode p� mere end <b><%=perAarLow %> md.</b>
            <br> - Mere end <b><%=maksAntalM %></b> medarbejdere. 
            
            
            <br /><br />

            <b>B)</b> En periode p� mere end <b>5 �r.</b><br /><br />

            <b>C)</b> 12 m�neders opdeling, mere en <b>20 job</b> og flere end <b>50 medarbejdere.</b><br /><br />

            <span style="color:red;">Du har er valgt: <b><%=antalm %></b> medarbejdere og <b><%=antJob %></b> job.</span><br /><br />

            <b>Tip:</b><br />
            S�g p� et specifikt job for at f� vist en l�ngere periode / flere medarbejdere.
			
			</td></tr>
			</table>
			</div>
			<br /><br /><br /><br /><br /><br /><br />&nbsp;
	<%
	
	x = 0

      response.End
    end if 'maks kriterier alm.
	
 
   


	
	
		'*****************************************************************************************************'
		'**** Job der skal indg� i oms�tning, budget og forbrugstal *******************************'
		'*****************************************************************************************************'
		
		
		
				
				
				
				'*** For at spare (trimme) p� SQL hvis der v�lges alle job alle kunder og vis kun for jobanssvarlige ikke er sl�et til ****
				'*** Og der ikke er s�gt p� jobnavn ***
				if cint(kundeid) = 0 AND cint(jobid) = 0 AND cint(jobans) = 0 AND cint(jobans2) = 0 AND cint(jobans3) = 0 AND cint(kundeans) = 0 AND len(trim(jobSogVal)) = 0 AND cint(aftaleid) = 0 AND cint(segment) = 0 then 
						
					'jobidFakSQLkri = " OR jobid <> 0 "
					jobnrSQLkri = " OR tjobnr <> '0' "
					jidSQLkri =  " OR id <> 0 "
					seridFakSQLkri = " OR aftaleid <> 0 "
						
				end if
	
	
		'**************** Trimmer SQL states ************************
		'len_jobidFakSQLkri = len(jobidFakSQLkri)
		'right_jobidFakSQLkri = right(jobidFakSQLkri, len_jobidFakSQLkri - 3)
		'jobidFakSQLkri =  right_jobidFakSQLkri
		
		len_jobnrSQLkri = len(jobnrSQLkri)
		right_jobnrSQLkri = right(jobnrSQLkri, len_jobnrSQLkri - 3)
		jobnrSQLkri =  right_jobnrSQLkri
		
		len_jidSQLkri = len(jidSQLkri)
		right_jidSQLkri = right(jidSQLkri, len_jidSQLkri - 3)
		jidSQLkri =  right_jidSQLkri
		
		
		len_seridFakSQLkri = len(seridFakSQLkri)
		right_seridFakSQLkri = right(seridFakSQLkri, len_seridFakSQLkri - 3)
		seridFakSQLkri =  right_seridFakSQLkri
		
	
		'*****************************************************************************************************
	%>
	
	<%'response.flush %>
	<!-- oCode -->
	<div id="printarea" style="position:relative; top:-18px;"><!-- zoom: 100%;-->
	<%	

            if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
			m = 55 '55 '55 '40
			x = 26500 '26500 '40500 '25500 '6500 '2500 (20000 = ca. 4-500 medarb.)
            else
            m = 80 '55 '55 '40
			x = 90500 '26500 '40500 '25500 '6500 '2500 (20000 = ca. 4-500 medarb.)
            end if
			lastjid = 0
			strMids = 0
			
			
			dim jobmedtimer ', jobtimerIaltX
			redim jobmedtimer(x,m) ', jobtimerIaltX(x)
			
			v = 1150
			j = 0
			
           
			dim medarb, medabTottimer, restimerTot, omsTot, fakbareTimerTot, normTimerTot, medabTotenh, medabRestimer, medabNormtimer
            dim subMedabRestimer, subMedabTottimer, subMedabTotenh, omsSubTot
            dim medabTottimerprMd, medabTotRestimerprMd, medabTotEnhprMd, medabTotOmsprMd, medabTotTpprMd

			Redim medabTottimer(v), restimerTot(v), omsTot(v), fakbareTimerTot(v), normTimerTot(v), medabTotenh(v), medabRestimer(v), medabNormtimer(v), medarb(v), medarbnavnognr(v), medarbnavnognr_short(v)
            redim subMedabRestimer(v), subMedabTottimer(v), subMedabTotenh(v), omsSubTot(v)

            redim medabTottimerprMd(v, 12), medabTotRestimerprMd(v, 12), medabTotEnhprMd(v, 12), medabTotOmsprMd(v, 12), medabTotTpprMd(v, 12)
							
			'dim  jobTottimer, jobRestimerTot, jobOmsTot, jobfakbareTimerTot
			'Redim  jobTottimer(0), jobRestimerTot(0), jobOmsTot(0), jobfakbareTimerTot(0)
			
			
			'*** Udspecificer p� aktiviteter? **'
			call akttyper2009(2)
			
             select case lto
            case "tec", "esn"
            aktypRealtimerSQLkri = replace(aty_sql_active, "tfaktim", "a.fakturerbar")
            case else 
            aktypRealtimerSQLkri = replace(aty_sql_realhours, "tfaktim", "a.fakturerbar")
            end select

		
			
			
			
			'**** Kun timer eller beregn fakbare timer og oms�tning ***'
			if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then 
			    if cint(visfakbare_res) = 1 then
                vgtTimePris = " COALESCE(sum(t.timer*(t.timepris*(t.kurs/100))),0) AS vaegtettimepris, "
                else 'kostpris
                vgtTimePris = " COALESCE(sum(t.timer*(t.kostpris*(t.kurs/100))),0) AS vaegtettimepris, "
                end if
			else
		    vgtTimePris = ""
			end if
			
			
            '** Vis kun fakurerbare ***
            if cint(vis_fakturerbare) = 1 then
            whrClaus = "("& replace(aty_sql_fakbar, "fakturerbar", "t.tfaktim") &")" 
			
            else
            
            select case lto
            case "tec", "esn"
            whrClaus = "("& aty_sql_active &")"
            case else 
            whrClaus = "("& aty_sql_realhours &")"
            end select
        
            end if

            'if session("mid") = 1 then
            'response.write "<br><br><br>vis_fakturerbare: "& vis_fakturerbare &" lto:" & lto &" whrClaus: " & whrClaus  & "<br><br>"
            'end if

            nDatoStart = day(sqlDatoStart) &"/"& month(sqlDatoStart) &"/"& year(sqlDatoStart)
			
			
			
			'***********************************************************************'
			'**** Job og medarbjedere med timer p� *********************************'
			'***********************************************************************'
			
			'*** Finder de medarbejdere af de valgte der er timer p� i den valgte periode ***'
            if cint(vis_medarbejdertyper_grp) = 1 then 'Fordeling p� medarbejertypegrp valgt
                medarbSQlKri = ""
                vlgtmtypgrp = thisMiduse '0 'mtypGrupper	    
                call mtyperIGrp_fn(vlgtmtypgrp,1)
                medarbSQlKri = medarbSQlKriMtypGrp
            else
			    if cint(jobans) = 1 OR cint(kundeans) = 1 then
			    medarbSQlKri = " m.mid <> 0 "
			    else
			    medarbSQlKri = medarbSQlKri 'sqlMedKri
			    end if
			end if
			
            'Response.write medarbSQlKri &"<hr>"
            'Response.Write "<br>whrClaus " & whrClaus
            'Response.Write "<br>jidSQLkri " & jidSQLkri
            'Response.Write "<br>Jobid "& jobid
            'Response.Write "<br><br><br>"

            
            'jobIsWrt = " AND ("

            ' if ja = 0 then 
			x = 0
			k = 0
			v = 0
            antaljob = 0
            TimeprisFaktiskTot = 0
			timeprisThis = 0
			strJobnbs = "#0#,"
			strMidsK = "#0#,"
            'end if

            'Response.write "jobSogVal " & jobSogVal
            'Response.Write "<br>upSpec" & upSpec 
            'Response.Write "<br>upSpec_all" & upSpec_all
            'Response.write "<br>jobid" & jobid

            'vis_medarbejdertyper = vis_medarbejdertyper

             '** Vis kun hvor aktnavn indg�r **'
            if cint(vis_aktnavn) = 1 then
            strSQLaktNavnKri = " AND taktivitetnavn LIKE '%"& trim(aktNavnSogVal) &"%'"
            else
            strSQLaktNavnKri = ""
            end if

            

            ja = 0
            for ja = 0 to upSpec '0 = nej, 1 ja
            
            
            if ja = 1 then
            strJobnbs = "#0#,"
            end if

            '*** Kriterier Job eller udspecificer p� akt.

            'Response.Write "ja: "& ja & " upSpec: "& upSpec &"jobid: "& jobid &"<br>"
            if ja = 1 OR (cint(upSpec) = 1 AND jobid <> 0)  then
				
				
				grpBySQL = "a.fase, a.id"
                
                if cint(vis_medarbejdertyper_grp) = 1 then        
                
                    grpBySQL = grpBySQL & ", m.medarbejdertype_grp"
                else
                        if cint(vis_medarbejdertyper) = 1 then
                        grpBySQL = grpBySQL & ", m.medarbejdertype"
                        else
                        grpBySQL = grpBySQL & ", m.mid"
                        end if
                
                end if	
                
                if cint(md_split) = 1 OR cint(md_split) = 2 then
                grpBySQL = grpBySQL & ",  MONTH(tdato)"
                end if

				orderBySQL = "k.kkundenavn, a.job, a.fase, a.sortorder, a.id"
			    
                if cint(vis_medarbejdertyper_grp) = 1 then        
                orderBySQL = orderBySQL & ", m.medarbejdertype_grp"
                
                else

                    if cint(vis_medarbejdertyper) = 1 then
                    orderBySQL = orderBySQL & ", m.medarbejdertype"
                    else
                    orderBySQL = orderBySQL & ", m.mid"
                    end if	
                end if

                if cint(md_split) = 1 OR cint(md_split) = 2 then
                orderBySQL = orderBySQL & ", tdato"
                end if	
			    

         

               
			
			'sqlJobKri = " j.id = "& jobid 
			aktSQLFields = ", taktivitetid, taktivitetnavn, a.id AS aid, a.job, a.navn AS anavn, a.fakturerbar AS afakbar, a.budgettimer AS abudgettimer, a.aktbudget AS abudget, a.fase, a.aktbudgetsum, a.bgr, a.antalstk, a.brug_fasttp "
			aktSQLTables = ", aktiviteter a "
			aktSQLWhere = " AND (a.job = j.id AND ("&aktypRealtimerSQLkri&")) " 'a.fakturerbar <> 2

            if cint(vis_medarbejdertyper) = 1 AND cint(vis_medarbejdertyper_grp) = 0 then
            aktSQLWhere = aktSQLWhere & " AND mty.sostergp <> -1" 
            end if

			aktjobTimerWhere = " t.taktivitetid = a.id "
			
            
			else
				
				grpBySQL = "jobnr"

                if cint(vis_medarbejdertyper_grp) = 1 then        
                
                    grpBySQL = grpBySQL & ", m.medarbejdertype_grp"
                else
                    if cint(vis_medarbejdertyper) = 1 then
                    grpBySQL = grpBySQL & ", m.medarbejdertype"
                    else
                    grpBySQL = grpBySQL & ", m.mid"
                    end if	
                end if
                
                if cint(md_split) = 1 OR cint(md_split) = 2 then
                grpBySQL = grpBySQL & ", MONTH(tdato)"
                end if

				orderBySQL = "k.kkundenavn, jobnr"
                
                if cint(vis_medarbejdertyper_grp) = 1 then        
                orderBySQL = orderBySQL & ", m.medarbejdertype_grp"
                
                else
                    if cint(vis_medarbejdertyper) = 1 then
                    orderBySQL = orderBySQL & ", m.medarbejdertype"
                    else
                    orderBySQL = orderBySQL & ", m.mid"
                    end if	
                end if    


                if cint(md_split) = 1 OR cint(md_split) = 2 then
                orderBySQL = orderBySQL & ", tdato"
                end if	
				
			aktSQLFields = ""
			aktSQLTables = ""
			aktSQLWhere = ""

            if cint(vis_medarbejdertyper) = 1 AND cint(vis_medarbejdertyper_grp) = 0 then
            aktSQLWhere = aktSQLWhere & " AND mty.sostergp <> -1"
            end if

			aktjobTimerWhere = " t.tjobnr = j.jobnr "
			end if


            if cint(vis_godkendte) = 1 then
            strSQLgk = " AND godkendtstatus = 1"
            else
            strSQLgk = ""
            end if

            'if cint(vis_fakturerbare) = 1 then
            'strSQLfk = " AND godkendtstatus = 1"
            'else
            'strSQLfk = ""
            'end if
            


            '****************************************************************
			if cint(ja) = 0 then '*** Job / aktiviteter *** 0 = Job, 1 = Akt

			
			
			        '*** Finder de job af de valgte der er timer p� i den valgte periode ***'
                    if cint(upSpec) = 1 AND jobid = 0 then
            
                        if upSpec_all = 1 then
                        sqlJobKritemp = replace(jidSQLkri, "id", "j.id") 
                        sqlJobKritemp0 = "j.id <> 0 AND ("
                        sqlJobKritemp1 = replace(sqlJobKritemp, "OR", "AND")
                        sqlJobKritemp1 = replace(sqlJobKritemp1, "=", "<>") 
                        sqlJobKritemp0 = sqlJobKritemp0 & sqlJobKritemp1 & ")"
                        else

                        '** Ved valgt job i dd og vis udspec sl�et til, skal job ikke vises p� job niveau
                        sqlJobKritemp0 = "j.id = 0 "
                        end if

                    else
			
                    'if (cint(upSpec) = 1 AND len(trim(jobSogVal)) = 0 AND jobid = 0 then 
                    ''** Ved alle job valgt i dd og vis udspec sl�et til, skal job ikke vises p� job niveau
                    'sqlJobKritemp0 = "j.id = 0 "
                    'else
                    sqlJobKritemp0 = replace(jidSQLkri, "id", "j.id") 'sqlJobKri
                    'end if
			
                    end if

            
                    aktidsSQLkriUse = ""
			        
                               
                                        
           

			                            strSQLj = "SELECT tmnr, j.id AS jid, jobnr, t.tmnr FROM timer t, job j "_
					                    &" WHERE ("& whrClaus &" AND tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoSlut &"' AND "& replace(medarbSQlKri, "m.mid", "t.tmnr") &" "& strSQLaktNavnKri &") "_
					                    &" AND ((j.jobnr = t.tjobnr) AND ("& sqlJobKritemp0 &")) "& jobstKri &" "& strSQLgk &" GROUP BY j.id, t.tmnr"
					
                                    
                                        'if session("mid") = 1 then
                                        'Response.write "<br><br>Finder Job: "& strSQLj
					                    'Response.flush
		                                'end if
        			
					                    sqlMedKri = " m.mid = 0 "
					                    sqlJobKri = " j.id = 0 "
					
					                    j = 0
					                    oRec.open strSQLj, oConn, 3
					                    while not oRec.EOF
					                    

                                        '** Forretningsomr�deKRI
                                        foromrKriOK = 1
                                        if strFomr_reljobids <> "0" then
                                       
                                            if instr(strFomr_reljobids, "#"& oRec("jid") &"#") = 0 then 'IKKE EN DEL AF FORRETNINGSOMR�DE
                                            foromrKriOK = 0
                                            end if                        
            
                                        end if


                                        if cint(foromrKriOK) = 1 then

                                        if instr(sqlJobKri, " OR j.id = " & oRec("jid") & "#") = 0 then  
					                    sqlJobKri = sqlJobKri & " OR j.id = " & oRec("jid") & "#"
                                        end if
                                        
                                        
                                        if instr(sqlMedKriStrTjk, ",#"& oRec("tmnr") &"#") = 0 then
                                        sqlMedKri = sqlMedKri & " OR m.mid = " & oRec("tmnr")
                                        end if
                                        
                                        sqlMedKriStrTjk = sqlMedKriStrTjk & ",#"& oRec("tmnr") &"#" 

                                        end if' foromrKriOK = 1
					
					                    j = j + 1
					                    oRec.movenext
					                    wend
					
					                    oRec.close


                                    
                                        '** Hvis Ressource timer er sl�et til skal de ogs� vises. ********
                                        if cint(vis_restimer) = 1 then
                
                                        sqlJobKritemp0res = replace(jidSQLkri, "id", "jobid")
                                        ressourceFCper = " ((aar >= "& year(sqlDatoStart)&" AND md >= "& month(sqlDatoStart) &") AND (aar <= "& year(sqlDatoSlut) &" AND md <= "& month(sqlDatoSlut) &"))"
                                        

                                        strSQLjobMforecast = "SELECT timer, jobid, medid FROM ressourcer_md AS rs WHERE " & ressourceFCper & " AND "& replace(medarbSQlKri, "m.mid", "rs.medid") &" AND ("& sqlJobKritemp0res &") GROUP BY jobid"

                                        'response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"& strSQLjobMforecast
                                        'response.Flush
                                        

                                        oRec.open strSQLjobMforecast, oConn, 3
					                    while not oRec.EOF                
                                        
                                          '** Forretningsomr�deKRI
                                        foromrKriOK = 1
                                        if strFomr_relaktids <> "0" then
                                       
                                            if instr(strFomr_relaktids, "#"& oRec("jobid") &"#") = 0 then 'IKKE EN DEL AF FORRETNINGSOMR�DE
                                            foromrKriOK = 0
                                            end if                        
            
                                        end if


                                        if cint(foromrKriOK) = 1 then	

                                        'sqlMedKri = sqlMedKri & " OR m.mid = " & oRec("medid")
                                        
                                        if instr(sqlJobKri, " OR j.id = " & oRec("jobid") & "#") = 0 then  
					                    sqlJobKri = sqlJobKri & " OR j.id = " & oRec("jobid") & "#"
                                        end if

                                        
                                        end if

                                        oRec.movenext
					                    wend
					
					                    oRec.close
                                        end if 'vis resttimer


                                       '*****************************************************************

                            
			
			
			  else 'aktiviteter
					
					
                    
                            sqlJobKri = " j.id = 0 "
				            sqlAktKritemp = replace(jidSQLkri, "id", "a.job")
					        strSQLa = "SELECT tmnr, a.id, t.taktivitetid, a.job FROM timer t, aktiviteter a "_
					        &" WHERE ("& whrClaus &" AND tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoSlut &"' "_
					        &" AND "& replace(medarbSQlKri, "m.mid", "t.tmnr") &" "& strSQLgk &" "& replace(strSQLaktNavnKri, "j.id", "a.job") &") AND (a.id = t.taktivitetid AND ("&sqlAktKritemp&")) GROUP BY a.id, t.tmnr"
					
                            '"& jobIsWrt &"
					
                            'if session("mid") = 1 then
					        'Response.write "<br><br><br><br><br><hr>Jobsogval: "& Jobsogval &"<br><br>AKtSQL: "& strSQLa & "<br><br>medarbSQlKri: " & medarbSQlKri
					        'Response.flush
		                    'end if			

					        aktidsSQLkriUse = " a.id = 0 "
					        sqlMedKri = " m.mid = 0 "
					
					
					        j = 0
					        oRec.open strSQLa, oConn, 3
					        while not oRec.EOF
		
                                        
                                         '** Forretningsomr�deKRI
                                        foromrKriOK = 1
                                        if strFomr_relaktids <> "0" then
                                       
                                            if instr(strFomr_relaktids, "#"& oRec("id") &"#") = 0 then 'IKKE EN DEL AF FORRETNINGSOMR�DE
                                            foromrKriOK = 0
                                            end if                        
            
                                        end if


                                        if cint(foromrKriOK) = 1 then			


					
					                    sqlMedKri = sqlMedKri & " OR m.mid = " & oRec("tmnr")
                                        'sqlJobKri = sqlJobKri & " OR j.id = " & oRec("job")
		                    
                                        if instr(sqlJobKri, " OR j.id = " & oRec("job") & "#") = 0 then  
					                    sqlJobKri = sqlJobKri & " OR j.id = " & oRec("job") & "#"
                                        end if			        

                                        
                                        if instr(aktidsSQLkriUse, " OR a.id = " & oRec("id") & "#") = 0 then  
					                    aktidsSQLkriUse = aktidsSQLkriUse & " OR a.id = "& oRec("id") & "#"
                                        end if	                                


                                        end if

                            
					
					        j = j + 1
					        oRec.movenext
					        wend
					
					        oRec.close


                                        '** Hvis Ressource timer er sl�et til skal de ogs� vises. ********
                                        if cint(vis_restimer) = 1 then
                                        ressourceFCper = " ((aar >= "& year(sqlDatoStart)&" AND md >= "& month(sqlDatoStart) &") AND (aar <= "& year(sqlDatoSlut) &" AND md <= "& month(sqlDatoSlut) &"))"
                                        sqlJobKritemp0res = replace(jidSQLkri, "id", "aktid")


                                        strSQLAktMforecast = "SELECT timer, jobid, aktid, medid FROM ressourcer_md AS rs WHERE " & ressourceFCper & " AND "& replace(medarbSQlKri, "m.mid", "rs.medid") &" AND ("& sqlJobKritemp0res &") GROUP BY aktid"

                                        'response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"& strSQLAktMforecast
                                        'response.Flush
                                        

                                        oRec.open strSQLAktMforecast, oConn, 3
					                    while not oRec.EOF                

                                         '** Forretningsomr�deKRI
                                        foromrKriOK = 1
                                        if strFomr_relaktids <> "0" then
                                       
                                            if instr(strFomr_relaktids, "#"& oRec("aktid") &"#") = 0 then 'IKKE EN DEL AF FORRETNINGSOMR�DE
                                            foromrKriOK = 0
                                            end if                        
            
                                        end if


                                        if cint(foromrKriOK) = 1 then
        
                                        sqlMedKri = sqlMedKri & " OR m.mid = " & oRec("medid")
                                        
                                        if instr(sqlJobKri, " OR j.id = " & oRec("jobid") & "#") = 0 then  
					                    sqlJobKri = sqlJobKri & " OR j.id = " & oRec("jobid") & "#"
                                        end if		
        			
                                        
                                        if instr(aktidsSQLkriUse, " OR a.id = " & oRec("aktid") & "#") = 0 then  
					                    aktidsSQLkriUse = aktidsSQLkriUse & " OR a.id = "& oRec("aktid") & "#"
                                        end if
                    
                                        end if

                                        
                                        oRec.movenext
					                    wend
					                    oRec.close
                                        
                                        'response.write "<br><br>aktidsSQLkriUse: " &aktidsSQLkriUse

            
                                        end if 'vis resttimer


                                       '*****************************************************************

					
					
					        aktidsSQLkriUse = " AND ("& aktidsSQLkriUse &")"	
			
			
			end if
            

            sqlJobKri = replace(sqlJobKri, "#", "")
            aktidsSQLkriUse = replace(aktidsSQLkriUse, "#", "")

            'Response.write "<br><br>sqlJobKri: " & sqlJobKri
            'Response.write "<br><br>medarbKri: " & replace(sqlMedKri, "OR m.mid =", "<br>")


			
			
            
			
			Response.flush
			
            select case md_split_cspan
            case 3
            etal = 2
            case 12
            etal = 11
            case else
            etal = 0
            end select
          
			
			
			'*******************'
			'*** Main SQL ******'
			'*******************'


           
           
            
			
			strSQL = "SELECT j.id AS jid, t.tjobnavn, t.tjobnr, j.jobnr AS jnr, j.jobnavn, j.jobans1, j.jobans2, "_
			&" COALESCE(sum(t.timer), 0) AS sumtimer, "& vgtTimePris &" t.tmnr, t.tmnavn, t.timepris, t.kostpris, t.tdato, m.mnavn, m.mnr, m.mid, m.init, "

           
			strSQL = strSQL &" j.jobTpris, j.fakturerbart, j.budgettimer, j.ikkebudgettimer, j.fastpris, j.usejoborakt_tp, m.medarbejdertype AS mtype "& aktSQLFields &","_
			&" k.kkundenavn, k.kkundenr, k.adresse, k.postnr, k.city, k.telefon, t.kurs AS oprkurs, t.valuta, "_
			&" m2.mnavn AS m2mnavn, m2.init AS m2init, m2.mnr AS m2mnr, m3.mnavn AS m3mnavn, m3.init AS m3init, m3.mnr AS m3mnr "
			
            if cint(vis_kpers) = 1 then
            strSQL = strSQL & ", kp.navn AS kontaktperson"
            end if

            if cint(vis_jobbesk) = 1 then
            strSQL = strSQL & ", j.beskrivelse"
            end if

			if cint(vis_enheder) = 1 then
			strSQL = strSQL & ", a2.faktor, COALESCE(SUM(t.timer * a2.faktor), 0) AS enheder"
			end if
			
            if cint(vis_medarbejdertyper) = 1 AND cint(vis_medarbejdertyper_grp) = 0 then
			strSQL = strSQL & ", mty.type AS mtypenavn"
            strSQL = strSQL & ", mty.sostergp"
			end if
            
           if cint(vis_medarbejdertyper_grp) = 1 then
            strSQL = strSQL & ", mtygrp.navn AS mtypegrpnavn"
            end if
        
           
           
			
			strSQL = strSQL &" FROM medarbejdere m "
            strSQL = strSQL &" LEFT JOIN job j ON ("& sqlJobKri &") "


            if len(trim(aktSQLTables)) <> 0 then 'udpecificer p� aktiviteter
            strSQL = strSQL &" LEFT JOIN aktiviteter AS a ON (a.job = j.id)"
            end if

           

			strSQL = strSQL &" LEFT JOIN kunder k ON (k.kid = j.jobknr)"
			
            if cint(vis_kpers) = 1 then
            strSQL = strSQL &" LEFT JOIN kontaktpers AS kp ON (j.kundekpers = kp.id)"
            end if

            strSQL = strSQL &" LEFT JOIN timer t ON ( "& aktjobTimerWhere &" AND t.tmnr = m.mid AND "& whrClaus &" AND t.tdato BETWEEN '"& sqlDatoStart &"' AND '"& sqlDatoSlut&"' "& strSQLgk &" "& strSQLaktNavnKri &")"


			strSQL = strSQL &" LEFT JOIN medarbejdere m2 ON (m2.mid = j.jobans1)"_
			&" LEFT JOIN medarbejdere m3 ON (m3.mid = j.jobans2)"
			'&" LEFT JOIN valutaer v ON (v.id = t.valuta)"
			
			if cint(vis_enheder) = 1 then 'OR cint(visfakbare_res) = 1 then
			strSQL = strSQL &" LEFT JOIN aktiviteter a2 ON (a2.id = t.taktivitetid)"
			end if

             if cint(vis_medarbejdertyper) = 1 AND cint(vis_medarbejdertyper_grp) = 0 then
             strSQL = strSQL &" LEFT JOIN medarbejdertyper AS mty ON (mty.id = m.medarbejdertype)" 'OR mty.sostergp = m.medarbejdertype)
             end if

            if cint(vis_medarbejdertyper_grp) = 1 then
             strSQL = strSQL &" LEFT JOIN medarbtyper_grp AS mtygrp ON (mtygrp.id = m.medarbejdertype_grp)"
            end if

            'if cint(vis_restimer) = 1 then 'OR cint(visfakbare_res) = 1 then
			'strSQL = strSQL &" LEFT JOIN ressourcer_md AS rmd ON (rmd.jobid = j.id AND rmd.medid = m.mid AND ((rmd.md >= MONTH('"& sqlDatoStart &"') AND aar = YEAR('"& sqlDatoStart &"')) "& orandval &" (rmd.md <= MONTH('"& sqlDatoSlut &"') AND aar = YEAR('"& sqlDatoSlut &"'))))"
			'end if
            

			if cint(visMedarbNullinier) = 0 then
            strSQL = strSQL &" WHERE ("& sqlMedKri &") "
            else 
			strSQL = strSQL &" WHERE ("& medarbSQlKri &") "
			end if

            strSQL = strSQL & aktSQLWhere &" "& aktidsSQLkriUse &" GROUP BY "& grpbySQL &" ORDER BY " & orderBySQL 

            '("& sqlJobKri &")
            'sqlMedKri
			
			'Response.write "<br><br><hr>"
			'Response.write "sqlMedKri: "& sqlMedKri &"<br>"
			'Response.write " sqlJobKri: "& sqlJobKri &"<br>"
			'Response.write "aktSQLFields: "& aktSQLFields &"<br>"
			'Response.write " aktSQLTables: "& aktSQLTables &"<br>"
			'Response.write "aktSQLWhere: "& aktSQLWhere &"<br>"
			'Response.write "aktjobTimerWhere: "& aktjobTimerWhere &"<hr>"
			'Response.Write "medarbSQlKri" & medarbSQlKri

            'if session("mid") = 1 then
			'response.write "<br><br><br><br><br><br><br><br><br><br>"_
            '&"<br><br><br><br><br>MAIN:" & strSQL & "<br><br>"
			'response.flush
			'end if
           

         
            jobmedtimerHigh = 0
			oRec.open strSQL, oConn, 3 
			while not oRec.EOF 
						
						
                        jobIsWrt = jobIsWrt & " a.job <> "& oRec("jid") &" AND"
						
						'Response.write x &", " & v &"<br>"
						
						'*** Job og medarb data ****
						jobmedtimer(x,0) = oRec("jid")
						jobmedtimer(x,1) = oRec("jobnavn")

                        if isNull(oRec("sumtimer")) <> true then
						jobmedtimer(x,3) = formatnumber(oRec("sumtimer"), 2)
                        else
                        jobmedtimer(x,3) = 0
                        end if
		
        				jobmedtimer(x,2) = oRec("mnavn") 
                        jobmedtimer(x,37) = oRec("mnr")
                            
                            if len(trim(oRec("init"))) <> 0 then
                            jobmedtimer(x,39) = "<br>" & oRec("init")
                            else
                            jobmedtimer(x,39) = ""
                            end if
                        

                        if cint(vis_medarbejdertyper) = 1 then
                        jobmedtimer(x,4) = oRec("mtype") 
						else
                        jobmedtimer(x,4) = oRec("mid")
                        end if
						
						'jobmedtimer(x,5) = oRec("restimer")
						jobmedtimer(x,6) = oRec("jnr")
						
                        if (cint(upSpec) = 1 AND jobid <> 0) then
                        jobmedtimer(x,38) = 1
                        else
                        jobmedtimer(x,38) = ja
                        end if
                       
                        'jobid = 1
						'*** Uspecificering af aktiviteter ***
						if cint(ja) <> 0 then
							
							jobmedtimer(x,12) = oRec("aid")
							jobmedtimer(x,13) = oRec("anavn")
							thisAktFakbar = oRec("afakbar")
							
							jobmedtimer(x,34) = oRec("bgr")
							
			                visning = 1
			                call akttyper(thisAktFakbar, visning)
			                jobmedtimer(x,14) = akttypenavn
			                
			                jobmedtimer(x,32) = oRec("fase")

                            jobmedtimer(x,40) = oRec("brug_fasttp")
                            
                        
                        else
						    
                            jobmedtimer(x,12) = 0
                            
                            jobmedtimer(x,13) = ""
                            jobmedtimer(x,14) = ""
                            jobmedtimer(x,34) = 0

						end if
						
						jobmedtimer(x,15) = oRec("mtype")
						
						
						'*** Antal Medarbejdere / Navne ***
						if instr(strMidsK, "#"& jobmedtimer(x,4) &"#") = 0 then
						strMidsK = strMidsK & ",#"& jobmedtimer(x,4) &"#"
							

                            'Response.Write strMidsK & "v: "& v &"<br>"

							'Redim preserve medarb(v)
							'Redim preserve medarbnavnognr(v)
							medarb(v) = jobmedtimer(x,4)

                            if cint(vis_medarbejdertyper) = 1 OR cint(vis_medarbejdertyper_grp) = 1 then
                            
                                    if cint(vis_medarbejdertyper_grp) = 1 then
                                        'if NOT isNull(oRec("mtypegrpnavn")) = true then                
                                        medarbnavnognr(v) = oRec("mtypegrpnavn")
                                        'else
                                        'medarbnavnognr(v) = "(ingen gruppe)"
                                        'end if
                                    else 
                                    medarbnavnognr(v) = oRec("mtypenavn")

                                                if oRec("sostergp") <> oRec("mtype") then
                                                strSQLsoster = "SELECT type FROM medarbejdertyper WHERE id = "& oRec("sostergp")
                                                oRec6.open strSQLsoster, oConn, 3
                                                mtypSosterNavn = ""
                                                if not oRec6.EOF then
                                                mtypSosterNavn = oRec6("type")
                                                end if
                                                oRec6.close

                                                if mtypSosterNavn <> "" then
                                                medarbnavnognr(v) = medarbnavnognr(v) & "<br> ("& mtypSosterNavn &")"
                                                end if

                                                end if
                                    end if

                            medarbnavnognr_short(v) = medarbnavnognr(v) 'left(medarbnavnognr(v), 10)

                            else
							


                            medarbnavnognr(v) = jobmedtimer(x,2) &" ["& jobmedtimer(x,39) & "]" '("& jobmedtimer(x,37) &")"
                            
                            if (lto = "wwf" OR lto = "wwf2") then
                            medarbnavnognr_short(v) = jobmedtimer(x,39)
                            else
                                
                                if cint(md_split_cspan) = 1 OR media = "print" then
                                thisM = left(jobmedtimer(x,2), 10) &" "& jobmedtimer(x,39) '<br>("& jobmedtimer(x,37) &")
                                else
                                thisM = left(jobmedtimer(x,2), 35) &" "& jobmedtimer(x,39) '<br>("& jobmedtimer(x,37) &")
                                end if
                            
                            
                            medarbnavnognr_short(v) = thisM
                            end if

                            
                            end if
							
                            v = v + 1
							
						else
						strMidsK = strMidsK
						end if
						
						
						'*** Jobtype ***
						if oRec("fastpris") = 1 then
						    'if oRec("usejoborakt_tp") <> 1 then
						    jobmedtimer(x,33) = 1 
						    jobmedtimer(x,8) = " - Fastpris"
						    'else
						    'jobmedtimer(x,8) = " - Fastpris (akt. ~ tiln�rm. timepris)"
						    'jobmedtimer(x,33) = 2
						    'end if
						else
						jobmedtimer(x,33) = 0
						jobmedtimer(x,8) = " - Lbn. timer" 
						end if
						
						
						'** Timepris v. fastpris job 
						timeprisThis = 0
						
						if oRec("budgettimer") <> 0 then
						bdgTim = oRec("budgettimer") + oRec("ikkebudgettimer")
						else
						bdgTim = 0
						end if
						
						if oRec("jobtpris") > 0 then
						jobbelob = oRec("jobtpris")
						else
						jobbelob = 0
						end if
						
						
						'*** Job eller aktiviteter VISES **'
						'** Budget / Forkalk **'
						if cint(jobmedtimer(x,38)) <> 0 then 'if oRec("usejoborakt_tp") <> 0 then '** Bruger akt forkalk 
						    '** budget **'
						    if oRec("aktbudgetsum") <> 0 then
						    jobmedtimer(x,18) = oRec("aktbudgetsum") 'oRec("abudgetsum")
						    else
						    jobmedtimer(x,18) = 0
						    end if
    						
					   
					        if oRec("abudgettimer") <> 0 then
					        jobmedtimer(x,19) = oRec("abudgettimer")
					        else
					        jobmedtimer(x,19) = 0
					        end if
						   
                            
                            jobmedtimer(x,35) = oRec("antalstk")
							
                            						    
						    select case oRec("bgr") 
						    case 0 'Ingen
						        sumbgr = jobmedtimer(x,19)
						    case 1 'Timer
						        sumbgr = jobmedtimer(x,19)
						    case 2 'stk
						        sumbgr = jobmedtimer(x,35)
						    end select
						
						else
						    '** Job forkalk ***'
						
						    '** budget **'
						    jobmedtimer(x,18) = jobbelob
						    '** timer **'
						    jobmedtimer(x,19) = bdgTim
						end if
						
						jobmedtimer(x,27) = oRec("oprkurs")
						
						
						
						'*** Timepriser bruges p� forkalk og hvis det er fastpris job ***'
						'if len(trim(jobmedtimer(x,3))) <> 0 AND jobmedtimer(x,3) <> 0 then
                        'realtimerTP = jobmedtimer(x,3)
                        'else
                        'realtimerTP = 1
                        'end if

                        if bdgTim > 0 then
						    '** job er basis **'
							if oRec("usejoborakt_tp") <> 1 then
                                timeprisThis = (jobbelob / bdgTim) 'realtimerTP '** Skal v�re real timer
						    else
						        '** Akt er basis og det er aktiviteter der vises **
						        if cint(jobmedtimer(x,38)) <> 0 then
						        timeprisThis = (jobmedtimer(x,18)/sumbgr) '** enhedspris p� akt
						        else
    						    timeprisThis = (jobbelob / bdgTim) 'realtimerTP
						        end if
						    end if
						else
						timeprisThis = (jobbelob / 1)
						end if
						
						
						
						            '*** Vis oms�tning *******'
						            if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
                                    '1: Oms�tning
                                    '2: kostpriser
            							
							                '*** Timepriser ***'
							                jobmedtimer(x,7) = 0
                                            
                                            '*** �NDRET pr. 1.9.2011 ** Viser altid den indtastede timepris pr. medarbejer.
                                            '*** Da det er blevet lettere at �ndre medarbejdertype timepris, og da det giver et mere rigtigt billede.
                                            '*** af at der er indtasts 20 dyre timer til en timepris, selvom den samlede bruttooms�tning ikke er s� h�j
                                            '*** Herved vil nedskrivning p� job ogs� nemmere komme tilsyne og det vil animere til at s�tte den rigtige timepris p� medarbejer typen p� det enkelte job.
                                            
                                            '*** og s� alligevel, indtil videre beholder vi den gamle beregning: 1 / :999 **'

                                            '** Fastpris **'
							                if cint(oRec("fastpris")) = 999 then '1
                							
                							
                							
						                        '** Bel�b **'jobmedtimer(x,7)
						                        belobThis = timeprisThis * oRec("sumtimer")
						                        call beregnValuta(belobThis,jobmedtimer(x,27),100)
						                        jobmedtimer(x,7) = valBelobBeregnet
                    							
						                        '*** Timepris ***'
						                        call beregnValuta(timeprisThis,jobmedtimer(x,27),100)
						                        jobmedtimer(x,10) = valBelobBeregnet
                							
                							
							                else
							                '** Bel�b **'
							                '*** Bliver beregnet i SQL kald ***'
							                jobmedtimer(x,7) = oRec("vaegtettimepris") 
                							
							                    '*** Timepris / Kostprise ***'
                                                if cint(visfakbare_res) = 1 then
							                    call beregnValuta(oRec("timepris"),jobmedtimer(x,27),100)
							                    jobmedtimer(x,10) = valBelobBeregnet
                							
                                                else 'kostpris

                                                call beregnValuta(oRec("kostpris"),jobmedtimer(x,27),100)
							                    jobmedtimer(x,10) = valBelobBeregnet

                                                end if


							                end if
            						
						            else
            							
                                        '*** bruge kun pseudo da, der ikke beregnes timepris ved vis kun timer
							            '*** Timepris ***'
							            call beregnValuta(oRec("timepris"),jobmedtimer(x,27),100)
							            jobmedtimer(x,10) = valBelobBeregnet
            						
						            end if
						
						
                        
                        jobmedtimer(x,11) = oRec("fastpris")
						
						if cint(vis_enheder) = 1 then
						jobmedtimer(x,25) = oRec("enheder")
						else
						jobmedtimer(x,25) = 0
						end if
						
						
							
							if cint(jobmedtimer(x,38)) <> 0 then '*** Udspec p� aktiviteter 
							jora_id = oRec("aid")
							else
							jora_id = oRec("jnr")
							end if
							
		                '*** 16 Sumtimer p� job ialt / 17 Oms�tning ***'
                        
		                if instr(strJobnbs, "#"& jora_id &"#") = 0 OR cint(csv_pivot) = 1 then
		                
                        'Response.Write "ja:" & ja & " navn: "& jobmedtimer(x,13) & " -- " & strJobnbs & " jora_id: " & jora_id & "<br>"
                        
                        strJobnbs = strJobnbs & ",#"& jora_id &"#"
		                
                								
							            if isNull(oRec("sumtimer")) <> true then
							            jobmedtimer(x,16) = oRec("sumtimer")
        								
								            if cint(vis_enheder) = 1 then
							                jobmedtimer(x,26) = oRec("enheder")
							                end if
            								
									            if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
									            jobmedtimer(x,17) = jobmedtimer(x,7) 'oRec("vaegtettimepris")
									            end if
            								
						                else
								            jobmedtimer(x,16) = 0
            								
            								
								            if cint(vis_enheder) = 1 then
							                jobmedtimer(x,26) =  0
							                end if
            								
									            if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
									            jobmedtimer(x,17) = 0
									            end if
							            end if
							            lastj = x
								
								
							else
							jobmedtimer(lastj,16) = (jobmedtimer(lastj,16) + oRec("sumtimer"))
							
							        if cint(vis_enheder) = 1 then
							        jobmedtimer(lastj,26) = (jobmedtimer(lastj,26) + oRec("enheder"))
							        end if
        								
								        if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
									        'if oRec("fastpris") = 1 then
									        jobmedtimer(lastj,17) = jobmedtimer(lastj,17) + jobmedtimer(x,7)
									        'else
									        'jobmedtimer(lastj,17) = (jobmedtimer(lastj,17) + jobmedtimer(x,7))
									        'end if
								        end if
							
							end if
						
						
                            if media <> "print" then
							jobmedtimer(x,20) = oRec("kkundenavn") & " ("& oRec("kkundenr") &")"
		                    else
                            jobmedtimer(x,20) = left(oRec("kkundenavn"), 20) & " ("& oRec("kkundenr") &")"
                            end if


        					jobmedtimer(x,21) = oRec("adresse")
							jobmedtimer(x,22) = oRec("postnr")
							jobmedtimer(x,23) = oRec("city") 
							jobmedtimer(x,24) = oRec("telefon")		
				            
				            if len(trim(oRec("m2mnavn"))) <> 0 then
				            jobmedtimer(x,28) = oRec("m2mnavn")  '&" ("&oRec("m2mnr")&")"
				            jobmedtimer(x,29) = oRec("m2init") 
				            else
				            jobmedtimer(x,28) = ""
				            jobmedtimer(x,29) = ""
				            end if
				            
				            if len(trim(oRec("m3mnavn"))) <> 0 then
				            jobmedtimer(x,30) = oRec("m3mnavn") '&" ("&oRec("m3mnr")&")"
				            jobmedtimer(x,31) = oRec("m3init")
				            else
				            jobmedtimer(x,30) = ""
				            jobmedtimer(x,31) = ""
				            end if 
				            
				            
				            jobmedtimer(x,36) = oRec("tdato")
                
                            if cint(vis_kpers) = 1 then
                            jobmedtimer(x,41) = oRec("kontaktperson") 
                            end if


                            if cint(vis_jobbesk) = 1 then
                            jobmedtimer(x,42) = oRec("beskrivelse")
                            end if

                            'if cint(vis_restimer) = 1 AND isNull(oRec("restimer")) <> true then
                            'jobmedtimer(x,37) = oRec("restimer")
                            'else
                            'jobmedtimer(x,37) = 0
                            'end if
				            
                            'jobmedtimer(x,38) = oRec("sumtimerPrev") 

			x = x + 1
            jobmedtimerHigh = x 
			'response.flush		
			oRec.movenext
			wend
			oRec.close 
			

            next 'ja
			
			'Response.write "x:" & x
			

            'for x = 0 to UBOUND (jobmedtimer, 1)
            'Response.Write jobmedtimer(x,4) & "<br>"

            'next

            'timeB = now
            'loadtime = datediff("s",timeA, timeB)
            'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>directexp: "& directexp &" <br>loadtid: " & loadtime
            
                

            'Response.end
            
            if cint(directexp) = 1 then
            Response.write "<br><br><br><div class=load>G�r din csv. fil klar, vent et �jeblik. Der g�r mellem 5 og 10 sekunder.</div>" 
            end if

            Response.flush
                
			
			'if cint(jobid) = 0 then
			jobaktOskrift = ""
			bgtd = "#ffffff"
			'else
			'jobaktOskrift = "Aktivitet"
			'bgtd = "#ffffff"
			'end if
			
			strJobLinie_top = ""
			strJobLinie = ""
			strJobLinie_total = ""
			timerTotSaldoGtotal = 0
			
			vCntVal = 0
			
            '*** Vis redakt�r = 1 / 0 hvis medarbejdertyper er sl�et til ****'
            if cint(directexp) <> 1 then 'direkte excel / print

            if media <> "print" then

            if vis_redaktor = 1 then
            strJobLinie_top = "<form action='"&strLink&"&func=dbupdatelist' id='redaktor' method=post>"
            else
            strJobLinie_top = ""
            end if

			strJobLinie_top = strJobLinie_top & "<br><br><table cellspacing=0 cellpadding=0 border=0>"
			strJobLinie_top = strJobLinie_top & "<tr><td Style='border:0px #8caae6 solid; padding:10px;' bgcolor='#ffffff'>"
			'strJobLinie_top = strJobLinie_top & "<h3>Timeforbrug, Oms�tning og Ressourcetimer.</h3>"
			
			if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
                if cint(visfakbare_res) = 1 then
			    strJobLinie_top = strJobLinie_top & "Realiserede fakturerbare timer og oms�tning."
                else
                strJobLinie_top = strJobLinie_top & "Realiserede fakturerbare timer og kost."
                end if
			else
                strJobLinie_top = strJobLinie_top & "Realiserede timer ialt."
            end if


            end if 'print

			if media <> "print" then
            strJobLinie_top = strJobLinie_top & " <br>Alle timer og bel�b er afrundet til 0 decimaler.<br>&nbsp;"
            else
            strJobLinie_top = strJobLinie_top & "<h4>Timeout - Grandtotal<br><span style=""font-size:9px;"">"& now &" - Periode: "& formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1) &"</span></h4>"
            end if

            strJobLinie_top = strJobLinie_top &  "<table cellspacing=0 cellpadding=0 border=0 bgcolor='#ffffff'>"
			


        end if'    if cint(directexp) <> 1 then 




        call medarboSkriftlinje
	    
        if cint(directexp) <> 1 then 
        'if cint(upSpec) <> 1 then
        strJobLinie_top = strJobLinie_top & strMedarbOskriftLinie
		'end if		
                            

		end if'    if cint(directexp) <> 1 then 
				
				'************************************************************'
				'*** Udskriver en ny linie med de valgte Job / Aktiviteter
				'************************************************************'
				'j = 0


                'LastMd = 1

				x = 0
				p = 0
				lastV = -1
				lastjid = "#0#"
                'lastWrtMdx = -1
                lastWrtX = -1
				
                lastWrtMd = -1
                thisWrtMd = 0
                nextWrtMd = 0

                thisWrtM = 0
                lastWrtM = 0
                nextWrtM = 0
                
                thisWrtJ = 0
                lastWrtJ = 0
                nextWrtJ = 0    
                lastFase = ""

                udSpecFirst = 0
                'lastJM = "0_0"

                firtsEfelt = 0
                'lastMidstrId = 0

				for x = 0 to jobmedtimerHigh - 1 'UBOUND(jobmedtimer, 1)
				                
                        
                      
				
				                if cint(jobmedtimer(x,38)) = 0 then '** Job
								jobaktId = jobmedtimer(x,0) & "_0"
								jobaktNr = jobmedtimer(x,6)
								jobaktNavn = jobmedtimer(x,1)
								jobaktType = jobmedtimer(x,8)
								
								jobans = jobmedtimer(x,28) 
								if len(trim(jobmedtimer(x,30))) <> 0 then
								jobans = jobans & ", "& jobmedtimer(x,30)
								end if
								
								
								else '*** Udspecificer Akt.
								
								jobaktId = jobmedtimer(x,0) & "_"& jobmedtimer(x,12)
								jobaktNr = jobmedtimer(x,12)
								jobaktNavn = jobmedtimer(x,13)
								jobaktType = jobmedtimer(x,14)
								jobaktFase = jobmedtimer(x,32)
								
								jobans = ""
								
								end if
								
                                
                                'Response.Write "jobaktNavn: " & jobaktNavn & "<hr>"
                           
                                  

									
									
									if (instr(lastjid, "#"&jobaktId&"#") = 0 AND len(trim(jobaktNavn)) <> 0 AND cint(csv_pivot) = 0) OR (cint(csv_pivot) = 1 AND jobmedtimer(x,16) <> 0) then 
                                    
                                    'AND lastMidstrId <> jobmedtimer(x,4) AND len(trim(jobmedtimer(x,4))) <> 0) then 'AND jobmedtimer(x,16) > 0


                                    
                                    antaljob = antaljob + 1
									'strJobLinie = strJobLinie & "</tr><tr>"
                                    
                                    if cint(directexp) <> 1 then                                

                                         if cint(upSpec) <> 1 then

                                                   select case antaljob
                                                   case (sideskiftlinier*1),(sideskiftlinier*2),(sideskiftlinier*3),(sideskiftlinier*4),(sideskiftlinier*5),(sideskiftlinier*6),(sideskiftlinier*7), (sideskiftlinier*8), (sideskiftlinier*9), (sideskiftlinier*10)
                         
                                                    
                                                    strJobLinie = strJobLinie & "<tr style='page-break-before:always'><td colspan=1000>&nbsp;</td></tr>"
                                                    strJobLinie = strJobLinie & strMedarbOskriftLinie '&"</tr>" 
                                                                   
                        
                                                 end select

                                         end if
                                                
                                    
                                    end if '   if cint(directexp) <> 1 then


									


									    if jobmedtimer(x,38) <> 0 then 
                                        '************************************************************************************
                                        '* job eller akt. udspec
                                        '************************************************************************************
                

                                                '*** AKT ***'
                                                if len(trim(jobmedtimer(x,0))) <> 0 AND (lastJob <> jobmedtimer(x,0)) then 

                                                            if udSpecFirst = 0 then
                                            
                                                                if cint(directexp) <> 1 then
                                                                strJobLinie = strJobLinie & "</tr><tr><td colspan=400 style='padding:20px 10px 2px 2px;' bgcolor=#FFFFFF>"_
    									                        &"Der er valgt udspecificering p� f�lgende job:</td>"
                                                                end if

                                                                    subbudgettimer = 0
                                                                    subbudget = 0
                                                                    TimeprisFaktiskSub = 0
                                                                    saldoRestimerSub = 0
                                                                    saldoJobSub = 0
                                                                    restimerSubJob = 0
                                                                    subtotaljboTimerIalt = 0
                                                                    subJobEnh = 0
                                                                    subtotaljobOmsIalt = 0
                                                                    subdbialt = 0
                                                                    timerTotSaldoSubGtotal = 0
            
                                                                    restimerSubGtotalJob = 0
                                                                    enhederPrevSaldoSub = 0
                                                                    subJobEnh = 0  
                                                                    enhederGSub = 0


                                                                for v = 0 to v - 1

                                                                 subMedabRestimer(v) = 0 
                                                                 subMedabTottimer(v) = 0
                                                                   subMedabTotenh(v) = 0
                                                                   omsSubTot(v) = 0

                                                                next



                                                            else
                                                
                                                                if cint(directexp) <> 1 then                                    
                                                
                                                                call subTotaler_gt

                                                                strJobLinie = strJobLinie & "<tr><td colspan=400 style='padding:20px 10px 2px 2px;'>"_
    									                        &"&nbsp;</td></tr>"
                                                 
                                                                end if'    if cint(directexp) <> 1 then

                                               
                                                            end if 'udSpecFirst

                                                

                                                
                                                    if cint(directexp) <> 1 then
    									                
                                                                'kunde
                                                                strJobLinie = strJobLinie & "</tr><tr><td style='padding:30px 10px 2px 2px; background-color:#F7F7F7; border-top:1px #CCCCCC solid;'>"_
    									                        &"<span style='color:#000000; font-size:10px;'>"& jobmedtimer(x,20) &"</span><br>"
        
        
                                                                 '**Vis kontaktperson
                                                                if cint(vis_kpers) = 1 then
                                                                    if len(trim(jobmedtimer(x,41))) <> 0 then
                                                                    strJobLinie = strJobLinie & "<span style='color:#999999; font-size:9px;'>"
									                                strJobLinie = strJobLinie & jobmedtimer(x,41) &"</span><br>"
                                                                    end if
                                                                end if 

                                                                'Jobnavn
                                                                strJobLinie = strJobLinie &"<b>"& jobmedtimer(x,1) &"</b> ("& jobmedtimer(x,6) &")"

                                                                 '**Vis jobbeskrivelse
                                                                if cint(vis_jobbesk) = 1 then
                                                                    if len(trim(jobmedtimer(x,42))) <> 0 then
                                                                    strJobLinie = strJobLinie & "<br><span style='color:#5582d2; font-size:9px; display:block; width:350px; white-space:normal;'>"
									                                strJobLinie = strJobLinie & left(trim(jobmedtimer(x,42)), 250) &"</span>"
                                                                    end if
                                                                end if 
                  
                                                                strJobLinie = strJobLinie &"</td>"

                                                                strJobLinie = strJobLinie &"<td class=lille valign=bottom style='border-top:1px #CCCCCC solid; background-color:#F7F7F7;'>Budget</td>"

                                                                if cint(visPrevSaldo) = 1 then
                                                                strJobLinie = strJobLinie &"<td class=lille valign=bottom style='border-top:1px #CCCCCC solid; background-color:#F7F7F7;'>Real. timer<br> f�r valgte periode</td>"
                                                                end if
                                                
                                                                strJobLinie = strJobLinie &"<td class=lille valign=bottom style='border-top:1px #CCCCCC solid; background-color:#F7F7F7;'>Real. timer<br>i periode</td>"
                                                                
                                                                select case lto
                                                                case "mmmi", "intranet - local"
                                                                case else
                                                                strJobLinie = strJobLinie &"<td class=lille valign=bottom style='border-top:1px #CCCCCC solid; background-color:#F7F7F7;'>Real. %</td>"
                                                                end select


                                                                if cint(visPrevSaldo) = 1 then
                                                                strJobLinie = strJobLinie &"<td class=lille valign=bottom style='border-top:1px #CCCCCC solid; background-color:#F7F7F7;'>Real. timer<br> ialt</td>"
                                                                end if
                                                
                                                                'if cint(visPrevSaldo) = 1 then
                                                                 'strJobLinie = strJobLinie & "<td style='border-top:1px #CCCCCC solid;'>&nbsp;</td><td style='border-top:1px #CCCCCC solid;'>&nbsp;</td>"
                                                                'end if        
                                                            
                                                                            if cint(upSpec) = 1 then
                                                                            '*** Overskrifter og medarb ***'
				                                                                    for v = 0 to v - 1
						                                            

                                                                                            select case right(v, 1)
                                                                                                    case 0,2,4,6,8
                                                                                                    bgthis = "#FFFFFF"
                                                                                                    case else
                                                                                                    bgthis = "#EFf3FF"
                                                                                                    end select
						
						                                                                    '** Medarb 1 Row ***
						                                                                    strJobLinie = strJobLinie & "<td colspan="& md_split_cspan &" "&tdstyleTimOms1&" bgcolor='"& bgthis &"'><span style='color:#000000; font-size:9px;'>"& medarbnavnognr_short(v) &"</span></td>"
	                                                              
						
						
				                                                                    next

                                                                            else
                                                                            strJobLinie = strJobLinie & "<td colspan=117 style='padding:30px 10px 2px 2px; border-top:1px #CCCCCC solid;' bgcolor=#FFFFFF>&nbsp;</td>"
                                                

                                                                            end if


                                                    udSpecFirst = 1
                                                    lastJob = jobmedtimer(x,0)

                                                    end if 'if cint(directexp) <> 1 then


    									        end if 'len(trim(jobmedtimer(x,0))) <> 0 AND (lastJob <> jobmedtimer(x,0))
									    

                                                
                                                if cint(directexp) <> 1 then
                                    

									            '*** Fase ***'
                                                select case lto
                                                case "cisu"

                                                case else

									                if len(trim(jobaktFase)) <> 0 AND (lastFase <> jobaktFase) then 
    									            strJobLinie = strJobLinie & "</tr><tr><td colspan=120 class=lille style='padding:2px; padding-right:10px; border-top:1px #CCCCCC solid;' bgcolor=#d6dff5>"_
    									            &" fase: <b>"& jobaktFase &"</b></td>"

                                                    lastFase = jobaktFase
    									            end if 
                                                end select 'lto                                    





                                                '*********************************'
                                                '*** job / aktnavn '*****   
    									        '*** timer og job **'
                                                '*********************************'
                                               
									            strJobLinie = strJobLinie & "</tr><tr><td style='padding:2px; padding-right:10px; border-top:1px #CCCCCC solid; width:250px; white-space:nowrap;' bgcolor="& bgtd &">"
                                                strJobLinie = strJobLinie & jobaktNavn 
                                                strJobLinie = strJobLinie & "<br><span style='color:#6CAE1C; font-size:9px;'> ("& jobaktType &")</span></td>" 
									    
									   
    									        end if'    if cint(directexp) <> 1 then 



    									
    									        ekspAkttype = jobaktType 'replace(jobaktType,"<font color=green>", "")
    									
									            expTxt = expTxt &"xx99123sy#z"
		                                        'kunde
        							            expTxt = expTxt & jobmedtimer(x,20) &";"
        
                                                'Kontaktperson
                                                if cint(vis_kpers) = 1 then
                                                expTxt = expTxt & jobmedtimer(x,41)&";"
                                                end if     

									            expTxt = expTxt & jobmedtimer(x,1) &";"& jobmedtimer(x,6) &";"

                                               
                                                'jobbesk
                                                if cint(vis_jobbesk) = 1 then
                                                call htmlparseCSV(jobmedtimer(x,42))
                                                expTxt = expTxt & htmlparseCSVtxt &";"
                                                end if
                                                
                                                select case lto
                                                case "cisu"
                                                expTxt = expTxt & jobaktNavn &";"& ekspAkttype &";"
                                                case else
                                                expTxt = expTxt & jobaktFase &";"& jobaktNavn &";"& ekspAkttype &";"
									            end select

                        
                                                'Jobansvarlig, Jobejer Init
                                                select case lto
                                                case "cisu"
                                               
                                                case else

                                                expTxt = expTxt & jobmedtimer(x,28)&";"
    									        expTxt = expTxt & jobmedtimer(x,29)&";"
    									        expTxt = expTxt & jobmedtimer(x,30)&";"
    									        expTxt = expTxt & jobmedtimer(x,31)&";"
									            end select
                                            
                                               

                                                expTxt = expTxt & formatnumber(jobmedtimer(x,19),2)&";"

			                            else

                                                '**** JOB *****
    	                                        if cint(directexp) <> 1 then 								
									            strJobLinie = strJobLinie & "</tr><tr><td style='padding:2px; padding-right:10px; border-top:1px #CCCCCC solid; width:250px; white-space:nowrap;' bgcolor="& bgtd &">" 
        
                                                strJobLinie = strJobLinie & "<span style='color:#5C75AA; font-size:9px;'>"
									            
                                                strJobLinie = strJobLinie & jobmedtimer(x,20) &"</span><br>"
                                               
                                                '**Vis kontaktperson
                                                if cint(vis_kpers) = 1 then
                                                    if len(trim(jobmedtimer(x,41))) <> 0 then
                                                    strJobLinie = strJobLinie & "<span style='color:#999999; font-size:9px;'>"
									                strJobLinie = strJobLinie & jobmedtimer(x,41) &"</span><br>"
                                                    end if
                                                end if                                        

        
                                                if media <> "print" then
                                                strJobLinie = strJobLinie & "<a href='joblog_timetotaler.asp?FM_jobsog="&jobaktNr&"&FM_udspec=1&nomenu=1&FM_medarb="&thisMiduse&"&FM_fomr="&fomr&"' target='_blank' class='vmenu'>[+] "& jobaktNavn &"</a>"
                                                else
                                                strJobLinie = strJobLinie & left(jobaktNavn, 25) 
                                                end if

                                                strJobLinie = strJobLinie &" ("&jobaktNr&") "

                                                if media <> "print" then
                                                strJobLinie = strJobLinie & "<span style='color:green; font-size:9px;'>"& jobaktType &"</span>" 
                                                end if

                                                if media <> "print" then
									                if (len(trim(jobmedtimer(x,28))) <> 0 OR len(trim(jobmedtimer(x,30))) <> 0) AND lto <> "wwf" then
     									            strJobLinie = strJobLinie & "<br><span style='color:#999999; font-size:9px;'>"& jobans & "</span>"
     									            end if
     									        end if

                                                '**Vis jobbeskrivelse
                                                if cint(vis_jobbesk) = 1 then
                                                    if len(trim(jobmedtimer(x,42))) <> 0 then
                                                    strJobLinie = strJobLinie & "<br><span style='color:#5582d2; font-size:9px; display:block; width:350px; white-space:normal;'>"
									                strJobLinie = strJobLinie & left(trim(jobmedtimer(x,42)), 250) &"</span>"
                                                    end if
                                                end if 
     									
									    
    									        strJobLinie = strJobLinie &"</td>"
			                                    end if						    
            

									            expTxt = expTxt &"xx99123sy#z"
                                                'kunde
									            expTxt = expTxt & jobmedtimer(x,20)&";"

                                                'kontaktperson
                                                if cint(vis_kpers) = 1 then
                                                expTxt = expTxt & jobmedtimer(x,41)&";"
                                                end if   

                                                'jobnavn og jobnr
									            expTxt = expTxt & jobaktNavn&";"&jobaktNr&";"
        
                                                'jobbesk
                                                if cint(vis_jobbesk) = 1 then
                                                call htmlparseCSV(jobmedtimer(x,42))
                                                expTxt = expTxt & htmlparseCSVtxt &";"
                                                end if


                                                'Jobansvarlig, Jobejer Init
                                                select case lto
                                                case "cisu"
        
                                                expTxt = expTxt &";"& jobmedtimer(x,8) &";"                                       

                                                case else

                                                expTxt = expTxt &";;"& jobmedtimer(x,8) &";"
                                                expTxt = expTxt & jobmedtimer(x,28)&";"
    									        expTxt = expTxt & jobmedtimer(x,29)&";"
    									        expTxt = expTxt & jobmedtimer(x,30)&";"
    									        expTxt = expTxt & jobmedtimer(x,31)&";"
									            end select
									          
    									
    									
    									
									    
									    
									            '** adr **'
    									        'call htmlparseCSV(jobmedtimer(x,21))
									            '" &htmlparseCSVtxt& "';"&jobmedtimer(x,22)&";"&jobmedtimer(x,23)&";jobmedtimer(x,24)&";"&"
									    
									            expTxt = expTxt & formatnumber(jobmedtimer(x,19),2)&";"
									    
    									        

                                              

    									
    									
									    end if
										
                                        
										
										'*************************************'
										'**** Forkalk. kolonne p� Job / akt **'
										'*************************************'
		                                if cint(directexp) <> 1 then 								

                                                if jobmedtimer(x,38) <> 0 then '** Uspec p� aktiviteter
									            
									                    strJobLinie = strJobLinie & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor='#EFF3FF'>"
									                    select case jobmedtimer(x,34) 
									                    case 0 'Ingen grundlag
									                    strJobLinie = strJobLinie & "" 'formatnumber(jobmedtimer(x,35),2) &" stk.<br>"& formatnumber(jobmedtimer(x,19),2)  
									                    case 1 'timer

                                                        if jobmedtimer(x,19) <> 0 then
									                    strJobLinie = strJobLinie & formatnumber(jobmedtimer(x,19),2)
                                                        else
                                                        strJobLinie = strJobLinie & ""
                                                        end if

									                    case 2 'stk

                                                        if jobmedtimer(x,35) <> 0 then
						                                strJobLinie = strJobLinie & formatnumber(jobmedtimer(x,35),2) &" stk." '<br>"& formatnumber(jobmedtimer(x,19),2) 
                                                        else
                                                        strJobLinie = strJobLinie & ""
                                                        end if
		
        							                    end select
									            
									            else
									            strJobLinie = strJobLinie & "<td class=lille align=right "&tdstyleTimOms2&" bgcolor='#EFF3FF'>"& formatnumber(jobmedtimer(x,19),2) 
									            end if

                                        

									    
									            if cint(vis_enheder) = 1 then
									            strJobLinie = strJobLinie & "<br>"
									            end if
									        
                                    
                                         end if  'if cint(directexp) <> 1 then 



									        if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
									        
                                            if cint(directexp) <> 1 then 
									        strJobLinie = strJobLinie & "<br><span style='color=#000000; font-size:9px;'>"& basisValISO &" "& formatnumber(jobmedtimer(x,18),2) & "</span>"
                                            end if        

                                                   if jobmedtimer(x,38) = 0 then '**Udspec p� aktiviteter
                                                          
                                                                  
                                                   expTxt = expTxt & formatnumber(jobmedtimer(x,18), 2)&";" 
                                                          
                                                   else

                                                   if cint(directexp) <> 1 then 
                                                   strJobLinie = strJobLinie & "</td>"
                                                   end if
                                                   
                                                   expTxt = expTxt & formatnumber(jobmedtimer(x,18),2)&";"
                                                   end if'jobid
        									
									        else
                                            
                                                if cint(directexp) <> 1 then 
									            strJobLinie = strJobLinie & "</td>"
		                                        end if							        

									        end if
									
									        totbudget = totbudget + jobmedtimer(x,18)
                                            subbudget = subbudget + jobmedtimer(x,18)
									        totbudgettimer = totbudgettimer + jobmedtimer(x,19)
                                            subbudgettimer = subbudgettimer + jobmedtimer(x,19)


                                          

									        '**********************************************
									        '**** Prev. Saldo ****'
                                            '**********************************************
                                            if cint(visPrevSaldo) = 1 then

                                            if jobmedtimer(x,38) <> 0 then
                                            strSaldoJobAktKri = "t2.taktivitetId = "& jobmedtimer(x,12) 'jobaktId
                                            else
                                            strSaldoJobAktKri = "t2.tjobnr = '" & jobmedtimer(x,6) & "'"
                                            end if

                                            timerPrevSaldo = 0
                                            enhederPrevSaldo = 0
                                            strSQLprevSaldo =  "SELECT sum(t2.timer) AS sumtimerPrev, SUM(t2.timer * a.faktor) AS enhederPrevSaldo FROM timer t2 LEFT JOIN aktiviteter AS a ON (a.id = t2.taktivitetid) WHERE "_
                                            &" "& strSaldoJobAktKri &" AND "& replace(whrClaus, "t.", "") &" AND (t2.tdato < '"& sqlDatoStart &"') "& strSQLgk &" GROUP BY tjobnr"

                                            'Response.write strSQLprevSaldo
                                            'Response.flush 

                                            oRec4.open strSQLprevSaldo, oConn, 3
                                            if not oRec4.EOF then

                                            timerPrevSaldo = oRec4("sumtimerPrev")
                                            enhederPrevSaldo = oRec4("enhederPrevSaldo")

                                            end if
                                            oRec4.close
                                                

                                               

                                            saldoJobTotal = (saldoJobTotal/1 + timerPrevSaldo/1) 
                                            saldoJobSub = (saldoJobSub/1 + timerPrevSaldo/1)

                                            if cint(directexp) <> 1 then 
											strJobLinie = strJobLinie & "<td class=lille valign=bottom align=right "&tdstyleTimOms2&">"
                                            end if 'if cint(directexp) <> 1 then 


                                            if cint(vis_restimer) = 1 then
                                            call resTimer(jobmedtimer(x,0), jobmedtimer(x,12), 0, 0, 0, md_split_cspan, 1)

                                            if len(trim(restimerThis)) <> 0 then
                                            restimerThisTxt = restimerThis
                                            else
                                            restimerThisTxt = 0
                                            end if

                                            if cint(directexp) <> 1 then 
                                            strJobLinie = strJobLinie & "<span style='color:#999999;'>"& restimerThisTxt &"</span><br>"
                                            end if 'if cint(directexp) <> 1 then 
                                            end if
                                            
                                            if cint(directexp) <> 1 then 
                                            strJobLinie = strJobLinie & formatnumber(timerPrevSaldo, 2) 
                                            end if 'if cint(directexp) <> 1 then  
                                            
                                            if cint(vis_enheder) = 1 then
                        
                                            if cint(directexp) <> 1 then 
                                            strJobLinie = strJobLinie & "<span style='color:#5c75AA; font-size:9px;'><br>enh. " & formatnumber(enhederPrevSaldo, 2) & "</span>"
                                            end if 'if cint(directexp) <> 1 then 

                                            enhederPrevSaldoTot = (enhederPrevSaldoTot/1 + enhederPrevSaldo/1)
                                            enhederPrevSaldoSub = (enhederPrevSaldoSub/1 + enhederPrevSaldo/1) 

                                            end if
                                            
                                            if cint(directexp) <> 1 then 
                                            strJobLinie = strJobLinie &"</td>"
                                            end if 'if cint(directexp) <> 1 then  
                                            
                                            
                                           
                                             expTxt = expTxt & formatnumber(timerPrevSaldo,2) &";"
                                            

                                            if cint(vis_restimer) = 1 then

                                            if len(trim(restimerThis)) <> 0 then
                                            restimerThis = restimerThis
                                            restimerPS = restimerThis
                                            else
                                            restimerThis = 0
                                            restimerPS = 0
                                            end if

                                            expTxt = expTxt & formatnumber(restimerThis, 0) &";"
                                            end if

                                             if cint(vis_enheder) = 1 then
                                             expTxt = expTxt & formatnumber(enhederPrevSaldo,2) &";"
                                             end if



                                            saldoRestimerTotal = (saldoRestimerTotal/1 + restimerThis/1) 
                                            saldoRestimerSub = (saldoRestimerSub/1 + restimerThis/1) 
                                            
                                          
                                            end if 'cint(visPrevSaldo) = 1 






                                     





											'*********************************************************************'
											'*** Timer og Oms�tning kolonne total p� job/akt. i valgte periode ***'
											'*********************************************************************'

                                            
                                          

											jobtimerIalt = 0
											jobOmsIalt = 0
											
											
											totaltotaljboTimerIalt = totaltotaljboTimerIalt + jobmedtimer(x,16) ' + jobtimerIalt i periode
											subtotaljboTimerIalt = subtotaljboTimerIalt + jobmedtimer(x,16) ' + jobtimerIalt i periode

								            if cint(vis_enheder) = 1 then
							                totalJobEnh = totalJobEnh + jobmedtimer(x,26) 
                                            subJobEnh = subJobEnh + jobmedtimer(x,26) 
							                end if
											

                                            if cint(directexp) <> 1 then 
											strJobLinie = strJobLinie & "<td class=lille align=right "&tdstyleTimOms10&" bgcolor='snow'>"
                                            end if 'if cint(directexp) <> 1 then                                     


                                            '** Res timer p� job ***'
                                            if cint(vis_restimer) = 1 then
                                            call resTimer(jobmedtimer(x,0), jobmedtimer(x,12), 0, 0, 0, md_split_cspan, 0)

                                          
                                            if len(trim(restimerThis)) <> 0 then
                                            restimerThis = restimerThis
                                            restimerPer = restimerThis
                                            restimerThisTxt = restimerThis
                                            else
                                            restimerPer = 0
                                            restimerThis = 0
                                            restimerThisTxt = 0
                                            end if

                                            if cint(directexp) <> 1 then 
                                            strJobLinie = strJobLinie & "<span style='color:#999999;'>" & restimerThisTxt &"</span><br>"
                                            end if 'if cint(directexp) <> 1 then 


                                            restimerTotalJob = (restimerTotalJob/1) + (restimerThis/1)
                                            restimerSubJob = (restimerSubJob/1) + (restimerThis/1)

                                            end if
                                            
                                            if cint(directexp) <> 1 then 
                                            '** REAL timer i periode
                                           strJobLinie = strJobLinie & formatnumber(jobmedtimer(x,16), 2)


                                          
											    if cint(vis_enheder) = 1 then
											    strJobLinie = strJobLinie & "<br><span style='color:#5c75AA; font-size:9px;'>enh. "& formatnumber(jobmedtimer(x,26), 2) & "</span>"
						                        end if
											 
		                                    end if 'if cint(directexp) <> 1 then  									

												'** Real. timer i % beregn
                                                if jobmedtimer(x,19) <> 0 then
                                                realTimerProc = formatnumber((jobmedtimer(x,16) / jobmedtimer(x,19) * 100), 0)
                                                realTimerProcTxt = realTimerProc

                                                if realTimerProc > 100 then
                                                realTimerProcWdt = 50
                                                else
                                                realTimerProcWdt = formatnumber(realTimerProc/2, 0)
                                                end if

                                                else
                                                realTimerProc = 0
                                                realTimerProcTxt = ""
                                                end if


                                                expTxt = expTxt & formatnumber(jobmedtimer(x,16), 2)&";"

                                                select case lto
                                                case "mmmi", "intranet - local"
                                                case else
                                                expTxt = expTxt & realTimerProc &";"
                                                end select
												
                                                if cint(vis_restimer) = 1 then
                                                expTxt = expTxt & restimerThis &";"
                                                end if

												if cint(vis_enheder) = 1 then
												expTxt = expTxt &formatnumber(jobmedtimer(x,26), 2)&";"
												end if
											


											if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
											
											'*** Oms�tning / kost.  ***'
											totaltotaljobOmsIalt = (totaltotaljobOmsIalt/1) + (jobmedtimer(x,17)/1) '+ jobOmsIalt
                                            subtotaljobOmsIalt = (subtotaljobOmsIalt/1) + (jobmedtimer(x,17)/1) 
											
											'*** ~ca timepris ved fastpris, aktiviteter grundlag og jobvisning ***
									        if cint(directexp) <> 1 then 
                                            strJobLinie = strJobLinie &"<br><span style='color=#000000; font-size:8px;'>"& basisValISO &" "& formatnumber(jobmedtimer(x,17)/1, 2)&"</span><br>"
											end if 'if cint(directexp) <> 1 then 

												
										    expTxt = expTxt & formatnumber(jobmedtimer(x,17), 2)&";"
											
											'*** Balance ***
											dbal = (jobmedtimer(x,17) - jobmedtimer(x,18))
                    
                                            if cint(directexp) <> 1 then 
											strJobLinie = strJobLinie & "<span style='color=#000000; font-size:8px;'>bal.: "& formatnumber(dbal, 2) &"</span></td>"
											end if 'if cint(directexp) <> 1 then 	

										    expTxt = expTxt & formatnumber(dbal, 2)&";"
											
											dbialt = dbialt + (dbal)
											subdbialt = subdbialt + (dbal)
											
											else
                                                if cint(directexp) <> 1 then 
												strJobLinie = strJobLinie & "</td>"
                                                end if 'if cint(directexp) <> 1 then 
											end if


											
                                            '**************************
                                            '** REAL timer i %
                                            '**************************
                                           
                                            
                                        select case lto
                                        case "mmmi", "intranet - local"
                                        case else
                                           
                                           

                                               if realTimerProc <> 0 then
                                                
                                                    if realTimerProc > 100 then
                                                    realTimerProcBgcol = "crimson"
                                                    procFntCol = "#ffffff"
                                                    else
                                                    realTimerProcBgcol = "yellowgreen"
                                                    procFntCol = "#000000"
                                                    end if

                                                strJobLinie = strJobLinie &"<td "&tdstyleTimOms&" bgcolor='snow' valign=bottom><div style='width:"& realTimerProcWdt &"px; background-color:"& realTimerProcBgcol &"; color:"& procFntCol &"; padding:1px 4px 1px 1px; font-size:8px; text-align:right;'>"& realTimerProcTxt & "</div></td>"
                                               else
                                                strJobLinie = strJobLinie &"<td "&tdstyleTimOms&" bgcolor='snow'>&nbsp;</td>"
                                               end if
                                           
                
                                          end select





                                            '************************************************************'        
                                            '** Real. Timer p� job (Grandtotal / ialt uanset periode) ***'
                                            '************************************************************'

                                            if cint(visPrevSaldo) = 1 then

                                            timerTotSaldo = 0
                                            enhederTot = 0
                                            strSQLprevSaldo =  "SELECT sum(t2.timer) AS sumtimerialt, SUM(t2.timer * a.faktor) AS enhederTot FROM timer AS t2 LEFT JOIN aktiviteter AS a ON (a.id = t2.taktivitetid) WHERE "_
                                            &" "& strSaldoJobAktKri &" AND "& replace(whrClaus, "t.", "") &" "& strSQLgk &" GROUP BY tjobnr"

                                            'Response.write strSQLprevSaldo
                                            'Response.flush 

                                            oRec4.open strSQLprevSaldo, oConn, 3
                                            if not oRec4.EOF then

                                            timerTotSaldo = oRec4("sumtimerialt")
                                            enhederTot = oRec4("enhederTot")

                                            end if
                                            oRec4.close


                                        
                                            if cint(directexp) <> 1 then 
                                            strJobLinie = strJobLinie & "<td class=lille valign=bottom align=right "&tdstyleTimOms20&">"
                                            end if 'if cint(directexp) <> 1 then 

                                             '** Res timer p� job ***'
                                            if cint(vis_restimer) = 1 then
                                            call resTimer(jobmedtimer(x,0), jobmedtimer(x,12), 0, 0, 0, md_split_cspan, 2)


                                            if len(trim(restimerThis)) <> 0 then
                                            restimerTot = restimerThis
                                            restimerThis = restimerThis
                                            else
                                            restimerThis = 0
                                            restimerTot = 0
                                            end if
                                            
                                            restimerThisTxt = restimerThis

                    
                                            

                                            if formatnumber(restimerTot) <> formatnumber(restimerPS/1+restimerPer) then
                                            fntRCol = "#999999"
                                            rsign = "!="
                                            else
                                            fntRCol = "#999999"
                                            rsign = "="
                                            end if

                                            if len(trim(restimerThis)) <> 0 then
                                                if cint(directexp) <> 1 then 
                                                strJobLinie = strJobLinie & "<span style='color:"&fntRCol&";'> "&rsign&" " & restimerThisTxt &"</span><br>"
                                                end if 'if cint(directexp) <> 1 then 
                                            end if
                                            
                                           

                                            restimerTotalGtotalJob = (restimerTotalGtotalJob/1) + (restimerThis/1)
                                            restimerSubGtotalJob = (restimerSubGtotalJob/1) + (restimerThis/1)

                                            end if


                                            if formatnumber(timerTotSaldo) <> formatnumber(timerPrevSaldo/1+jobmedtimer(x,16)) then
                                            fntCol = "darkred"
                                            tSign = "!="
                                            else
                                            fntCol = "#000000"
                                            tSign = "="
                                            end if

                                            if cint(directexp) <> 1 then 
                                            strJobLinie = strJobLinie &"<span style='color:"& fntCol &";'> "&tSign&" "& formatnumber(timerTotSaldo, 2) &"</span>"
                                            end if 'if cint(directexp) <> 1 then                                     

                                            if cint(vis_enheder) = 1 then
                                                
                                            if formatnumber(enhederTot) <> formatnumber(enhederPrevSaldo/1+jobmedtimer(x,26))  then
                                            efntCol = "#3b5998"
                                            eSign = "~"
                                            else
                                            efntCol = "#5c75AA"
                                            eSign = "="
                                            end if

                                            if cint(directexp) <> 1 then 
                                            strJobLinie = strJobLinie &"<br><span style='color:"& efntCol &"; font-size:9px;'> "&eSign&" enh. "& formatnumber(enhederTot, 2) & "</span>"
                                            end if 'if cint(directexp) <> 1 then 
                        

                                            enhederGTot = (enhederGTot/1 + enhederTot/1)
                                            enhederGSub = (enhederGSub/1 + enhederTot/1) 

                                            end if

                                            if cint(directexp) <> 1 then 
                                            strJobLinie = strJobLinie &"</td>"
                                            end if 'if cint(directexp) <> 1 then 

                                            timerTotSaldoSubGtotal = timerTotSaldoSubGtotal + timerTotSaldo '(timerPrevSaldo/1+jobmedtimer(x,16)) 
                                            timerTotSaldoGtotal = timerTotSaldoGtotal + timerTotSaldo '(timerPrevSaldo/1+jobmedtimer(x,16)) 

                                           

                                             expTxt = expTxt & formatnumber(timerTotSaldo,2) &";"
                                            
                                             if cint(vis_restimer) = 1 then
                                             expTxt = expTxt & restimerThis &";"
									         end if

                                             if cint(vis_enheder) = 1 then
                                             expTxt = expTxt & formatnumber(enhederTot, 2) &";"
									         end if
										
                                             

                                             end if 'cint(visPrevSaldo) = 1





											
											
											lastjid = lastjid &",#"&jobaktId&"#"
									
									


									
									end if 'instr lastjob
										
								
								
								
				
						'*************************************************************'
                        '**** MAIN
						'**** Timer p� hver Medarbjeder LOOP ***'
						'*************************************************************'
						
                        'lastWrtMd = 1
						'cnt = 1  
                           
                            'Response.Write "jobmedtimer(x,12): "& jobmedtimer(x,12) & "upSpec: "& upSpec & "jobid: "& jobid & "<br>"
                            if (cdbl(jobmedtimer(x,12)) <> 0 AND cint(upSpec) = 1) OR upSpec = 0 OR jobid = 0 then

							for v = 0 to v - 1 'UBOUND(medarb) 'v - 1
								

                                'if jobmedtimer(x,12) <> 0 then
                                'strJobLinie = strJobLinie & "<td>mv: "& medarb(v) &" v:"& v &" (x,4): "& jobmedtimer(x,4) &" | </td>"
                                'end if

                                if cint(directexp) <> 1 then 

                                select case right(v, 1)
                                case 0,2,4,6,8
                                bgthis = "#FFFFFF"
                                case else
                                bgthis = "#EFf3FF"
                                end select

                                end if 'if cint(directexp) <> 1 then 

                                lloops = 1

                                
                              

								if medarb(v) = jobmedtimer(x,4) then
                                
                                
								'*** timer total ***
                                'lastJMisWrt = 0 

                                if lastWrtM <> jobmedtimer(x,4) then
                                lastWrtMd = -1
                                end if
                                
                                if jobmedtimer(x,38) = 0 then
                                lastWrtJTjk = jobmedtimer(x,0)
                                else
                                lastWrtJTjk = jobmedtimer(x,12)
                                end if

                                if lastWrtJTjk <> lastWrtJ then
                                lastWrtMd = -1
                                end if

                               if isnull(datepart("m", jobmedtimer(x,36), 2,2)) = true then
                               thisWrtMd = 0
                               else
                               thisWrtMd = datepart("m", jobmedtimer(x,36), 2,2)
                               end if

                                'thisWrtM = jobmedtimer(x,4)

                                nextWrtM = jobmedtimer(x+1, 4)
                                


                                if jobmedtimer(x,38) = 0 then
                                nextWrtJ = jobmedtimer(x+1, 0)
                                'thisWrtJ = jobmedtimer(x,0)
                                else
                                nextWrtJ = jobmedtimer(x+1, 12)
                                'thisWrtJ = jobmedtimer(x,12)
                                end if


								
                                'for md = lastWrtMd to tjkMd  'md_split_cspan 
                                loopsFO = -1 'mdStart

                               

                                    for md = 0 To etal '2 'mdStart to mdSlut 'md_split_cspan 

                                
                                        '** Res timer ***'
                                        if md = 0 then
                                        md_year = dateAdd("m", 0, datoStart)
                                        md_md = month(md_year)
                                        md_year = year(md_year)
                                        else
                                        md_year = dateAdd("m", md, datoStart)
                                        md_md = month(md_year)
                                        md_year = year(md_year)
                                        end if



                                    if (datepart("m", jobmedtimer(x,36), 2,2) = md_md) OR cint(md_split_cspan) = 1 then


                                                  

                                                    '*** Tomme felter ved 3 mds opsplitning ***'
                                                    if cint(md_split_cspan) = 3 then '*** 3 md tomme felter

                                                        select case md 'md 'datepart("m", jobmedtimer(x,36), 2,2)
                                                        case 0
                                      
                                                        case 1, 2
                                                             
                                                                   for m = 1 to 3
                                                         
                                                                     'mdThis = dateadd("m", -(12-m), datoSlut)
                                                                     'mdThis = month(mdThis)
                                                                 
                                                                     'Response.Write "cnt: "& cnt &" mdThis: "& mdThis & " m: "& m & " md:" & md & "md_md:" & md_md & "<br>" 

                                                                     if m <= md then
                                                                         if cint(csv_pivot) <> 1 then
                                                                            call tomtdfelt_a
                                                                         end if
                                                                         'cnt = cnt + 1
                                                                     end if

                                                                  next
                                                            
                                    
                                      
                                                
                                                        end select



                                                    end if '3 md tomme felter





                                                    '*** Tomme felter ved 12 mds opsplitning ***'
                                                    if cint(md_split_cspan) = 12 then '*** 3 md tomme felter

                                                        select case md 'md 'datepart("m", jobmedtimer(x,36), 2,2)
                                                        case 0
                                      
                                                        case 1,2,3,4,5,6,7,8,9,10,11

                                                     
                                                                  for m = 1 to 12
                                                         
                                                                     'md_year = dateadd("m", -(12-m), datoSlut)
                                                                     'mdThis = month(mdThis)
                                                                 
                                                                     'Response.Write "cnt: "& cnt &" mdThis: "& mdThis & " m: "& m & " md:" & md & "md_md:" & md_md & "<br>" 

                                                                     if m <= md then

                                                                         if cint(csv_pivot) <> 1 then
                                                                         call tomtdfelt_a
                                                                         'cnt = cnt + 1
                                                                         end if
                                                                     end if

                                                                  next
                                                              
                                                             
                                                                  

                                                            'case else
                                                            'strJobLinie = strJobLinie & "<td class=lille align=right "& tdstyleTimOms3 &" bgcolor='"& bgthis &"'>&nbsp;M:"& m &"</td>"

                                                        end select



                                                    end if ' 12 md tommefelter

                                                
                                               
        
                                           

                                              call timerTdFelt

                                        '*************************************************************************'

                                
                               

                                    end if

                                   'end if '** md = 
								
                                    lloops = lloops + 1
                                
                                    next '**md spilt








                                      if md_split_cspan <> 1 then

                                                        
                                                         


                                                        if (cint(vis_medarbejdertyper) = 0 AND ((nextWrtJ <> lastWrtJ) OR (nextWrtM <> lastWrtM))) OR _
                                                         (cint(vis_medarbejdertyper) = 1 AND antalV = 1 AND ((nextWrtJ <> lastWrtJ) OR (nextWrtM <> lastWrtM)) AND (lastWrtX >= x)) OR _
                                                         (cint(vis_medarbejdertyper) = 1 AND antalV <> 1 AND ((nextWrtJ <> lastWrtJ) OR (nextWrtM <> lastWrtM)) AND (jobmedtimer(x,4) <> nextWrtM) ) then 'AND (jobmedtimer(x,4) <> nextWrtM OR loopsFO > 0)    then 'AND (jobmedtimer(x,4) <> nextWrtM OR nextWrtM = "")   ' AND ((jobmedtimer(x,4) <> nextWrtM) OR (jobmedtimer(x,4) = nextWrtM AND lastWrtM <> 0))) then

                                                       
                                                        loopsFO = loopsFO + 1
                                                        

                                                        

                                                        for f = (loopsFO) to etal 'month(datoStart) + 2

                                
                                                        if cint(directexp) <> 1 then 
                                                        strJobLinie = strJobLinie & "<td class=lille align=right "& tdstyleTimOms3 &" bgcolor='"& bgthis &"'>&nbsp;"
                                                        end if 'if cint(directexp) <> 1 then 
                        
                                                        'M: "& jobmedtimer(x,4) &" lastM: "& lastWrtM &" - nextM: "& nextWrtM & " - x: "& x & " nxtJ: "& nextWrtJ &" <> lstJ: "& lastWrtJ & " loopsFO:" &loopsFO & " antalV: "& antalV & " lloops:"& lloops & " lastWrtX: "& lastWrtX 'vsiMtyp: "& cint(vis_medarbejdertyper)
                                                        'strJobLinie = strJobLinie & "<br> medarb(v): " &  medarb(v) & " Medid: <b>"& jobmedtimer(x,4) & "</b><br>"
                                                        '&" loopsFO = "& loopsFO &" lastWrtMd: "& lastWrtMd &" nextWrtJ: "& nextWrtJ &" lastWrtJ: "& lastWrtJ &" nextWrtM: "& nextWrtM &" lastWrtM: "& lastWrtM &""_
                                                        '&"<br>lastJM:"& lastJM &""
                                                        'Response.Write v &  " medarb(v): " &  medarb(v) & " jobmedtimer(x,4): "& jobmedtimer(x,4) & "<br>"

                                                        expTxt = expTxt &";"

                                                        '"& md &" - "& lloops &" "_
                                                        '&" loopsFO = "& loopsFO &" lastWrtMd: "& lastWrtMd &" nextWrtJ: "& nextWrtJ &" lastWrtJ: "& lastWrtJ &" nextWrtM: "& nextWrtM &" lastWrtM: "& lastWrtM &""_
                                                        '&"<br>lastJM:"& lastJM &""



                                                                     '** Res timer ***'
                                                                     if cint(vis_restimer) = 1 then
                                    
                                                                        li = "e"
                                    
                                                                         md_year = dateAdd("m", f, datoStart)
                                                                         md_md = month(md_year)
                                                                         md_year = year(md_year)
                                  

                                                                        call resTimer(jobmedtimer(x,0), jobmedtimer(x,12), jobmedtimer(x,4), md_md, md_year, md_split_cspa, 0)
                                 
                                                                        if cint(directexp) <> 1 then 
                                                                        strJobLinie = strJobLinie & "<span style='color:#999999;'>"& restimerThis &"</span><br>&nbsp;"
                                                                        end if

                                                                                if len(trim(restimerThis)) <> 0 then
                                                                                restimerThis = restimerThis
                                                                                restimerThisExp = restimerThis
                                                                                else
                                                                                restimerThis = 0
                                                                                restimerThisExp = ""
                                                                                end if

                                                                                medabRestimer(v) = medabRestimer(v) + restimerThis
                                                                                subMedabRestimer(v) = subMedabRestimer(v) + restimerThis

                                                                                if firtsEfelt = 0 then
                                                                                firtsEfelt = 1
                                                                                lastWrtMdRStimerCountM = lastWrtMd + 1
                                                                                else
                                                                                lastWrtMdRStimerCountM = lastWrtMdRStimerCountM + 1
                                                                                end if
                                                                    
                                                                                medabTotRestimerprMd(v, lastWrtMdRStimerCountM) = (medabTotRestimerprMd(v, lastWrtMdRStimerCountM)/1 + restimerThis/1)

                                                                                'strJobLinie = strJobLinie & " - "& lastWrtMdRStimerCountM


                                                                                expTxt = expTxt & restimerThisExp &";"


                               
                                                                     end if
                                                                     '*****

                                 
                                                                        if cint(vis_enheder) = 1 then
							                                            expTxt = expTxt &";" 
							                                            end if
							
							                                            if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
							                                            expTxt = expTxt &";;"
							                                            end if    
                                                                            
                                                                         'if cint(vis_normtimer) = 1 then ALDRING MED P� eksport n�r der opldeles p� 3 og 12 MD
                                                                         'expTxt = expTxt &";" 
                                                                         'end if


                                                                        if vis_redaktor = 1 AND f = etal then
                                                                            
                                                                            if jobmedtimer(x,12) <> 0 then
                                                                            thisTp = "n"
                                                                            else
                                                                            thisTp = ""
                                                                            end if

                                                                            '*** Finder tp p� job / akt ****
                                                                            strSQLtp = "SELECT id, 6timepris FROM timepriser WHERE jobid = "& jobmedtimer(x,0) &" AND aktid = "& jobmedtimer(x,12) &" AND medarbid = "& jobmedtimer(x,4)
                                                                            'Response.write strSQLtp & "<br><br>"
                                                                            oRec.open strSQLtp, oConn, 3
                                                                            if not oRec.EOF then 
            
                                                                                thisTp = formatnumber(oRec("6timepris"))

                                                                            else
                                                                                '** Find prisen ved hj�lp af medarb. typen ***'
                                                                                '** Kun p� job, s� man kan se om akt nedarver **'

                                                                                if jobmedtimer(x,12) = 0 then
                                                                                strSQLtpmt = "SELECT m.medarbejdertype, mt.timepris FROM medarbejdere AS m "_
                                                                                &" LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) WHERE m.mid = "& jobmedtimer(x,4) 


                                                                                'Response.write strSQLtpmt & "<br>"
                                                                                'Response.flush

                                                                                 oRec3.open strSQLtpmt, oConn, 3
                                                                                 if not oRec3.EOF then
                                                                                 
                                                                                 thisTp = formatnumber(oRec3("timepris")) 

                                                                                 end if
                                                                                 oRec3.close

                                                                                
                                                                                else 
                                                                                'Aktiviteter
                                                                                'Viser tp fra job hvis akt st�r til at nedarve


                                                                                 
                                                                                strSQLtpnedarv = "SELECT id, 6timepris FROM timepriser WHERE jobid = "& jobmedtimer(x,0) &" AND aktid = 0 AND medarbid = "& jobmedtimer(x,4)
                                                                                'Response.write strSQLtpmt & "<br>"
                                                                                'Response.flush

                                                                                 oRec3.open strSQLtpnedarv, oConn, 3
                                                                                 if not oRec3.EOF then
                                                                                 
                                                                                 thisTp = formatnumber(oRec3("6timepris")) 

                                                                                 end if
                                                                                 oRec3.close


                                                                                 end if

                                                                                
            
                                                                            end if
                                                                            oRec.close
                            

                                                                        if cint(directexp) <> 1 then 
                                                                        strJobLinie = strJobLinie &"<br><input type='text' class='f_tp_"&jobmedtimer(x,4)&"_"&jobmedtimer(x,0)&"_"&jobmedtimer(x,12)&"' value='"& thisTp & "' name='FM_tp_j' style='color=#999999; padding:0px; width:40px; height:14px; line-height:9px; font-size:9px;'> /t."
                                                                        strJobLinie = strJobLinie &"<input type='hidden' value='#' name='FM_tp_j'>"
                                                                        strJobLinie = strJobLinie &"<input type='hidden' value='"& jobmedtimer(x,6) &"' name='FM_jobnr_j'>"
                                                                        strJobLinie = strJobLinie &"<input type='hidden' value='"& jobmedtimer(x,0) &"' name='FM_jobid_j'>"
                                                                        strJobLinie = strJobLinie &"<input type='hidden' value='"& jobmedtimer(x,12) &"' name='FM_aktid_j'>"
                                                                        strJobLinie = strJobLinie &"<input type='hidden' value='"& jobmedtimer(x,4) &"' name='FM_medid_j'>"
                                                                        end if 'if cint(directexp) <> 1 then                                                                  


                                                                        end if
                                                                        
                                                        if cint(directexp) <> 1 then 
                                                        strJobLinie = strJobLinie & "</td>"
                                                        end if' if cint(directexp) <> 1 then 
        
                                                        next

                                                       'lastJMisWrt = 1'lastWrtJ &"_"&lastWrtM

                                                       Response.flush

                                                       end if  '** nextWrtJ <> lastWrtJ) OR (nextWrtM <> lastWrtM

                                           end if '** md spilt


                               'else
                               'Response.Write v &"<br>"

								'*** For at krydstjekke bliver restimer og balance summeret som totaler p� medarb **'
								'*** Mens Timer ialt og fakbare timer og Oms. bliver summeret p� jobtotaler **'
								'restimerTotTot = restimerTotTot + restimerThis
								
                                
                                end if '*** medarb(v) = 
								

                               
                                
							firtsEfelt = 0
							next 'medarb
                            
						    

                            end if 'udSpec
						
						
						
					
					response.flush
						
				'end if 'jobmedtimer(x,16) og Nulfilter
				

                 firtsEfelt = 0
				'*** Tilprint ***
				next 'job
				
			if antaljob <> 0 then	
			antaljob = antaljob	
            else
            antaljob = 1
            end if
			
            if cint(directexp) <> 1 then 
			strJobLinie = strJobLinie & "</tr>"

            '*** Subtotal i bunden ***'
             if cint(upSpec) = 1 then
              call subTotaler_gt
            end if
			
            end if 'if cint(directexp) <> 1 then 

			
			'***************************************
			'**** Udskriver Tables *****************
			'***************************************
            if cint(directexp) <> 1 then 
			Response.write strJobLinie_top
			

			'*** Skal job vises eller vises kun totaler?
			Response.write strJobLinie
			end if 'if cint(directexp) <> 1 then 
			
			
			
			'***************************'
			'*** totaler i bunden ******'
			'***************************'
			
            if cint(csv_pivot) <> 1  then

				
				expTxt = expTxt &"xx99123sy#z"
				
                select case lto
                case "cisu"
				expTxt = expTxt &"Grandtotal;;;;;"
                case else
                        select case lto
                        case "mmmi", "intranet - local"
                        expTxt = expTxt &"Grandtotal;;;;;;;;;;"
                        case else
                        expTxt = expTxt &"Grandtotal;;;;;;;;;;;"
                        end select
                end select
				
				

                 'if cint(vis_restimer) = 1 then
                 'expTxt = expTxt &";"
                 'end if

				if cint(directexp) <> 1 then 
			    strJobLinie_total = "<tr bgcolor=lightpink>"
				strJobLinie_total = strJobLinie_total & "<td style='padding:4px; border-top:1px #CCCCCC solid;' valign=Top><b>Grandtotal:</b></td>"
						
						strJobLinie_total = strJobLinie_total & "<td class=lille valign=bottom align=right "&tdstyleTimOms2&" >" 
						strJobLinie_total = strJobLinie_total & formatnumber(totbudgettimer, 2) 

                        if cint(vis_enheder) = 1 then
					    strJobLinie_total = strJobLinie_total & "<br>"
						end if

                end if 'if cint(directexp) <> 1 then 
						
					    expTxt = expTxt & formatnumber(totbudgettimer, 2) &";"
								
						if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then

                        if cint(directexp) <> 1 then
						strJobLinie_total = strJobLinie_total & "<br><span style='color=#000000; font-size:8px;'>"& basisValISO &" "& formatnumber(totbudget, 2)&"</span></td>"
                        end if 'if cint(directexp) <> 1 then                

                        expTxt = expTxt & formatnumber(totbudget, 0) &";"
                            
                           
                            
                                  
					
						end if
						
                        '***************'
                        '**** Saldo ***'
                        '***************'
                        if cint(visPrevSaldo) = 1 then

                        if cint(directexp) <> 1 then
                        strJobLinie_total = strJobLinie_total & "<td class=lille valign=bottom align=right "&tdstyleTimOms2&">"
                        end if 'if cint(directexp) <> 1 then                

                        expTxt = expTxt & formatnumber(saldoJobTotal,2) &";"
                        'expTxt = expTxt & formatnumber(timerTotSaldoGtotal,2) &";"
                            
                            if cint(vis_restimer) = 1 then
                                if cint(directexp) <> 1 then
                                strJobLinie_total = strJobLinie_total & "<span style='color:#999999;'>"& formatnumber(saldoRestimerTotal,0) &"</span><br>"
                                end if 'if cint(directexp) <> 1 then
                            expTxt = expTxt & formatnumber(saldoRestimerTotal,0) &";"
                            end if

                         if cint(directexp) <> 1 then
                         strJobLinie_total = strJobLinie_total & formatnumber(saldoJobTotal, 2) 
                         end if 'if cint(directexp) <> 1 then


                         if cint(vis_enheder) = 1 then
                             if cint(directexp) <> 1 then
                             strJobLinie_total = strJobLinie_total &"<br><span style='color:#5c75AA; font-size:9px;'>enh. "& formatnumber(enhederPrevSaldoTot, 2) & "</span>" 
                             end if 'if cint(directexp) <> 1 then 
        
                         expTxt = expTxt & formatnumber(enhederPrevSaldoTot,2) &";"
                         end if
                        
                        if cint(directexp) <> 1 then
                        strJobLinie_total = strJobLinie_total &"</td>"
                        end if 'if cint(directexp) <> 1 then

                        end if 

						
                        '***********************'
						'*** Jobtotaler i periode ***'
						'***********************'
                        if cint(directexp) <> 1 then
                        strJobLinie_total = strJobLinie_total & "<td class=lille valign=bottom align=right "&tdstyleTimOms10&">" 
						

                        '**** Res timer ***'
                        if cint(vis_restimer) = 1 then
                        strJobLinie_total = strJobLinie_total & "<span style='color:#999999;'>fc:"& formatnumber(restimerTotalJob,0) &"</span><br>"
                        end if

                         strJobLinie_total = strJobLinie_total & formatnumber(totaltotaljboTimerIalt,2)

						'*** Enheder ***'
						if cint(vis_enheder) = 1 then
					    strJobLinie_total = strJobLinie_total & "<br><span style='color:#5c75AA; font-size:9px;'>enh. " & formatnumber(totalJobEnh,2) & "</span>" 
					    end if

                             

                        end if 'if cint(directexp) <> 1 then      
								
								expTxt = expTxt & formatnumber(totaltotaljboTimerIalt,2) &";"


                                select case lto
                                case "mmmi", "intranet - local"
                                
                                case else
                                '** Real. timer % GT
                                expTxt = expTxt &";"
                                end select

                                if cint(vis_restimer) = 1 then
                                expTxt = expTxt & formatnumber(restimerTotalJob,0) &";"
                                end if
								
								if cint(vis_enheder) = 1 then
								expTxt = expTxt & formatnumber(totalJobEnh, 2)&";"
								end if
								
						
						if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
                    
                        if cint(directexp) <> 1 then
						strJobLinie_total = strJobLinie_total & "<br><span style='color=#000000; font-size:8px;'>"& basisValISO &" "&formatnumber(totaltotaljobOmsIalt, 2)& "</span>" 
						strJobLinie_total = strJobLinie_total & "<br><span style='color=#000000; font-size:8px;'>bal.: "&formatnumber(dbialt, 2)&"</span></td>"
		                end if 'if cint(directexp) <> 1 then				

								expTxt = expTxt & formatnumber(totaltotaljobOmsIalt, 2) &";"
								expTxt = expTxt & formatnumber(dbialt, 2) &";"
								
						else
                            if cint(directexp) <> 1 then
		    				strJobLinie_total = strJobLinie_total & "</td>"
                            end if
						end if
						
            
                        '** Real. % **'
                        select case lto
                        case "mmmi", "intranet - local"

                        case else
        
                        if cint(directexp) <> 1 then
                        strJobLinie_total = strJobLinie_total & "<td class=lille valign=bottom align=right "&tdstyleTimOms10&">&nbsp;</td>"
                        end if 
				
                        end select



                '***********************************'		
				'**** Grandtotal uanset periode ****'
                '***********************************'		
                if cint(visPrevSaldo) = 1 then
                            
                            if cint(directexp) <> 1 then
                            strJobLinie_total = strJobLinie_total & "<td class=lille valign=bottom align=right "&tdstyleTimOms20&">"

                                if cint(vis_restimer) = 1 then

                                if restimerTotalGtotalJob <> (restimerTotalJob+saldoRestimerTotal) then
                                rsSign = "!="
                                rsFontC = "#999999"
                                else
                                rsSign = "="
                                rsFontC = "#999999"
                                end if

                                strJobLinie_total = strJobLinie_total & "<span style='color:"& rsFontC &";'>fc: "&rsSign&" "& formatnumber(restimerTotalGtotalJob,0) &"</span><br>"
                            
                           
                                end if
                            

                                if timerTotSaldoGtotal <> (totaltotaljboTimerIalt+saldoJobTotal) then
                                tsSign = "!="
                                tsFontC = "darkred"
                                else
                                tsSign = "="
                                tsFontC = ""
                                end if
                            
                            strJobLinie_total = strJobLinie_total & "<span style='color:"&tsFontC&";'>"& tsSign &" "& formatnumber(timerTotSaldoGtotal, 2) &"</span>"
                            end if 'if cint(directexp) <> 1 then                    


                            if cint(vis_enheder) = 1 then

                                if formatnumber(enhederGTot) <> formatnumber(totalJobEnh/1+enhederPrevSaldoTot)  then
                                efntCol = "#3b5998"
                                eSign = "!="
                                else
                                efntCol = "#5c75AA"
                                eSign = "="
                                end if

                                if cint(directexp) <> 1 then
                                strJobLinie_total = strJobLinie_total &"<br><span style='color:"& efntCol &"; font-size:9px;'> "&eSign&" enh. "& formatnumber(enhederGTot, 2) & "</span>"
                                end if 'if cint(directexp) <> 1 then
                            
                            end if

                            if cint(directexp) <> 1 then
                            strJobLinie_total = strJobLinie_total &"<br><br>&nbsp;</td>"
                            end if 'if cint(directexp) <> 1 then 

                             expTxt = expTxt & formatnumber(timerTotSaldoGtotal,2) &";"
                             
                             if cint(vis_restimer) = 1 then
                              expTxt = expTxt & formatnumber(restimerTotalGtotalJob,0) &";"
                              end if

                              if cint(vis_enheder) = 1 then
                              expTxt = expTxt & formatnumber(enhederGTot,2) &";"
                              end if
				

                end if
				
				

                '***************************************************'
                '******** LOOP medarbejdere Grandtotal i bunden ****'
                '***************************************************'

				'if cint(jobid) = 0 then
				
						for v = 0 to v - 1

                            
                            
                                            if cint(vis_normtimer) = 1 then

                                                call normtimerPer(medarb(v), datoStart, perinterval, io)
                                                medabNormtimer(v) = ntimPer

                                            end if
                
                            
                                    if md_split_cspan <> 1 then

                                                for mth = 0 to md_split_cspan - 1
                                        
                                                if cint(directexp) <> 1 then
						                        strJobLinie_total = strJobLinie_total & "<td class=lille valign=bottom align=right "&tdstyleTimOms3&">"
                            
                                                    if cint(vis_restimer) = 1 then

                                                    if len(trim(medabTotRestimerprMd(v, mth))) <> 0 then
                                                    strJobLinie_total = strJobLinie_total & "<span style='color=#999999'>"& formatnumber(medabTotRestimerprMd(v, mth), 0) & "</span><br>"
                                                    else
                                                    strJobLinie_total = strJobLinie_total &"&nbsp;"
                                                    end if

                                                    end if


                                                end if 'if cint(directexp) <> 1 then

                                                if len(trim(medabTottimerprMd(v, mth))) <> 0 then
                                                    if cint(directexp) <> 1 then
                                                    strJobLinie_total = strJobLinie_total & formatnumber(medabTottimerprMd(v, mth), 2)
                                                    end if 'if cint(directexp) <> 1 then
                                                expTxt = expTxt & formatnumber(medabTottimerprMd(v, mth), 2) &";"
                                                else
                                                    if cint(directexp) <> 1 then
                                                    strJobLinie_total = strJobLinie_total &"&nbsp;"
                                                    end if' if cint(directexp) <> 1 then
                                                expTxt = expTxt &";"
                                                end if

                                                    if cint(vis_restimer) = 1 then

                                                    if len(trim(medabTotRestimerprMd(v, mth))) <> 0 then
                                                    expTxt = expTxt & formatnumber(medabTotRestimerprMd(v, mth), 2) &";"
                                                    else
                                                    expTxt = expTxt &";"
                                                    end if

                                                    end if


                                                    if cint(vis_enheder) = 1 then

                                                        if len(trim(medabTotEnhprMd(v, mth))) <> 0 then
                                                            if cint(directexp) <> 1 then
                                                            strJobLinie_total = strJobLinie_total & "<br><span style='color=#5c75AA; font-size:9px;'>enh. "& formatnumber(medabTotEnhprMd(v, mth), 2) & "</span>"
                                                            end if 'if cint(directexp) <> 1 then
                                                        expTxt = expTxt & formatnumber(medabTotEnhprMd(v, mth), 2) &";"
                                                        else
                                                            if cint(directexp) <> 1 then
                                                            strJobLinie_total = strJobLinie_total &"&nbsp;"
                                                            end if
                                                        expTxt = expTxt &";"
                                                        end if

                                                    end if


                                                    if cint(vis_normtimer) = 1 then

            
                                                        '** NORM ALDRIG MED P� VISNING OPDELT P� MD                                  
                                                        'if len(trim(medabNormtimer(v))) <> 0 then
                                                        'expTxt = expTxt & formatnumber(medabNormtimer(v), 2) &";"
                                                        'else
                                                        'expTxt = expTxt &";"
                                                        'end if

                                                

                                                    end if

                                                if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then

                                                    if len(trim(medabTotOmsprMd(v, mth))) <> 0 then
                                                    if cint(directexp) <> 1 then
                                                    strJobLinie_total = strJobLinie_total & "<br><span style='color=#000000; font-size:8px;'>"& basisValISO &" "& formatnumber(medabTotOmsprMd(v, mth), 2) & "</span>"
                                                    end if 'if cint(directexp) <> 1 then
                                                    expTxt = expTxt & formatnumber(medabTotOmsprMd(v, mth), 2) &";;"
                                                    else
                                                    if cint(directexp) <> 1 then
                                                    strJobLinie_total = strJobLinie_total &"&nbsp;"
                                                    end if 'if cint(directexp) <> 1 then
                                                    expTxt = expTxt &";;"
                                                    end if

                                                end if
                            


                                                if cint(directexp) <> 1 then

                                                if mth <> md_split_cspan - 1 then

                                            
                                                        if cint(vis_restimer) = 1 then
                                                             'strJobLinie_total = strJobLinie_total & "<br>"
                                                        end if


                                                        if cint(vis_enheder) = 1 then
                                                            'strJobLinie_total = strJobLinie_total & "<br>"
                                                        end if

                                                        if cint(vis_normtimer) = 1 then
                                                            'strJobLinie_total = strJobLinie_total & "<br>"
                                                        end if

                                                strJobLinie_total = strJobLinie_total & "</td>"
                                                else
                                                strJobLinie_total = strJobLinie_total & "<br>Ialt i periode:<br> "
                                                end if
							  
                                                end if 'if cint(directexp) <> 1 then
						    


						                        next


                                    else
                                        if cint(directexp) <> 1 then
                                        strJobLinie_total = strJobLinie_total & "<td class=lille align=right valign=bottom "&tdstyleTimOms3&">"
                                        end if 'if cint(directexp) <> 1 then
                                    end if


                                                        'strJobLinie_total = strJobLinie_total & "<td class=lille valign=top align=right "&tdstyleTimOms3&" bgcolor=#eff3ff>"
                                                         if cint(directexp) <> 1 then                


                                                         if cint(vis_restimer) = 1 then
                                                            if medabRestimer(v) <> 0 then 
                                                            strJobLinie_total = strJobLinie_total & "<span style='color:#999999;'>fc: "& formatnumber(medabRestimer(v),0) &"</span><br>"
                                                            else
                                                            strJobLinie_total = strJobLinie_total & "&nbsp;<br>"
                                                            end if
                                                         end if
                        
                                                        if medabTottimer(v) <> 0 then
                                                        strJobLinie_total = strJobLinie_total & formatnumber(medabTottimer(v), 2)  
                                                        else
                                                        strJobLinie_total = strJobLinie_total & "&nbsp;"
                                                        end if
						
						                                if cint(vis_enheder) = 1 then
                                                            if medabTotenh(v) <> 0 then
						                                    strJobLinie_total = strJobLinie_total & "<br><span style='color:#5c75AA; font-size:9px;'>enh. " & formatnumber(medabTotenh(v), 2) & "</span>"
                                                            else
                                                            strJobLinie_total = strJobLinie_total & "&nbsp;"
                                                            end if
						                                end if

                        
                            
                                                         if cint(vis_normtimer) = 1 then

                                                            if medabNormtimer(v) <> 0 then 
                                                            strJobLinie_total = strJobLinie_total & "<br><span style='color:#999999;'>n: "& formatnumber(medabNormtimer(v),2) &"</span>"
                                                            else
                                                            strJobLinie_total = strJobLinie_total & "&nbsp;<br>"
                                                            end if
                                                         end if


                                                        end if 'if cint(directexp) <> 1 then



						        
                                  if md_split_cspan = 1 then 'Ingen opdeling p� md
                                    
                                 
                               

								    expTxt = expTxt & formatnumber(medabTottimer(v), 2) &";"


                                     if cint(vis_restimer) = 1 then
                                    expTxt = expTxt & formatnumber(medabRestimer(v),2) &";"
                                    end if
								
								    if cint(vis_enheder) = 1 then
								    expTxt = expTxt & formatnumber(medabTotenh(v), 2)&";"
								    end if


                                           if cint(vis_normtimer) = 1 then

                                                
                                                    if len(trim(medabNormtimer(v))) <> 0 then
                                                    expTxt = expTxt & formatnumber(medabNormtimer(v), 2) &";"
                                                    else
                                                    expTxt = expTxt &";"
                                                    end if

                                              

                                            end if
							
							    if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
                                        
                                        if cint(directexp) <> 1 then
                                            if omsTot(v) <> 0 then
							                strJobLinie_total = strJobLinie_total & "<br><span style='color=#000000; font-size:8px;'>"& basisValISO &" "&formatnumber(omsTot(v), 2)&"</span></td>"
                                            else
                                            strJobLinie_total = strJobLinie_total & "<br>&nbsp;</td>"
                                            end if
                                        end if 'if cint(directexp) <> 1 then
								    expTxt = expTxt & formatnumber(omsTot(v), 2) &";;"
							    else
                                    if cint(directexp) <> 1 then
							        strJobLinie_total = strJobLinie_total & "<br>&nbsp;</td>"
                                    end if 'if cint(directexp) <> 1 then
							    end if


                                end if
                        	
						next
				 ' end if 'if visMedarb = 1 then

            
        
            if cint(directexp) <> 1 then

			strJobLinie_total = strJobLinie_total & "</tr></table></td></tr></table>"
		    strJobLinie_total = strJobLinie_total & "<br><table cellpadding=2 cellspacing=1 width=450 border=0><tr><td>"
            
            if vis_redaktor = 1 then
            
            updTpDatoM = month(datoStart)
            updTpDatoY = year(datoStart)

           
            strJobLinie_total = strJobLinie_total & "<b>Opdater timepriser</b><br>Opdaterer timeregistreringer i valgte m�ned (gr�nne).<br><br>"_
            &"<input type='Checkbox' value='1' name='FM_opdater_tp_job'> Opdater samtidig de timepriser der er indtastet p� job / aktivitet, s� fremtidige timeregistreringer f�lger denne pris (g�lder ikke de aktiviteter der er sat til ens pris for alle).<br><br>n: nedarver fra job<br><br>Alle priser er "& basisValISO &". Valuta �ndres p� jobbet."_
            &"</td><tr><td align=right><input type='submit' value='Opdater >>'></td></tr></table></form>"
            else
            strJobLinie_total = strJobLinie_total & no_redaktorTxt &"</td></tr></table>"
            end if

			
			Response.write strJobLinie_total
			
		    end if 'if cint(directexp) <> 1 then    

            end if 'if cint(csv_pivot) <> 1 then   

            'end if ' max size antal job/periode	
			
            




               

            '******************* Eksport **************************' 
            if cint(directexp) = 1 then  
            'if media = "export" then



                  
                                    
                call TimeOutVersion()
    


            
    

	            ekspTxt = replace(expTxt, "xx99123sy#z", vbcrlf)
	
	            datointerval = request("datointerval")
	
	
	            filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	            filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				            Set objFSO = server.createobject("Scripting.FileSystemObject")
				
				            if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\timereg\joblog_timetotaler.asp" then
					            Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\gtexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					            Set objNewFile = nothing
					            Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\gtexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				            else
					            Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\gtexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					            Set objNewFile = nothing
					            Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\gtexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				            end if
				
				
				
				            file = "gtexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
				
				            '**** Eksport fil, kolonne overskrifter ***
                            'strOskrifter = "Medarbejder; Init; Kunde; Kundenr; Job; Jobnr; Aktivitet; Kommentar;"

                            
				
				            objF.writeLine("Periode afgr�nsning: "& datoStart &" - "& datoSlut & vbcrlf)
				            'objF.WriteLine(strOskrifter & chr(013))
				            objF.WriteLine(ekspTxt)
				            objF.close
				
				            %>
				
	                        <table border=0 cellspacing=1 cellpadding=0 width="400" id="tcsv">
	                        <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	                        <img src="../ill/outzource_logo_200.gif" />
	                        </td>
	                        </tr>
	                        <tr>
	                        <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
                            Du har valgt at udskrive direkte til .cvs fil.<br /> Du �bner den ved at klikke p� linket herunder:<br />
	                        <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank">Din CSV. fil er klar >></a>
	                        </td></tr>
	                        </table>
	                        
                            </div><!-- oCode (n�r der kommer response.end herunder) --> 
	          
	            
	                        <%
                
                
                            Response.end
	                        'Response.redirect "../inc/log/data/"& file &""	
				



            end if

            '***************************************** END export *****************************************
            %>















			<br><br>&nbsp;
			</div><!-- sideindhold -->
			
			<%


              if len(trim(md_splitVal)) <> 0 AND isNull(md_splitVal) <> true then
                    md_splitVal = md_splitVal
                else
                    md_splitVal = 0
                end if

                
                    
                prntLnk = "nomenu="& nomenu &""_
                &"&FM_job="& jobid &""_
                &"&FM_jobsog="& jobSogValPrint &""_
                &"&FM_start_dag="& strDag &""_
	            &"&FM_start_mrd="&strMrd&"&FM_start_aar="&strAar&"&FM_slut_dag="&strDag_slut&""_
	            &"&FM_slut_mrd="&strMrd_slut&"&FM_slut_aar="&strAar_slut&""_
                &"&FM_progrp="& progrp &"&vis_fakbare_res="& visfakbare_res &""_
                &"&vis_aktnavn="& vis_aktnavn &"&csv_pivot="&csv_pivot&"&FM_kundejobans_ell_alle="& visKundejobans &"&FM_kundeans="&kundeans&"&FM_jobans="& jobans1Val & ""_   
                &"&FM_jobans2=" & jobans2 &"&FM_jobans3="& jobans3 &""_
                &"&FM_medarb="& thisMiduse &""_
                &"&FM_fomr="& fomr &"&FM_kunde="& kundeid &"&FM_visMedarbNullinier="& visMedarbNullinier &"&FM_aftaler="& aftaleid & ""_
                &"&FM_md_split="& md_splitVal &""_
                &"&FM_visnuljob="& visnuljob &"&FM_visPrevSaldo="& visPrevSaldo &""_
                &"&FM_udspec="& upSpec &"&FM_directexp="& directexp &"&FM_udspec_all="& upSpec_all & ""_
                &"&vis_fakturerbare="& vis_fakturerbare & "&vis_kpers="& vis_kpers &"&vis_jobbesk="& vis_jobbesk &"&vis_enheder="& vis_enheder & "&vis_restimer=" & vis_restimer & "&vis_normtimer="& vis_normtimer & "&FM_vis_medarbejdertyper="& vis_medarbejdertyper & "&FM_vis_medarbejdertyper_grp="& vis_medarbejdertyper_grp & ""_
                &"&vis_redaktor=" & vis_redaktor & "&FM_segment="& segmentLnk   


			if print <> "j" AND x >= 1 AND media <> "print" then
            
            ptop = pr_Top
            pleft = 830
            pwdt = 165
            
            call eksportogprint(ptop, pleft, pwdt)
            %>

            <form action="joblog_timetotaler_eksport.asp" target="_blank" method="post" name=theForm2 onsubmit="BreakItUp2()"> <!--  -->
			<input type="hidden" name="datointerval" id="datointerval" value="<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>">
			<input type="hidden" name="txt1" id="txt1" value="">
			<input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=expTxt%>">
			<input type="hidden" name="txt20" id="txt20" value="">
			
            <tr>
            
                <%if len(strJobLinie) > 30000000 then %>
        <td class=lille>
        M�ngden af data er for stor til eksport. V�lg et mindre interval ell. f�rre 
        medarbejdere. St�rrelsen p� data er: <%=len(strJobLinie) %> og den m� ikke overstige 30000000 bytes.
        <%else %>
    
        <td><input type="submit" id="sbm_csv" value="CSV. fil eksport >>" style="font-size:9px;" /></td>
   
        <%end if %>
            
              
                </tr>
                </form>
                

                <!--

                <form action="printversion.asp?media=print" method="post" target="_blank" name="theForm" onsubmit="BreakItUp()"> 

			    <input type="hidden" name="datointerval" id="datointerval" value="<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>">
			    <input type="hidden" name="txt1" id="txt1" value="<%=strJobLinie_top%>">
			    <input type="hidden" name="BigTextArea" id="BigTextArea" value="<%=strJobLinie%>">
			    <input type="hidden" name="txt20" id="txt20" value="<%=strJobLinie_total%>">
                <tr>
               
                <td><br /><input type="submit" value="Print version >>" style="font-size:9px;" /> </td>
               </tr>
               </form>

                -->


                <%

              
               

                
                   

                %>

               


                 <form action="joblog_timetotaler.asp?media=print&FM_usedatointerval=1&<%=prntLnk%>" method="post" target="_blank"> 
			    <input type="hidden" name="datointerval" id="datointerval" value="<%=formatdatetime(strDag&"/"&strMrd&"/"&strAar, 1) & " - " & formatdatetime(strDag_slut&"/"&strMrd_slut&"/"&strAar_slut, 1)%>">
			  
                <tr>
               
                <td><br /><input type="submit" value="Print >>" style="font-size:9px;" /> </td>
               </tr>
               </form>
	            
               </table>
            </div>
            
            
            
            
         
            
            
            
            <%end if
                
                'if session("mid") = 1 then
                'response.write "prntLnk: <br>" & prntLnk
                'end if
                
                %>
		
		    <% 'if media = "print" then
                'Response.Write("<script language=""JavaScript"">window.print();</script>")
              'end if  %>
		    
		  
		
<%end if
    
    
    
            'timeB = now
            'loadtime = datediff("s",timeA, timeB)
            'Response.write "<br><br><br><br><br><br><br><br><br><br><br><br><br>loadtid: " & loadtime
    %>



  </div>
<!--#include file="../inc/regular/footer_inc.asp"-->