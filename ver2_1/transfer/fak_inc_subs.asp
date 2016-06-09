<% 

public timeforbrug 
public lastFakdato
public strfaktidspkt
function faktureredetimerogbelob()
        
        if jobid <> 0 then
        jobaftKri = " jobid = " & jobid 
        else
        jobaftKri = " aftaleid = "& aftid 
        end if
        
        
        
		strSQL = "SELECT fid, beloeb, timer, fakdato, tidspunkt, faktype FROM fakturaer WHERE "& jobaftKri &" AND fid <> "& id &" ORDER BY fakdato"
		
		'Response.Write strSQL
		'Response.Flush
		
	    oRec.open strSQL, oConn, 3
		while not oRec.EOF 
				
				strSQL2 = "SELECT Fid, parentfak FROM fakturaer WHERE parentfak = " &oRec("Fid")&" AND faktype = 2"
				oRec2.open strSQL2, oConn, 3 
				hasfak = 0
				if not oRec2.EOF then
				hasfak = 1
				end if
				oRec2.close
				
				
				Select case oRec("faktype") 
				case 0
					if hasfak = 0 then
					faktureretBeloeb = faktureretBeloeb + oRec("beloeb")
					timeforbrug = timeforbrug + oRec("timer")
					else
					faktureretBeloeb = faktureretBeloeb
					timeforbrug = timeforbrug
					end if
				case 1
				faktureretBeloeb = faktureretBeloeb - oRec("beloeb")
				timeforbrug = timeforbrug - oRec("timer")
				case 2
				faktureretBeloeb = faktureretBeloeb + oRec("beloeb")
				timeforbrug = timeforbrug + oRec("timer")
				end select
			
			
			lastFakdato = convertDateYMD(oRec("fakdato")) 
			strfaktidspkt = oRec("tidspunkt")
		oRec.movenext
		Wend
		oRec.close



end function




'***************************************************************
'Fak, Venter og Beløb
'***************************************************************
public medarbejderBeloeb 
public txt 
public medarbid 
public showonfak
public showventer
public venter
public fak
public venter_brugt
public venter_ultimo





'***********************************************************************************
'**** Sum-Aktiviteter subtotal func **********
'***********************************************************************************
	
	public totalbelob
	public intTimer
	public strAktsubtotal
	public subtotalbelob
	public subtotaltimer
	'public antalenhialt
	
	
	
	
	function subtotalakt(x, a)
	
	
	if a < 0 then
	useMedarbejderTimepris = -1
	thisAkttimer(0) = ""
	thisAktBeloeb(0) = ""
	nul = 1
	bgcol = "#ffffff"
	bgcol2 = "#FFFFFF"
	
	hidden = "hidden"
	display = "none"
	
	if func = "red" then
	akt_showonfak = 0
	end if
	
	'hidden = "visible"
	'display = ""
	
	end if
	
	
	
	    '********** Andre? sum-aktiviteten *****
	    if a = 0 then
	
	
		        if thisaktfunc(x) = "red" then
				'*** Finder timepris på ekstra /andre? sum-aktiviteten **
				strSQL = "SELECT enhedspris FROM faktura_det WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &"" 
				
				oRec.open strSQL, oConn, 3 
				while not oRec.EOF 
					
					if instr(isTPused, "#"&SQLBlessDOT(formatnumber(oRec("enhedspris"), 2))&"#") = 0 then
					useEkstraMedarbejderTimepris = oRec("enhedspris")
					end if
					
				oRec.movenext
				wend
				oRec.close 
				
				
				if len(useEkstraMedarbejderTimepris) <> 0 then
				useMedarbejderTimepris = useEkstraMedarbejderTimepris
				else
				useMedarbejderTimepris = -2
				end if
				
				useEkstraMedarbejderTimepris = ""
				
				
				'** Finder showonfak, samt rabat på 0 aktiviteten **
				strSQL2 = "SELECT showonfak, rabat FROM faktura_det WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &" AND enhedspris = "& SQLBless(useMedarbejderTimepris) 'SQLBless(thisTimePris(x)) 
				oRec2.open strSQL2, oConn, 3
				if not oRec2.EOF then
				akt_showonfak = oRec2("showonfak")
				aktrabat = oRec2("rabat")
                end if
				oRec2.close
				
		else
		useMedarbejderTimepris = -2
		thisAkttimer(x) = 0
		thisAktBeloeb(x) = 0
		'akt_rabat = 0
		end if
	
	nul = 2
	bgcol = "#ffffff"
	bgcol2 = "#FFFFFF"
	hidden = "visible"
	display = ""
	end if
	
	if a > 0 then
	'Response.Write "a" & a & " MedarbejderTimepris: " & MedarbejderTimepris
	useMedarbejderTimepris = MedarbejderTimepris
	nul = 1
	bgcol = "#FFFFFF"
	bgcol2 = "#FFFFFF"
	
				
				if thisaktfunc(x) = "red" then
						
						'** Finder showonfak på aktiviteten **
						strSQL2 = "SELECT showonfak, rabat, enhedsang FROM faktura_det WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &" AND enhedspris = "& SQLBless(useMedarbejderTimepris) 'SQLBless(thisTimePris(x)) 
						'Response.Write strSQL2
						'Response.Flush
						oRec2.open strSQL2, oConn, 3
						if not oRec2.EOF then
							
							akt_showonfak = oRec2("showonfak")
							aktrabat = oRec2("rabat")
							aktenhedsang = oRec2("enhedsang")
							
						end if
						oRec2.close
						
						'*** Skal aktiviteten vises? 
						if akt_showonfak = 1 then
						display = ""
						hidden = "visible"
						else
						display = "none"
						hidden = "hidden"
						end if
					
				else
				display = ""
				hidden = "visible"
				end if
	end if
	
		
		
		
		
		strAktsubtotal = strAktsubtotal & "<tr><td valign=top>"_
		&"<div name='sumaktdiv_"&x&"_"&a&"' id='sumaktdiv_"&x&"_"&a&"' style='position:relative; visibility:"&hidden&"; display:"&display&"; background-color:"&bgcol2&"; border-top:1px lightgrey solid; padding:5px 5px 5px 0px;'>"_
		&"<table cellspacing=0 cellpadding=0 border=0 width=650><tr>"
		
		strAktsubtotal = strAktsubtotal &"<td valign=top width=80>"
		strAktsubtotal = strAktsubtotal &"<input type=hidden name='FM_aktid_"&x&"_"&a&"' value='"&thisaktid(x)&"'>"_
		&"<input type=hidden name='FM_hidden_timepristhis_"&x&"_"&a&"' id='FM_hidden_timepristhis_"&x&"_"&a&"' value='"&useMedarbejderTimepris&"'>"
		
		'*** Beregner timer ***********
		'===
		if thisaktfunc(x) = "red" then
				
				
					strSQL2 = "SELECT sum(antal) AS antaltimer FROM faktura_det WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &" AND enhedspris = "& SQLBless(useMedarbejderTimepris) 'SQLBless(thisTimePris(x)) 
					oRec2.open strSQL2, oConn, 3
					
					if not oRec2.EOF then
						hiddentimer = oRec2("antaltimer")
						valueTimer = oRec2("antaltimer")
					end if
					oRec2.close
					
					if len(hiddentimer) <> 0 then
					hiddentimer = hiddentimer
					valueTimer = valueTimer
					else
					hiddentimer = 0
					valueTimer = 0
					end if
					
		else
			if usefastpris(x) <> 1 then 
					if a > 0 then
						
						
						strSQL2 = "SELECT sum(timer.timer) AS antaltimer FROM timer WHERE taktivitetid =" & thisaktid(x) &""_
						&" AND timepris = "&SQLBless(useMedarbejderTimepris)&" AND "_
						&" (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')" 
						'AND tfaktim = 1
						oRec2.open strSQL2, oConn, 3
						'Response.write strSQL2
						if not oRec2.EOF then
							hiddentimer = oRec2("antaltimer")
							valueTimer = oRec2("antaltimer")
						end if
						oRec2.close
						
						
						if hiddentimer > 0 then
						hiddentimer = hiddentimer
						valueTimer = valueTimer 
						else
						hiddentimer = 0
						valueTimer = 0 
						end if
						
						'* Hvis en medarbejder har skiftet timepris og der derfor ikke kan 
						'* findes timer med den beregnede timepris for medarbejderen angives dette med en alert.
						'if thisAkttimer(x) <> 0 AND hiddentimer = 0 then
						'alerttp = "<font color=red><b>!</b></font>"
						'else
						'alerttp = ""
						'end if
						
					else
						
						hiddentimer = 0
						valueTimer = 0
						
					end if
			else
					
					if len(thisAkttimer(x)) <> 0 then
					hiddentimer = thisAkttimer(x)
					else
					hiddentimer = 0
					end if
					
					if len(thisAkttimer(x)) <> 0 then
					valueTimer = thisAkttimer(x)
					else
					valueTimer = 0
					end if
					
			end if
		end if
		
		
		
	if thisaktfunc(x) = "red" then
	timepris = SQLBlessDot(formatnumber(useMedarbejderTimepris, 2))
	else
	timepris = useMedarbejderTimepris
	end if	
	
		
		'Response.write "func" & func & " akt_showonfak " & akt_showonfak &"<br>"
		
		if func = "red" AND thisaktfunc(x) <> "red" then
		    
		    chk = ""
		
		else
		
		    if (thisaktfunc(x) <> "red" AND a > 0) OR (thisaktfunc(x) = "red" AND akt_showonfak = 1) then
		    '========================================================== 
		    'Denne kan være misvisende da en alternativ akt. (a = 0 akt)
		    'godt kan være slået til men med 0 timer
			    if hiddentimer <> 0 AND cint(thisaktfakbar(x)) = 1 then
			    chk = "CHECKED"
			    else
				    if (thisaktfunc(x) = "red" AND akt_showonfak = 1) then
				    chk = "CHECKED"
				    else
				    chk = ""
				    end if
			    end if
		    else
		    chk = ""
		    end if
		
		end if
		
		'**** Show ***
		strAktsubtotal = strAktsubtotal & "<input type=hidden name='FM_hidden_timerthis_"&x&"_"&a&"' id='FM_hidden_timerthis_"&x&"_"&a&"' value='"&hiddentimer&"'>"_
		&"<input type=checkbox name='FM_show_akt_"&x&"_"&a&"' id='FM_show_akt_"&x&"_"&a&"' value='1' "&chk&" onClick='setTimerTot2("&x&"), setBeloebTot2("&x&"), nulstilfaktimer("&x&","&a&")'>&nbsp;"
		'&"a: "& a &" hiddtiemr: "& hiddentimer &" showonfak: "& akt_showonfak &" func: "& func &" thisaktfunc(x) "& thisaktfunc(x) &"
		
		'**** Timer (antal) ***
		if nul = 1 then
			strAktsubtotal = strAktsubtotal &"<div style='position:relative; left:25px; top:-10px; width:60px;' name='timeprisdiv_"&x&"_"&a&"' id='timeprisdiv_"&x&"_"&a&"'><b>"& valueTimer &"</b></div>"
			strAktsubtotal = strAktsubtotal &"<input type='hidden' name='FM_timerthis_"&x&"_"&a&"' id='FM_timerthis_"&x&"_"&a&"' value='"&valueTimer&"'></td>"
		else
			strAktsubtotal = strAktsubtotal &"<input type=text name='FM_timerthis_"&x&"_"&a&"' id='FM_timerthis_"&x&"_"&a&"' onKeyup='tjektimer("&x&","&a&"), showborder2("&x&","&a&")' value='"&valueTimer&"' size=4 style='!border: 1px; background-color: #FFFFFF; font-size:10px; border-color: yellowgreen; border-style: solid;'>"_
			& "<br><input type='button' name='beregn2_"&x&"_"&a&"_a' id='beregn2_"&x&"_"&a&"_a' value='Calc' onclick='tjektimer("&x&","&a&"), enhedspris("&x&","&a&"), setBeloebThis2("&x&","&a&"), hideborder2("&x&","&a&")' style='font-size:8px'></td>"
	    end if
		
		'***** Aktnavn / Text *****
		if thisaktfunc(x) = "red" then
				
				strSQL2 = "SELECT beskrivelse FROM faktura_det WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &" AND enhedspris = "& SQLBless(useMedarbejderTimepris) 
				oRec2.open strSQL2, oConn, 3
					
				if not oRec2.EOF then
				thistxt = oRec2("beskrivelse")
				end if
				oRec2.close
				
				if a <> 0 then
				thistxt = thistxt
				else
					if func = "red" then
					thistxt = thistxt
					else
					thistxt = "Andet?"
					end if
				end if
				
		
		else
			if a <> 0 then
			thistxt = thisaktnavn(x) '&"#"& oRec3("tmnr") '& "x:" & x & "a:" & a & " mid "&  oRec3("mid")
			else
			thistxt = "Andet?"
			end if
		end if
		
		if nul = 1 then
		pd = "0px"
		else
		pd = "5px"
		end if
		
		strAktsubtotal = strAktsubtotal &"<td valign=top style='padding:2px 12px 2px "&pd&";'><textarea name='FM_aktbesk_"&x&"_"&a&"' id='FM_aktbesk_"&x&"_"&a&"' style='width:320px; height:40px; font-size:11px;'>"&thistxt&"</textarea></td>"
		'<input id='button_bold_"&x&"_"&a&"' type='button' value='< fed >' style=""font-size:8px;"" /><input id='button_kurs_"&x&"_"&a&"' type='button' value='< kursiv >' style=""font-size:8px;"" /><input id='button_br_"&x&"_"&a&"' type='button' value='< br >' style=""font-size:8px;""  onClick=""addbr('"&x&"', '"&a&"')"" /><br>
		
		
		'**** Enhedspris ***
		strAktsubtotal = strAktsubtotal & "<td valign=top width=60 align='right' style='padding:1px 4px 2px 2px;'>"
			
		
		if nul = 1 then
			strAktsubtotal = strAktsubtotal & "<div style='position:relative; left:0px; top:6px;' align='right' name='enhprisdiv_"&x&"_"&a&"' id='enhprisdiv_"&x&"_"&a&"'><b>"& timepris &"</b></div>"
			strAktsubtotal = strAktsubtotal & "<input type=hidden name='FM_enhedspris_"&x&"_"&a&"' id='FM_enhedspris_"&x&"_"&a&"' value='"&timepris&"'></td>"
		else
			strAktsubtotal = strAktsubtotal &"<input type=text name='FM_enhedspris_"&x&"_"&a&"' id='FM_enhedspris_"&x&"_"&a&"' onKeyup='tjektimepris("&x&"), showborder2("&x&","&a&")' value='"&timepris&"'"_ 
			&" style='width:45px; font-size:8px;'></td>"
		end if
		
		
		'*** Enheds angivelse ****
		 if thisaktfunc(x) = "red" then 
         thisakt_enhedsang = aktenhedsang
         else
         thisakt_enhedsang = 0
         end if
		
		select case thisakt_enhedsang
		case 0
		ehSel_nul = "SELECTED"
		ehSel_et = ""
		ehSel_to = ""
		case 1
		ehSel_nul = ""
		ehSel_et = "SELECTED"
		ehSel_to = ""
		case 2
		ehSel_nul = ""
		ehSel_et = ""
		ehSel_to = "SELECTED"
		case else
		ehSel_nul = "SELECTED"
		ehSel_et = ""
		ehSel_to = ""
		end select
		
		
		strAktsubtotal = strAktsubtotal & "<td valign=top style='padding:2px 2px 2px 2px;'>"_
		&"<select name='FM_akt_enh_"&x&"_"&a&"' id='FM_akt_enh_"&x&"_"&a&"' style=""width:70px; font-size:8px;"" onchange=""setenhedpaakt('"&a&"','"&x&"','2')"">"_
        &"<option value=0 "& ehSel_nul &">kr. pr. time</option>"_
        &"<option value=1 "& ehSel_et &">kr. pr. stk.</option>"_
        &"<option value=2 "& ehSel_to &">kr. pr. enhed</option>"_
        &"</select></td>"
		
		
		
		
		'*** Rabat ****
	    
	     if thisaktfunc(x) = "red" then 
         thisakt_rabat = aktrabat
         else
             if intRabat <> 0 then
             thisakt_rabat = (intRabat/100)
             else
             thisakt_rabat = 0
             end if
         end if
	    
	    if nul = 1 then
	    
	    strAktsubtotal = strAktsubtotal & "<td width=60 valign=top align='right'><div style='position:relative; left:0px; top:6px;' align='right' name='rabatdiv_"&x&"_"&a&"' id='rabatdiv_"&x&"_"&a&"'><b>"& (thisakt_rabat*100) &" %</b></div>"
	    strAktsubtotal = strAktsubtotal & "<input type=hidden name='FM_rabat_"&x&"_"&a&"' id='FM_rabat_"&x&"_"&a&"' value='"&thisakt_rabat&"'></td>"
		
	    else
        
         'onChange='enhedspris("&x&","&a&"), setBeloebThis2("&x&","&a&")'
         
            
         
            select case (thisakt_rabat*100)
            case 0
            akt_rSel0 = "SELECTED"
            akt_rSel10 = ""
            akt_rSel25 = ""
            akt_rSel40 = ""
            akt_rSel50 = ""
            akt_rSel60 = ""
            akt_rSel75 = ""
            case 10
            akt_rSel0 = ""
            akt_rSel10 = "SELECTED"
            akt_rSel25 = ""
            akt_rSel40 = ""
            akt_rSel50 = ""
            akt_rSel60 = ""
            akt_rSel75 = ""
            case 25
            akt_rSel0 = ""
            akt_rSel10 = ""
            akt_rSel25 = "SELECTED"
            akt_rSel40 = ""
            akt_rSel50 = ""
            akt_rSel60 = ""
            akt_rSel75 = ""
            case 40
            akt_rSel0 = ""
            akt_rSel10 = ""
            akt_rSel25 = ""
            akt_rSel40 = "SELECTED"
            akt_rSel50 = ""
            akt_rSel60 = ""
            akt_rSel75 = ""
            case 50
            akt_rSel0 = ""
            akt_rSel10 = ""
            akt_rSel25 = ""
            akt_rSel40 = ""
            akt_rSel50 = "SELECTED"
            akt_rSel60 = ""
            akt_rSel75 = ""
            case 60
            akt_rSel0 = ""
            akt_rSel10 = ""
            akt_rSel25 = ""
            akt_rSel40 = ""
            akt_rSel50 = ""
            akt_rSel60 = "SELECTED"
            akt_rSel75 = ""
            case 75
            akt_rSel0 = ""
            akt_rSel10 = ""
            akt_rSel25 = ""
            akt_rSel40 = ""
            akt_rSel50 = ""
            akt_rSel60 = ""
            akt_rSel75 = "SELECTED"
            end select
        
        strAktsubtotal = strAktsubtotal &"<td width=60 valign=top align='right' style='padding:2px 0px 0px 0px;'><select id='FM_rabat_"&x&"_"&a&"' name='FM_rabat_"&x&"_"&a&"' onChange='setbgcol("&x&");' style='background-color:#ffffff; width:50px; font-size:8px;'>"
        strAktsubtotal = strAktsubtotal &"<option value='0' "&akt_rSel0&">0%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.10' "&akt_rSel10&">10%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.25' "&akt_rSel25&">25%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.40' "&akt_rSel40&">40%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.50' "&akt_rSel50&">50%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.60' "&akt_rSel60&">60%</option>"
        strAktsubtotal = strAktsubtotal &"<option value='0.75' "&akt_rSel75&">75%</option></select>"
        
        strAktsubtotal = strAktsubtotal &"<div id='angivtxt_"&x&"_"&a&"' style='position:relative; left:-10px; top:2px; visibility:hidden; display:none;'><input id=asum_rbt_"&x&"_"&a&" type=button value='Match timepris' style='font-size:9px; width:100px;' onClick='genberegntimeprissumakt("&x&")'/></div></td>"
        
		end if                           
		
		
		
		
		'**** Beløb ****
		prisThis = 0
		if thisaktfunc(x) = "red" then 'func = "red" then 
				
					strSQL2 = "SELECT sum(aktpris) AS aktpris FROM faktura_det WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &" AND enhedspris = "& SQLBless(useMedarbejderTimepris) 
					oRec2.open strSQL2, oConn, 3
					
					if not oRec2.EOF then
						
						'** her er sidste rettelse **
						if len(oRec2("aktpris")) <> 0 then
						prisThis = oRec2("aktpris")
						end if
						
					end if
					oRec2.close
					
					
		else
			prisThis = valueTimer * timepris		
		end if
		
		
		'*** Rabatberegning ****
	    if thisaktfunc(x) = "red" then
		intRabatUse = prisThis 
		else
		intRabatUse = (prisThis) - (prisThis * (intRabat/100))
		end if
		
		if len(trim(intRabatUse)) <> 0 then
		intRabatUse = intRabatUse
		else
		intRabatUse = 0
		end if
		
		thisbel = SQLBlessDot(formatnumber(intRabatUse, 2))
		useforTotal = intRabatUse
		
		strAktsubtotal = strAktsubtotal & "<td valign=top width=80 style='padding:1px 0px 0px 2px;'>" 
		
		if nul = 1 then
			strAktsubtotal = strAktsubtotal &"<div style='position:relative; left:10px; top:6px; border-bottom:1px #8caae6 solid; width:50px;' align='right' name='belobdiv_"&x&"_"&a&"' id='belobdiv_"&x&"_"&a&"'><b>"& thisbel &"</b></div>"
			strAktsubtotal = strAktsubtotal &"<input type=hidden name='FM_beloebthis_"&x&"_"&a&"' id='FM_beloebthis_"&x&"_"&a&"' value='"&thisbel&"'></td>"
		else
			strAktsubtotal = strAktsubtotal &"&nbsp;&nbsp;<input type=text name='FM_beloebthis_"&x&"_"&a&"' id='FM_beloebthis_"&x&"_"&a&"' value='"&thisbel&"' onkeyup='tjekBeloeb("&x&","&a&"), setBeloebTot2("&x&","&a&")'"_
			&" style='width:50px; font-size:8px;'>"
		end if
		
		'*** Afslut tabel ***
		strAktsubtotal = strAktsubtotal &"</td></tr></table></div></td></tr>"
	
		
	if a >= 0 AND chk = "CHECKED" then
	intTimer = intTimer + valueTimer 'thisAkttimer(x)
	totalbelob = totalbelob + useforTotal 'cdbl(thisAktBeloeb(x))
	subtotalbelob = subtotalbelob + useforTotal 
	subtotaltimer = subtotaltimer + valueTimer
	'timeprisThisjob = thisTimePris(x)
	end if
	
	if func = "red" then
	isTPused = isTPused & ",#"& useMedarbejderTimepris &"#"
	end if
	
	hiddentimer = 0
	valuetimer = 0
	end function
	
	
	
	
	
	
	
	
	'*********************************************************************
	'******* Medarbejder linier ***********
	'*********************************************************************
	
	
	public medarbstrpris, faktot, mbelobtot, n, x, useMedarbejderTimepris
	public nialt, MedarbejderTimepris, chkMedarblinier
	function fakturaMedarbLinjer()
	                                    
	                                '*****************************
	                                '*** VenteTime Saldo Primo ***
	                                '***************************** 
	                                
	                                
	                                if id <> 0 then
	                                sqlFaknrKri =  " f.faknr < "& strFaknr &" AND "
	                                else
	                                sqlFaknrKri =  "  "
	                                end if
	                                
	                                
	                                
									strSQL2 = "SELECT f.fid, fms.venter FROM fakturaer f "_
									&" LEFT JOIN fak_med_spec fms ON (fms.fakid = f.fid AND fms.aktid = "& thisaktid(x) &""_
									&" AND fms.mid = "& oRec3("mid") &") "_
									&" WHERE "& sqlFaknrKri &" f.jobid = "& jobid &" ORDER BY f.faknr DESC, aktid" 
                                    
                                    'Response.Write strSQL2
                                    'Response.flush
                                    
                                    
                                    oRec2.open strSQL2, oConn, 3 
                                    if not oRec2.EOF then 
                                        venter = oRec2("venter")
                                    end if
                                    oRec2.close
                                    
                                    if len(trim(venter)) <> 0 then
                                    venter = venter 
                                    else
                                    venter = 0
                                    end if
	
	                                showventer = venter
	                           
	                           
	                           
	                            if thisaktfunc(x) = "red" then
									
									
								  
								    venter_brugt = oRec3("venter_brugt")
									venter_ultimo = oRec3("venter_ultimo")
									
									fak = oRec3("fak")
									medarbejderTimepris = oRec3("enhedspris")
									medarbejderBeloeb = oRec3("beloeb")
									txt = oRec3("tekst")
									medarbid = oRec3("mid")		
									showonfak = oRec3("showonfak")
									akt_showonfak = 0 
									
									
									
									if len(medarbejderTimepris) <> 0 then
									medarbejderTimepris = SQLBlessDOT(formatnumber(medarbejderTimepris, 2))
									else
									medarbejderTimepris = SQLBlessDot(formatnumber(0, 2))
									end if
									
									medrabat = (oRec3("medrabat") * 100)
									medenhedsang = oRec3("enhedsang")
									
									
									
									
								
								else
								
								    
										
										'***************************************************************
										'Timer medarbejder 
										'***************************************************************
										
										strSQL2 = "SELECT sum(timer) AS timer FROM timer WHERE Tmnr = "&oRec3("medarbejderid")&" "_
										&" AND taktivitetid ="& thisaktid(x) &" AND "_
										&" (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')"
										
										'AND tfaktim = 1 
										'OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"'
										
										'Response.Write strSQL2 & "<br>"
										oRec2.open strSQL2, oConn, 3 
										
										if not oRec2.EOF then 
											medarbejderTimer = oRec2("timer")
										end if
										oRec2.close
										
										
										'medarbejderTimer = oRec3("timer")
										if len(medarbejderTimer) <> 0 then
											medarbejderTimer = medarbejderTimer
										else
											medarbejderTimer = 0
										end if
										
										usedmedabId = usedmedabId & " Tmnr <> "& oRec3("medarbejderid") &" AND "
										
										'***************************************************************
										'Timepris pr. medarbejder ***
										'***************************************************************
										%>
										<!--#include file="fak_inc_mtimepris_opret.asp" -->
										<%
										
										'***************************************************************
										'Fak, Venter og Beløb
										'***************************************************************
										
										
										'call fakventerbelob(thisaktid(x), oRec3("mid"), oRec3("mnavn"), medarbejderTimer, medarbejderTimepris)
										
										
										fak = medarbejderTimer 'oRec3("timer")
                                        if len(fak) <> 0 then
                                        fak = fak
                                        else
                                        fak = 0
                                        end if


                                        venter_brugt = 0
                                        venter = 0
                                        venter_ultimo = 0

                                        medarbejderBeloeb = (fak * medarbejderTimepris) 
                                        txt = oRec3("mnavn") 
                                        medarbid = oRec3("mid")
                                        showonfak = 0
                                        										
										
											
								end if '** red / opr
								
							
							
							'**************************************************************************
							'Medarbejder timepris samme som forrige?? 
							'Kalder sum-aktivitet
							'**************************************************************************
							
								if instr(medarbstrpris, "#"&MedarbejderTimepris&"#") = 0 then
									newa = newa + 1
									Redim preserve aval(newa)
									aval(newa) = MedarbejderTimepris
									call subtotalakt(x, newa)
									sa = sa + 1
								end if
							'Response.write "a: "& newa  &"aval(a): "&aval(newa)&"<br>" 
							'***
							
							
							'** Henter tom sum-akt ***
							call subtotalakt(x, -(n))
							
							
							
							
							'***** Medarbejder timer på akt. timer skt. pris og omsætning  ************	
							
							if thisaktfunc(x) <> "red" AND flereTp = 1 then
							mbg = "#ffff99"
							else
							
							select case right(n, 1)
							case 0,2,4,6,8
							mbg = "#ffffff"
							case 1,3,5,7,9
							mbg = "lightgrey"
							end select
							
							end if
							%>	
							<tr bgcolor="<%=mbg %>">
								<td valign=top style="padding:2px 2px 0px 2px;"><input type="button" name="beregn_<%=n%>_<%=x%>_a" id="beregn_<%=n%>_<%=x%>_a" value="Calc" onClick="enhedsprismedarb('<%=x%>','<%=n%>'), hideborder('<%=x%>','<%=n%>')" style="font-size:8px;"></td>
								
								<!-- Antal -->
								<td valign=top style="padding:0px 2px 0px 2px;"><input type="hidden" name="FM_mid_<%=n%>_<%=x%>" value="<%=medarbid%>">
								<%
								if showonfak = 1 then
								chkshow = "CHECKED"
								else
									if chkMedarblinier = 1 AND thisaktfunc(x) <> "red" then
									chkshow = "CHECKED"
									else
									chkshow = ""
									end if
								end if
								
								'request.cookies("tvmedarb") = "j"
								
								'*** finder den næste a- værdi i rækken. Der må ikke være "huller"
								'*** da javascriptet så fejler
								
								if instr(medarbstrpris, "#"&MedarbejderTimepris&"#") = 0 then
								a = newa
								else
									
									for intcounter = 0 to newa - 1
										
										if MedarbejderTimepris = aval(intcounter) then
											a = intcounter
										end if
										
									next
								end if
								
								
								 
								%>
								<input type="hidden" name="FM_m_aval_opr_<%=n%>_<%=x%>" id="FM_m_aval_opr_<%=n%>_<%=x%>" value="<%=a%>">
								<input type="hidden" name="FM_m_aval_<%=n%>_<%=x%>" id="FM_m_aval_<%=n%>_<%=x%>" value="<%=a%>">
								
								<%
								'*** Kun medarbejdere med mere end 0 timer skal være checked, hvis dette ellers er tilvalgt. ****
								if (fak = 0 OR cint(thisaktfakbar(x)) <> 1) AND thisaktfunc(x) <> "red" then
								chkshow = ""
								else
								chkshow = chkshow
								end if
								%>
								
								<!-- Vis på faktura -->
								<input type="checkbox" name="FM_show_medspec_<%=n%>_<%=x%>" <%=chkshow%> id="FM_show_medspec_<%=n%>_<%=x%>" value="show">
								
								<!-- Timer til fakturering -->
								<input type="text" maxlength="<%=maxl%>" name="FM_m_fak_<%=n%>_<%=x%>" id="FM_m_fak_<%=n%>_<%=x%>" onKeyup="offsetfak('<%=x%>','<%=n%>'), tjektimer2('<%=x%>','<%=n%>'), showborder('<%=x%>','<%=n%>')" value="<%=fak%>" style="border:1px yellowgreen solid; font-size:9px; width:30px;">&nbsp;<font class=lillegray><%=fak%></font>
								<input type="hidden" name="FM_m_fak_opr_<%=n%>_<%=x%>" id="FM_m_fak_opr_<%=n%>_<%=x%>" value="<%=fak%>" style="font-size:8px;">
								</td>
								
								<!-- Vente timer -->
								<td align=center valign=top>
								<font class=lilleroed><%=showventer%></font> &nbsp;:
								<input type="text" name="FM_m_venterbrugt_<%=n%>_<%=x%>" id="FM_m_venterbrugt_<%=n%>_<%=x%>" onKeyup="offsetventer('<%=x%>','<%=n%>'), tjektimer4('<%=x%>','<%=n%>')" value="<%=venter_brugt %>" size=1 style="border:1px yellowgreen dashed; font-size:8px; width:25px;">
							   : <input type="text" name="FM_m_vent_<%=n%>_<%=x%>" id="FM_m_vent_<%=n%>_<%=x%>" onKeyup="offsetventer('<%=x%>','<%=n%>'), tjektimer4('<%=x%>','<%=n%>')" value="<%=venter_ultimo%>" size=1 style="border:1px #999999 dashed; font-size:8px; width:25px;">
							   
								<%if thisaktfunc(x) <> "red" then%>
								<input type="hidden" name="FM_m_vent_opr_<%=n%>_<%=x%>" id="FM_m_vent_opr_<%=n%>_<%=x%>" value="<%=fak%>">
								<%else%>
							    <input type="hidden" name="FM_m_vent_opr_<%=n%>_<%=x%>" id="FM_m_vent_opr_<%=n%>_<%=x%>" value="<%=venter%>">
								<%end if%>
								</td>
								
								<!-- Medarbejder Tekst -->
								<td valign=top><input type="text" name="FM_m_tekst_<%=n%>_<%=x%>" value="<%=txt%>" style="width:140px; font-size:9px;">
								<%if thisaktfunc(x) <> "red" AND flereTp = 1 then%>
								<%=tp%>
								<%end if%>
								</td>
								
								
								<!-- Pris pr. enhed -->
							    <td valign=top align=right style="padding:0px 2px 0px 2px;"><input type="hidden" name="FM_mtimepris_opr_<%=n%>_<%=x%>" id="FM_mtimepris_opr_<%=n%>_<%=x%>" value="<%=medarbejderTimepris%>">
								<input type="text" name="FM_mtimepris_<%=n%>_<%=x%>" id="FM_mtimepris_<%=n%>_<%=x%>" onKeyup="offsetmtp('<%=x%>','<%=n%>'), tjektimer3('<%=x%>','<%=n%>'), showborder('<%=x%>','<%=n%>')" value="<%=medarbejderTimepris%>" style="font-size:9px; width:50px;"> 
								<!-- onFocus="higlightakt('<=x%>','<=n%>','1')" -->
								</td>
								
								<!-- Enheds angivelse -->
								<%
								'*** Enheds angivelse ****
		                        if thisaktfunc(x) = "red" then 
                                thismed_enhedsang = medenhedsang
                                else
                                thismed_enhedsang = 0
                                end if
                        		
		                        select case thismed_enhedsang
		                        case 0
		                        ehSel_nul = "SELECTED"
		                        ehSel_et = ""
		                        ehSel_to = ""
		                        case 1
		                        ehSel_nul = ""
		                        ehSel_et = "SELECTED"
		                        ehSel_to = ""
		                        case 2
		                        ehSel_nul = ""
		                        ehSel_et = ""
		                        ehSel_to = "SELECTED"
		                        case else
		                        ehSel_nul = "SELECTED"
		                        ehSel_et = ""
		                        ehSel_to = ""
		                        end select
								
								%>
								<td valign=top align=right style="padding-right:10;">
								<select id="FM_med_enh_<%=n%>_<%=x%>" name="FM_med_enh_<%=n%>_<%=x%>" style="width:70px; font-family:verdana; font-size:8px;" onchange="setenhedpaakt('<%=n %>','<%=x %>','1')";>
                                    <option value=0 <%=ehSel_nul %>>kr. pr. time</option>
                                    <option value=1 <%=ehSel_et %>>kr. pr. stk.</option>
                                    <option value=2 <%=ehSel_to %>>kr. pr. enhed</option>
                                </select>
								</td>
								
								<!-- Rabat -->
								<td valign=top align=right>
                                <%
                                if thisaktfunc(x) = "red" then 
                                thismed_rabat = medrabat
                                else
                                thismed_rabat = intRabat
                                end if
                                
                                 
		                        select case thismed_rabat
		                        case 0
		                        rSel0 = "SELECTED"
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 10
		                        rSel0 = ""
		                        rSel10 = "SELECTED"
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 25
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = "SELECTED"
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
	                            case 40
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = "SELECTED"
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 50
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = "SELECTED"
		                        rSel60 = ""
		                        rSel75 = ""
		                        case 60
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = "SELECTED"
		                        rSel75 = ""
		                        case 75
		                        rSel0 = ""
		                        rSel10 = ""
		                        rSel25 = ""
		                        rSel40 = ""
		                        rSel50 = ""
		                        rSel60 = ""
		                        rSel75 = "SELECTED"
		                        end select
		                        %>
                                    <select id="FM_mrabat_<%=n%>_<%=x%>" name="FM_mrabat_<%=n%>_<%=x%>" style="width:50px; font-family:verdana; font-size:9px;" onChange="enhedsprismedarb('<%=x%>','<%=n%>')">
                                        <option value="0"  <%=rSel0%>>0%</option>
                                        <option value="0.10" <%=rSel10%>>10%</option>
                                        <option value="0.25" <%=rSel25%>>25%</option>
                                        <option value="0.40" <%=rSel40%>>40%</option>
                                        <option value="0.50" <%=rSel50%>>50%</option>
                                        <option value="0.60" <%=rSel60%>>60%</option>
                                        <option value="0.75" <%=rSel75%>>75%</option>
                                        </select>
								
								 </td>
								
								<!-- Beløb ialt. -->
								<td valign=top style="padding-left:2;">
								<%
								if thisaktfunc(x) = "red" then
								
								mbelob = medarbejderBeloeb
								
								else
								
								if len(medarbejderBeloeb) <> 0 then
								    if intRabat <> 0 then
								    mbelob = (medarbejderBeloeb) - (medarbejderBeloeb * (intRabat/100)) 
								    else
								    mbelob = medarbejderBeloeb
								    end if
								else
								mbelob = 0
								end if
								
								end if
								
								if len(trim(mbelob)) <> 0 then
								mbelob = mbelob
								else
								mbelob = 0
								end if
								%>
								<div style="position:relative; left:0px; top:0px; width:55px; border-bottom:1px #8caae6 solid; background-color:<%=mbg %>; padding-right:3px;" align="right" id="medarbbelobdiv_<%=n%>_<%=x%>"><b><%=SQLBlessDot(formatnumber(mbelob, 2))%></b></div>
								<input type="hidden" name="FM_mbelob_<%=n%>_<%=x%>" id="FM_mbelob_<%=n%>_<%=x%>" value="<%=SQLBlessDot(formatnumber(mbelob, 2))%>">
								</td>
								<td valign=top>
								<input type="button" value="+" onClick="decimaler('plus', '<%=n%>', '<%=x%>')" style="font-size:9px; background-color:#8caae6; width:14px; height:17;">
							<input type="button" value="-" onClick="decimaler('minus', '<%=n%>', '<%=x%>')" style="font-size:9px; background-color:#8caae6; width:10px; height:17;">
							</td>
							</tr>
							<% 
					
					medarbstrpris = medarbstrpris & ", #"&MedarbejderTimepris&"#"
					faktot = faktot + fak
					mbelobtot = mbelobtot + mbelob
					
					fak = 0 
					venter = 0
					medarbejderBeloeb = 0
					n = n + 1
	
	
	
	
	
	end function
	
	
    
    public 	gp1, gp2, gp3, gp4, gp5, gp6, gp7, gp8, gp9, gp10 
    function projektgrupper()
    
    
    
                            '**** projektgrupper på job ******************************************
							'Tjekker at der den projgp der findes på akt også findes på jobbet. 
							'Gælder kun job oprettet før 5/11-2004, da der her er lavet en 
							'funktion der tjekker det samme når et job redigeres. 
							'*********************************************************************
							
								jobPgstring = ""
							
								strSQL = "SELECT id, projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM job WHERE id = "& jobid
								oRec2.open strSQL, oConn, 3
									if not oRec2.EOF then
									
									strProjektgr1 = oRec2("projektgruppe1")
									strProjektgr2 = oRec2("projektgruppe2")
									strProjektgr3 = oRec2("projektgruppe3")
									strProjektgr4 = oRec2("projektgruppe4")
									strProjektgr5 = oRec2("projektgruppe5")
									strProjektgr6 = oRec2("projektgruppe6")
									strProjektgr7 = oRec2("projektgruppe7")
									strProjektgr8 = oRec2("projektgruppe8")
									strProjektgr9 = oRec2("projektgruppe9")
									strProjektgr10 = oRec2("projektgruppe10")
									
									jobPgstring = "#"&strProjektgr1&"#,#"&strProjektgr2&"#,#"&strProjektgr3&"#,#"&strProjektgr4&"#,#"&strProjektgr5&"#,#"&strProjektgr6&"#,#"&strProjektgr7&"#,#"&strProjektgr8&"#,#"&strProjektgr9&"#,#"&strProjektgr10&"#"
									
									end if
								oRec2.close
							
							
							
							'*** Medarbejder udspecificering ***
							strSQL3 = "SELECT projektgruppe1, projektgruppe2, projektgruppe3, projektgruppe4, projektgruppe5, projektgruppe6, projektgruppe7, projektgruppe8, projektgruppe9, projektgruppe10 FROM aktiviteter WHERE id = "& thisaktid(x)
							oRec3.open strSQL3, oConn, 3
							if not oRec3.EOF then
							
								gp1 = oRec3("projektgruppe1")
								gp2 = oRec3("projektgruppe2")
								gp3 = oRec3("projektgruppe3")
								gp4 = oRec3("projektgruppe4")
								gp5 = oRec3("projektgruppe5")
								gp6 = oRec3("projektgruppe6")
								gp7 = oRec3("projektgruppe7")
								gp8 = oRec3("projektgruppe8")
								gp9 = oRec3("projektgruppe9")
								gp10 = oRec3("projektgruppe10")
							
							end if
							oRec3.Close 
							
							'*** Findes gruppen også på jobbet?? ****
							if len(gp1) <> 0 AND instr(jobPgstring, "#"&gp1&"#") > 0 then
							gp1 = gp1
							else
							gp1 = 1
							end if
							
							if len(gp2) <> 0 AND instr(jobPgstring, "#"&gp2&"#") > 0 then
							gp2 = gp2
							else
							gp2 = 1
							end if
							
							if len(gp3) <> 0 AND instr(jobPgstring, "#"&gp3&"#") > 0 then
							gp3 = gp3
							else
							gp3 = 1
							end if
							
							if len(gp4) <> 0 AND instr(jobPgstring, "#"&gp4&"#") > 0 then
							gp4 = gp4
							else
							gp4 = 1
							end if
							
							if len(gp5) <> 0 AND instr(jobPgstring, "#"&gp5&"#") > 0 then
							gp5 = gp5
							else
							gp5 = 1
							end if
							
							if len(gp6) <> 0 AND instr(jobPgstring, "#"&gp6&"#") > 0 then
							gp6 = gp6
							else
							gp6 = 1
							end if
							
							if len(gp7) <> 0 AND instr(jobPgstring, "#"&gp7&"#") > 0 then
							gp7 = gp7
							else
							gp7 = 1
							end if
							
							if len(gp8) <> 0 AND instr(jobPgstring, "#"&gp8&"#") > 0 then
							gp8 = gp8
							else
							gp8 = 1
							end if
							
							if len(gp9) <> 0 AND instr(jobPgstring, "#"&gp9&"#") > 0 then
							gp9 = gp9
							else
							gp9 = 1
							end if
							
							if len(gp10) <> 0 AND instr(jobPgstring, "#"&gp10&"#") > 0 then
							gp10 = gp10
							else
							gp10 = 1
							end if
    
    
    
    
    end function
    
    
    public matSubTotal, matSubTotalAll
    public rSel0, rSel10, rSel25, rSel40, rSel50, rSel60
    
    
    function materialer(sw)
    
    if func = "red" then
        if sw = 2 then
        chkVis = ""
        else
        chkVis = "CHECKED"
        end if
    else
    chkVis = "CHECKED"
    end if
    
    %>
    <tr>
        
        <td style="padding:0px 2px 0px 2px;"><input type="button" name="beregn_<%=m%>_m" id="beregn_<%=m%>_m" value="Calc" onClick="beregnmatpris('<%=m%>','1'), hidebordermat('<%=m%>')" style="font-size:8px;"></td>
        <td>
            <input id="FM_vis_<%=m %>" name="FM_vis_<%=m %>" value="1" <%=chkVis %> type="checkbox" onclick="beregnmatpris('<%=m%>','1')" />
            <input id="FM_matid_<%=m %>" name="FM_matid_<%=m %>" value="<%=oRec("matid") %>" type="hidden" /></td>
        <td>
            <input id="FM_matnavn_<%=m %>" name="FM_matnavn_<%=m %>" type="text" value="<%=oRec("matnavn") %>" style="font-family:verdana; font-size:9px; width:275px;" /></td>
        <td>
            <input id="FM_matvarenr_<%=m %>" name="FM_matvarenr_<%=m %>" type="text" value="<%=oRec("matvarenr") %>" style="font-family:verdana; font-size:9px; width:60px;" /></td>
        
        <td>
            <input id="FM_matantal_opr_<%=m %>" type="hidden" value="<%=oRec("matantal") %>" />
            <input id="FM_matantal_<%=m %>" name="FM_matantal_<%=m %>" type="text" value="<%=SQLBlessDOT(formatnumber(oRec("matantal"), 2)) %>" onkeyup="offsetmatant('<%=m %>'), tjekmatantal('<%=m %>'), showbordermat('<%=m%>')" style="font-family:verdana; font-size:9px; width:50px;" /></td>
        <td> 
            <input id="FM_matenhedspris_opr_<%=m %>" type="hidden" value="<%=oRec("matsalgspris") %>" />
            <input id="FM_matenhedspris_<%=m %>" name="FM_matenhedspris_<%=m %>" type="text" value="<%=SQLBlessDOT(formatnumber(oRec("matsalgspris"), 2)) %>" onkeyup="offsetmatenhpris('<%=m %>'), tjekmatehpris('<%=m %>'), showbordermat('<%=m%>')" style="font-family:verdana; font-size:9px; width:80px;" /></td>
        <td> <input id="FM_matenhed_<%=m %>" name="FM_matenhed_<%=m %>" type="text" value="<%=oRec("matenhed") %>" style="font-family:verdana; font-size:9px; width:40px;" /></td>
        <td>
                <select id="FM_matrabat_<%=m%>" name="FM_matrabat_<%=m%>" style="width:50px; font-family:verdana; font-size:9px;" onChange="beregnmatpris('<%=m%>','1')">
                                        <option value="0"  <%=rSel0%>>0%</option>
                                        <option value="0.10" <%=rSel10%>>10%</option>
                                        <option value="0.25" <%=rSel25%>>25%</option>
                                        <option value="0.40" <%=rSel40%>>40%</option>
                                        <option value="0.50" <%=rSel50%>>50%</option>
                                        <option value="0.60" <%=rSel60%>>60%</option>
                                        <option value="0.75" <%=rSel75%>>75%</option>
                                        </select>
        <td>
        <%
        if intRabat <> 0 then
        matSubTotal = ((oRec("matantal") * oRec("matsalgspris")) - (oRec("matantal") * oRec("matsalgspris") * (intRabat/100)))
        else
        matSubTotal = oRec("matantal") * oRec("matsalgspris")
        end if
        
        if len(trim(matSubTotal)) <> 0 then
        matSubTotal = matSubTotal
        else
        matSubTotal = 0
        end if
        
        %>
        <input id="FM_matbeloeb_<%=m %>" name="FM_matbeloeb_<%=m %>" type="hidden" value="<%=SQLBlessDOT(formatnumber(matSubTotal, 2)) %>" style="font-family:verdana; font-size:9px; width:80px;" />
        <div id="matbelobdiv_<%=m %>" style="position:relative; left:0px; top:1px; width:50px; font-size:10px; border-bottom:1px #5582d2 solid;" align="right"><%=SQLBlessDOT(formatnumber(matSubTotal, 2)) %></div>
        </td>
       
    </tr>
    <%
    if chkVis = "CHECKED" then
    matSubTotalAll = matSubTotalAll + matSubTotal
    end if
    
    end function
    
    
    
    
    
    
   
			
%>