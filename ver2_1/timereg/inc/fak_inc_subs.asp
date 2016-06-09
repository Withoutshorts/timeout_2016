<%

public timeforbrug 
public lastFakdato
public strfaktidspkt
function faktureredetimerogbelob()

		strSQL = "SELECT fid, beloeb, timer, fakdato, tidspunkt, faktype FROM fakturaer WHERE jobid = " & jobid &" ORDER BY fakdato"
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

function fakventerbelob(aktid, medid, mednavn, medtim, medtpris)
strSQL2 = "SELECT sum(venter) AS venter FROM fak_med_spec WHERE aktid = "& aktid &" AND mid = "& medid &" ORDER BY aktid" 
'Response.write strSQL2
oRec2.open strSQL2, oConn, 3 
if not oRec2.EOF then 
	venter = oRec2("venter")
end if
oRec2.close

if len(venter) <> 0 then
venter = venter 
else
venter = 0
end if


fak = medtim '+ venter 'oRec3("timer")
if len(fak) <> 0 then
fak = fak
else
fak = 0
end if

showventer = venter
if len(showventer) <> 0 then
showventer = showventer
else
showventer = 0
end if
venter = 0


medarbejderBeloeb = (fak * medtpris) 'oRec3("timepris")
txt = mednavn
medarbid = medid
showonfak = 0
end function















'***********************************************************************************
'Sum-Aktiviteter subtotal func
'***********************************************************************************
	
	public totalbelob
	public intTimer
	public strAktsubtotal
	public subtotalbelob
	public subtotaltimer
	
	
	function subtotalakt(x, a)
	
	
	if a < 0 then
	useMedarbejderTimepris = -1
	'thisAkttimer(x) = ""
	'thisAktBeloeb(x) = ""
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
	
	
		if func = "red" then
				'*** Finder timepris på ekstra aktiviteten **
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
				
				
				'** Finder showonfak på 0 aktiviteten **
				strSQL2 = "SELECT showonfak FROM faktura_det WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &" AND enhedspris = "& SQLBless(useMedarbejderTimepris) 'SQLBless(thisTimePris(x)) 
				oRec2.open strSQL2, oConn, 3
				if not oRec2.EOF then
				akt_showonfak = oRec2("showonfak")
				end if
				oRec2.close
				
		else
		useMedarbejderTimepris = -2
		thisAkttimer(x) = 0
		thisAktBeloeb(x) = 0
		end if
	
	nul = 2
	bgcol = "#ffffff"
	bgcol2 = "#FFFFFF"
	hidden = "visible"
	display = ""
	end if
	
	if a > 0 then
	useMedarbejderTimepris = MedarbejderTimepris
	nul = 1
	bgcol = "#FFFFFF"
	bgcol2 = "#FFFFFF"
	
				
				if func = "red" then
						
						'** Finder showonfak på aktiviteten **
						strSQL2 = "SELECT showonfak FROM faktura_det WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &" AND enhedspris = "& SQLBless(useMedarbejderTimepris) 'SQLBless(thisTimePris(x)) 
						oRec2.open strSQL2, oConn, 3
						if not oRec2.EOF then
							akt_showonfak = oRec2("showonfak")
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
	
		
		strAktsubtotal = strAktsubtotal & "<tr><td valign=top colspan=6 bgcolor=#ffffff><div name='sumaktdiv_"&x&"_"&a&"' id='sumaktdiv_"&x&"_"&a&"' style='position:relative; visibility:"&hidden&"; display:"&display&"; background-color:"&bgcol2&";'>"_
		&"<table cellspacing=0 cellpadding=0 border=0 bgcolor='#FFFFFF'><tr>"
		
		strAktsubtotal = strAktsubtotal &"<td valign=top width='120' bgcolor='#ffffff'>"
		strAktsubtotal = strAktsubtotal &"<input type=hidden name='FM_aktid_"&x&"_"&a&"' value='"&thisaktid(x)&"'>"_
		&"<input type=hidden name='FM_hidden_timepristhis_"&x&"_"&a&"' id='FM_hidden_timepristhis_"&x&"_"&a&"' value='"&useMedarbejderTimepris&"'>"
		
		'*** Beregner timer ***********
		'===
		if func = "red" then
				
				
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
						
						
						strSQL2 = "SELECT sum(timer.timer) AS antaltimer FROM timer WHERE taktivitetid =" & thisaktid(x) &" AND tfaktim = 1 AND timepris = "&SQLBless(useMedarbejderTimepris)&" AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"')" 
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
		
		
		
	if func = "red" then
	timepris = SQLBlessDot(formatnumber(useMedarbejderTimepris, 2))
	else
	timepris = useMedarbejderTimepris
	end if	
	
		
		'Response.write "func" & func & " akt_showonfak " & akt_showonfak &"<br>"
		if (func <> "red" AND a > 0) OR (func = "red" AND akt_showonfak = 1) then
		'========================================================== 
		'Denne kan være misvisende da en alternativ akt. (a = 0 akt)
		'godt kan være slået til men med 0 timer
			if hiddentimer <> 0 then
			chk = "CHECKED"

              

			else
				if (func = "red" AND akt_showonfak = 1) then
				chk = "CHECKED"
				else
				chk = ""
				end if
			end if
		else
		chk = ""
		end if


        
                    

		
		'**** Show ***
		strAktsubtotal = strAktsubtotal & "<input type=hidden name='FM_hidden_timerthis_"&x&"_"&a&"' id='FM_hidden_timerthis_"&x&"_"&a&"' value='"&hiddentimer&"'>"_
		&"<input type=checkbox name='FM_show_akt_"&x&"_"&a&"' id='FM_show_akt_"&x&"_"&a&"' value='1' "&chk&" onClick='setTimerTot2("&x&"), setBeloebTot2("&x&"), nulstilfaktimer("&x&","&a&")'>&nbsp;"
		
		
		'**** Timer (antal) ***
		if nul = 1 then
			strAktsubtotal = strAktsubtotal &"<div style='position:relative; left:-55; top:-17;' align='right' name='timeprisdiv_"&x&"_"&a&"' id='timeprisdiv_"&x&"_"&a&"'><b>"& valueTimer &"</b></div>"_
			&"<input type=hidden name='FM_timerthis_"&x&"_"&a&"' id='FM_timerthis_"&x&"_"&a&"' value='"&valueTimer&"'>"_
			&"</td>"
		else
			strAktsubtotal = strAktsubtotal &"<input type=text name='FM_timerthis_"&x&"_"&a&"' id='FM_timerthis_"&x&"_"&a&"' onKeyup='tjektimer("&x&","&a&"), showborder2("&x&","&a&")' value='"&valueTimer&"' size=4 style='!border: 1px; background-color: #FFFFFF; border-color: limegreen; border-style: solid;'>"_
			& "&nbsp;<input type='button' name='beregn2_"&x&"_"&a&"_a' id='beregn2_"&x&"_"&a&"_a' value='Calc' onClick='tjektimer("&x&","&a&"), enhedspris("&x&","&a&"), setBeloebThis2("&x&","&a&"), hideborder2("&x&","&a&")' style='font-size:8px'></td>"
		end if
		
		'***** Aktnavn / Text *****
		if func = "red" then
				
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
		
		
		strAktsubtotal = strAktsubtotal &"<td valign=top width=400 bgcolor='#ffffff'><textarea rows=2 name='FM_aktbesk_"&x&"_"&a&"' id='FM_aktbesk_"&x&"_"&a&"' style='background-color: "&bgcol&"; width:400;'>"&thistxt&"</textarea>&nbsp;&nbsp;</td>"
		
		
		'**** Enhedspris ***
		strAktsubtotal = strAktsubtotal & "<td bgcolor='#ffffff' valign=top align='right' width=75>"
			
		
		if nul = 1 then
			strAktsubtotal = strAktsubtotal & "<div style='position:relative; left:-22; top:1;' align='right' name='enhprisdiv_"&x&"_"&a&"' id='enhprisdiv_"&x&"_"&a&"'><b>"& timepris &"</b></div>"
			strAktsubtotal = strAktsubtotal & "<input type=hidden name='FM_enhedspris_"&x&"_"&a&"' id='FM_enhedspris_"&x&"_"&a&"' value='"&timepris&"'></td>"
		else
			strAktsubtotal = strAktsubtotal &"<input type=text name='FM_enhedspris_"&x&"_"&a&"' id='FM_enhedspris_"&x&"_"&a&"' onKeyup='tjektimepris("&x&"), showborder2("&x&","&a&")' value='"&timepris&"' size=4"_ 
			&" style='!border: 1px; background-color: #FFFFFF; border-color: #999999; border-style: solid;'>"_
			& "&nbsp;<input type='button' name='beregn2_"&x&"_"&a&"_b' id='beregn2_"&x&"_"&a&"_b' value='Calc' onClick='enhedspris("&x&","&a&"), setBeloebThis2("&x&","&a&"), hideborder2("&x&","&a&")' style='font-size:8px'>"_
			&"&nbsp;</td>"
		end if
		
		
		'**** Beløb ****
		thisbel = 0
		if func = "red" then 
				
					strSQL2 = "SELECT sum(aktpris) AS aktpris FROM faktura_det WHERE fakid = "& id &" AND aktid = "& thisaktid(x) &" AND enhedspris = "& SQLBless(useMedarbejderTimepris) 
					oRec2.open strSQL2, oConn, 3
					
					if not oRec2.EOF then
						
						'** her er sidste rettelse **
						if len(oRec2("aktpris")) <> 0 then
						thisbel = SQLBlessDot(formatnumber(oRec2("aktpris"), 2))
						useforTotal = oRec2("aktpris")
						end if
						
					end if
					oRec2.close
					
					if len(thisbel) <> 0 then
					thisbel = thisbel
					useforTotal = useforTotal
					else
					thisbel = 0
					useforTotal = 0
					end if
					
		else
			thisbel = SQLBlessDot(formatnumber(valueTimer * timepris, 2))
			useforTotal = valueTimer * timepris		
		end if
		
		
		strAktsubtotal = strAktsubtotal & "<td valign=top valign=top width=90 align=right bgcolor='#ffffff'>" 
		
		if nul = 1 then
			strAktsubtotal = strAktsubtotal &"<div style='position:relative; left:-8;' align='right' name='belobdiv_"&x&"_"&a&"' id='belobdiv_"&x&"_"&a&"'><b>"& thisbel &"</b></div>"_
			&"<input type=hidden name='FM_beloebthis_"&x&"_"&a&"' id='FM_beloebthis_"&x&"_"&a&"' value='"&thisbel&"'>"_
			&"</td>"
		else
			strAktsubtotal = strAktsubtotal &"&nbsp;&nbsp;<input type=text name='FM_beloebthis_"&x&"_"&a&"' id='FM_beloebthis_"&x&"_"&a&"' value='"&thisbel&"' onkeyup='tjekBeloeb("&x&","&a&"), setBeloebTot2("&x&","&a&")'"_
			&" size='8' style='!border: 1px; background-color: #FFFFFF; border-color: #86B5E4; border-style: solid;'>"
		end if
		
		strAktsubtotal = strAktsubtotal &"</td></tr></table></div></td></tr>"
	
		
	if a >= 0 then
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
	
	
			
%>