<%
'***************************************************************
'Timepris pr. medarbejder ***
'***************************************************************

'if medarbejderTimer = 0 then
medarbejderTimepris = ""
'end if


if usefastpris(x) = 1 then

	medarbejderTimepris = akttimepris

else	
		
		strSQL2 = "SELECT timepris AS medarbejderTimepris FROM timer WHERE Tmnr = "&oRec3("medarbejderid")&" AND taktivitetid ="& thisaktid(x) &" AND tfaktim = 1 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"') ORDER BY timepris" 'OR Tdato = '" & stdatoKri &"' AND timer.tidspunkt > '"& LastFakTid &"'
		oRec2.open strSQL2, oConn, 3 
		tp = "  "
		while not oRec2.EOF 
			
			medarbejderTimepris = oRec2("medarbejderTimepris")
			
			if instr(tp, medarbejderTimepris&",") = 0 then
			tp = tp & medarbejderTimepris & ", "
			end if
			
		oRec2.movenext
		wend 
		oRec2.close
		
		len_tp = len(tp)
		tp = left(tp, len_tp - 2)
		
		if instr(tp, ",") <> 0 then
		tp = "<font color=red><b>!</b></font> <font class=megetlillesort> Timepriser: ("&trim(tp)&")</font>"
		flereTp = 1
		else
		tp = ""
		flereTp = 0
		end if
		
		
		
		'** Hvis ikke der findes timeregistreringer bruges
		'** den aktuelle timepris på aktiviteten
		'** for hver enkelt medarbejder.
		
		if len(medarbejderTimepris) > 0 then 'AND noregs = 1 then
		medarbejderTimepris = medarbejderTimepris
		else
			
			'*** aktiviteten ***
			ftp = 0
			
			strSQL2 = "SELECT timeprisalt, 6timepris, medarbejdertype FROM timepriser LEFT JOIN medarbejdere ON (mid = medarbid) WHERE medarbid = "&oRec3("medarbejderid")&" AND aktid ="& thisaktid(x) 
			oRec2.open strSQL2, oConn, 3 
			if not oRec2.EOF then
			
			
				if oRec2("timeprisalt") <> 6 then
						
						strSQL4 = "SELECT timepris, timepris_a1, timepris_a2, timepris_a3, timepris_a4, timepris_a5 FROM medarbejdertyper WHERE id = "& oRec2("medarbejdertype")
						oRec4.open strSQL4, oConn, 3 
						if not oRec4.EOF then
							select case oRec2("timeprisalt")
							case 0
							medarbejderTimepris = oRec4("timepris")
							case 1
							medarbejderTimepris = oRec4("timepris_a1")
							case 2
							medarbejderTimepris = oRec4("timepris_a2")
							case 3
							medarbejderTimepris = oRec4("timepris_a3")
							case 4
							medarbejderTimepris = oRec4("timepris_a4")
							case 5
							medarbejderTimepris = oRec4("timepris_a5")
							end select
							
							ftp = 1		
						end if
						oRec4.close 
				else
				medarbejderTimepris = oRec2("6timepris")
				ftp = 1
				end if
			end if
			oRec2.close
			
			
			'** Hvis der ikke findes en timepris på aktiviteten, bruges timeprisen fra jobbet. 
			if ftp <> 1 then 
			'** jobbet **
			strSQL2 = "SELECT timeprisalt, 6timepris, medarbejdertype FROM timepriser LEFT JOIN medarbejdere ON (mid = medarbid) WHERE medarbid = "&oRec3("medarbejderid")&" AND aktid = 0 AND jobid = "& jobid
			 
			oRec2.open strSQL2, oConn, 3 
			if not oRec2.EOF then
			
				if oRec2("timeprisalt") <> 6 then
						
						strSQL4 = "SELECT timepris, timepris_a1, timepris_a2, timepris_a3, timepris_a4, timepris_a5 FROM medarbejdertyper WHERE id = "& oRec2("medarbejdertype")
						oRec4.open strSQL4, oConn, 3 
						if not oRec4.EOF then
							select case oRec2("timeprisalt")
							case 0
							medarbejderTimepris = oRec4("timepris")
							case 1
							medarbejderTimepris = oRec4("timepris_a1")
							case 2
							medarbejderTimepris = oRec4("timepris_a2")
							case 3
							medarbejderTimepris = oRec4("timepris_a3")
							case 4
							medarbejderTimepris = oRec4("timepris_a4")
							case 5
							medarbejderTimepris = oRec4("timepris_a5")
							end select
							
									
						end if
						oRec4.close 
				else
				medarbejderTimepris = oRec2("6timepris")
				end if
				
				'ftp = 2 
			end if
			oRec2.close
			end if
			
		end if
end if 'fastpris

if len(medarbejderTimepris) <> 0 then
medarbejderTimepris = SQLBlessDot(formatnumber(medarbejderTimepris, 2))
else
medarbejderTimepris = SQLBlessDot(formatnumber(0, 2))
end if
%>
