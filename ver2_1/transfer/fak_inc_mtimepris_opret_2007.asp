<%
'***************************************************************'
'Timepris pr. medarbejder                                    ***'
'***************************************************************'




'if medarbejderTimer = 0 then
medarbejderTimepris = ""
'end if


if usefastpris(x) = 1 then

	
	
	'*************************************************************************************'
    '*** Finder timepris på aktivitet og job for den valgte medarbejder *****'
    '********************************************************************************'
	
	 '*** aktiviteten ***'
			ftp = 0
			
			strSQL2 = "SELECT timeprisalt, 6timepris, 6valuta, valutakode, medarbejdertype, v.kurs FROM timepriser "_
			&" LEFT JOIN medarbejdere ON (mid = medarbid) "_
			&" LEFT JOIN valutaer v ON (v.id = 6valuta) "_
			&" WHERE "_
			&" medarbid = "&oRec3("medarbejderid")&" AND aktid ="& thisaktid(x) 
			oRec2.open strSQL2, oConn, 3 
			if not oRec2.EOF then
			
			
				if oRec2("timeprisalt") <> 6 then
						
				call timeprisertilFak(1, oRec2("medarbejdertype"))
							
					
				else
				'medarbejderTimepris = oRec2("6timepris")
				valutaKode = oRec2("valutakode")
		        valutaID = oRec2("6valuta")
		        valutaKodeKurs = oRec2("kurs")
				
				ftp = 1
				end if
			end if
			oRec2.close
			
			
			'** Hvis der ikke findes en timepris på aktiviteten, bruges timeprisen fra jobbet. **'
			if ftp <> 1 then 
			'** jobbet **
			
			strSQL2 = "SELECT timeprisalt, 6timepris, 6valuta, valutakode, medarbejdertype, v.kurs FROM timepriser "_
			&" LEFT JOIN medarbejdere ON (mid = medarbid) "_
			&" LEFT JOIN valutaer v ON (v.id = 6valuta) "_
			&" WHERE medarbid = "&oRec3("medarbejderid")&" AND aktid = 0 AND jobid = "& jobid
			 
			oRec2.open strSQL2, oConn, 3 
			if not oRec2.EOF then
			
				if oRec2("timeprisalt") <> 6 then
						
				call timeprisertilFak(2, oRec2("medarbejdertype"))
			    else
				
				'medarbejderTimepris = oRec2("6timepris")
				valutaKode = oRec2("valutakode")
		        valutaID = oRec2("6valuta")
		        valutaKodeKurs = oRec2("kurs")
		        
				end if
				
				'ftp = 2 
			end if
			oRec2.close
		end if
		
		'** Ved fastpris overskrives medabejder timepris med den beregnede timepris fra job **'
		medarbejderTimepris = akttimepris
		

else	
		
		
		'** Medarbejder timepris kun fakturerbare aktivteter **'
		strSQL2 = "SELECT timepris AS medarbejderTimepris, valutakode, t.valuta, v.kurs FROM timer t "_
		&" LEFT JOIN valutaer v ON (v.id = t.valuta) WHERE "_
		&" Tmnr = "& oRec3("medarbejderid") &" AND taktivitetid ="& thisaktid(x) &""_
		&" AND tfaktim = 1 AND (Tdato BETWEEN '" & stdatoKri &"' AND '"& slutdato &"') "_
		&" ORDER BY timepris" 
		
		
		oRec2.open strSQL2, oConn, 3 
		tp = "  "
		while not oRec2.EOF 
			
			medarbejderTimepris = oRec2("medarbejderTimepris")
			valutaKode = oRec2("valutakode")
			valutaID = oRec2("valuta")
			valutaKodeKurs = oRec2("kurs")
			
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
		
		
		
		'** Hvis ikke der findes timeregistreringer bruges  **'
		'** den aktuelle timepris på aktiviteten            **'
		'** for hver enkelt medarbejder.                    **'
		
		if len(medarbejderTimepris) > 0 then '*AND noregs = 1 then*'
		medarbejderTimepris = medarbejderTimepris
		valutaKode = valutaKode
		valutaID = valutaID  
		valutaKodeKurs = valutaKodeKurs
		else
			
			
			'*************************************************************************************'
            '*** Finder timepris på aktivitet og job for den valgte medarbejder *****'
            '********************************************************************************'
			
			 '*** aktiviteten ***'
			ftp = 0
			
			strSQL2 = "SELECT timeprisalt, 6timepris, 6valuta, valutakode, medarbejdertype, v.kurs FROM timepriser "_
			&" LEFT JOIN medarbejdere ON (mid = medarbid) "_
			&" LEFT JOIN valutaer v ON (v.id = 6valuta) "_
			&" WHERE "_
			&" medarbid = "&oRec3("medarbejderid")&" AND aktid ="& thisaktid(x) 
			oRec2.open strSQL2, oConn, 3 
			if not oRec2.EOF then
			
			
				if oRec2("timeprisalt") <> 6 then
						
				call timeprisertilFak(1, oRec2("medarbejdertype"))
							
					
				else
				medarbejderTimepris = oRec2("6timepris")
				valutaKode = oRec2("valutakode")
		        valutaID = oRec2("6valuta")
		        valutaKodeKurs = oRec2("kurs")
				
				ftp = 1
				end if
			end if
			oRec2.close
			
			
			'** Hvis der ikke findes en timepris på aktiviteten, bruges timeprisen fra jobbet. **'
			if ftp <> 1 then 
			'** jobbet **
			
			strSQL2 = "SELECT timeprisalt, 6timepris, 6valuta, valutakode, medarbejdertype, v.kurs FROM timepriser "_
			&" LEFT JOIN medarbejdere ON (mid = medarbid) "_
			&" LEFT JOIN valutaer v ON (v.id = 6valuta) "_
			&" WHERE medarbid = "&oRec3("medarbejderid")&" AND aktid = 0 AND jobid = "& jobid
			 
			oRec2.open strSQL2, oConn, 3 
			if not oRec2.EOF then
			
				if oRec2("timeprisalt") <> 6 then
						
				call timeprisertilFak(2, oRec2("medarbejdertype"))
			    else
				
				medarbejderTimepris = oRec2("6timepris")
				valutaKode = oRec2("valutakode")
		        valutaID = oRec2("6valuta")
		        valutaKodeKurs = oRec2("kurs")
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
