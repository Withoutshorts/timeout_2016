<%

		
    lastJobnr = 0
	fakTotalBeloeb = 0
	faktotalTimer = 0

	faktimerJan = 0
	fakbeloebJan = 0
	
	faktimerFeb = 0
	fakbeloebFeb = 0
	
	faktimerMar = 0
	fakbeloebMar = 0
	
	faktimerApr = 0
	fakbeloebApr = 0
	
	faktimerMaj = 0
	fakbeloebMaj = 0
	
	faktimerJun = 0
	fakbeloebJun = 0
	
	faktimerJul = 0
	fakbeloebJul = 0
	
	faktimerAug = 0
	fakbeloebAug = 0
	
	faktimerSep = 0
	fakbeloebSep = 0
	
	faktimerOkt = 0
	fakbeloebOkt = 0
	
	faktimerNov = 0
	fakbeloebNov = 0
	
	faktimerDec = 0
	fakbeloebDec = 0
		
		
		dim mrdOmsBudgetTOT
		mrdOmsBudgetTOT = 0
		dim mrdOmsBudget(12)
		mrdOmsBudget(1) = 0
		mrdOmsBudget(2) = 0
		mrdOmsBudget(3) = 0
		mrdOmsBudget(4) = 0
		mrdOmsBudget(5) = 0
		mrdOmsBudget(6) = 0
		mrdOmsBudget(7) = 0
		mrdOmsBudget(8) = 0
		mrdOmsBudget(9) = 0
		mrdOmsBudget(10) = 0
		mrdOmsBudget(11) = 0
		mrdOmsBudget(12) = 0
		
		dim mrdAftBudgetTOT
		mrdAftBudgetTOT = 0
		dim mrdAftBudget(12)
		mrdAftBudget(1) = 0
		mrdAftBudget(2) = 0
		mrdAftBudget(3) = 0
		mrdAftBudget(4) = 0
		mrdAftBudget(5) = 0
		mrdAftBudget(6) = 0
		mrdAftBudget(7) = 0
		mrdAftBudget(8) = 0
		mrdAftBudget(9) = 0
		mrdAftBudget(10) = 0
		mrdAftBudget(11) = 0
		mrdAftBudget(12) = 0
		
		
		
		public mrdNotFakTimer(12)
		mrdNotFakTimer(1) = 0
		mrdNotFakTimer(2) = 0
		mrdNotFakTimer(3) = 0
		mrdNotFakTimer(4) = 0
		mrdNotFakTimer(5) = 0
		mrdNotFakTimer(6) = 0
		mrdNotFakTimer(7) = 0
		mrdNotFakTimer(8) = 0
		mrdNotFakTimer(9) = 0
		mrdNotFakTimer(10) = 0
		mrdNotFakTimer(11) = 0
		mrdNotFakTimer(12) = 0
		
		public mrdFakTimer(12)
		mrdFakTimer(1) = 0
		mrdFakTimer(2) = 0
		mrdFakTimer(3) = 0
		mrdFakTimer(4) = 0
		mrdFakTimer(5) = 0
		mrdFakTimer(6) = 0
		mrdFakTimer(7) = 0
		mrdFakTimer(8) = 0
		mrdFakTimer(9) = 0
		mrdFakTimer(10) = 0
		mrdFakTimer(11) = 0
		mrdFakTimer(12) = 0
		
		public mrdOms(12)
		mrdOms(1) = 0
		mrdOms(2) = 0
		mrdOms(3) = 0
		mrdOms(4) = 0
		mrdOms(5) = 0
		mrdOms(6) = 0
		mrdOms(7) = 0
		mrdOms(8) = 0
		mrdOms(9) = 0
		mrdOms(10) = 0
		mrdOms(11) = 0
		mrdOms(12) = 0
		
		public mrdKostpris(12)
		mrdKostpris(1) = 0
		mrdKostpris(2) = 0
		mrdKostpris(3) = 0
		mrdKostpris(4) = 0
		mrdKostpris(5) = 0
		mrdKostpris(6) = 0
		mrdKostpris(7) = 0
		mrdKostpris(8) = 0
		mrdKostpris(9) = 0
		mrdKostpris(10) = 0
		mrdKostpris(11) = 0
		mrdKostpris(12) = 0

		
		
		
		
		

if kli > 5 then
usescala = "1000"
divider = 2000
maxoms = 920
divider2 = 20
showmaxoms = 460
else
usescala = "50"
divider = 200
divider2 = 2
maxoms = 900
showmaxoms = 900
end if



	
	
	'***************************************************************************
	'***** Funktioner ****
	'***************************************************************************
	function valgtemedarbejdere()
	
	end function
	
	
	
	
	function omsBeregningprmd(counter)

    'Response.Write "intJobTpris/intBudgetTime" & intJobTpris &" / "& intBudgetTimer & "<br>"

		if cint(aty_fakbar) = 1 then 'aktivitet fakbar?
			
            if oRec("fastpris") = 9999 then '1: 1.9.2011 Deaktiveret, timepris følger altid medarbejertypen /:999 'fastpris eller lbn timer
			call kommaFunc()
                
                if intBudgetTimer <> 0 then
                tpThis = intJobTpris/intBudgetTimer
                else
                tpThis = intJobTpris
                end if
			
            mrdOms(counter) = mrdOms(counter) + ((tpThis) * oRec("timer"))
			mrdFakTimer(counter) = mrdFakTimer(counter) + oRec("timer")
			else
				'if aktNotFakbar = "y" then
				mrdOms(counter) = mrdOms(counter) + (oRec("Timer") * oRec("timepris") * (oRec("kurs")/100)) '*** Omregning til basisVal **'
				mrdFakTimer(counter) = mrdFakTimer(counter) + oRec("timer")
				'else
				'mrdOms(counter) = mrdOms(counter)
				'mrdNotFakTimer(counter) = mrdNotFakTimer(counter) + oRec("timer")
				'end if
			end if
		else
		mrdNotFakTimer(counter) = mrdNotFakTimer(counter) + oRec("timer")
		mrdOms(counter) = mrdOms(counter)
		end if
		
			if len(oRec("kostpris")) <> 0 then
			mrdKostpris(counter) = mrdKostpris(counter) + (oRec("Timer") * oRec("kostpris")) '** Altid basisVal **'
			else
			mrdKostpris(counter) = mrdKostpris(counter)
			end if
		
		dotwhere = instr(mrdOms(counter), ".")
		if dotwhere <> "0" then
			mrdOms(counter) = left(mrdOms(counter), (dotwhere-1))
		end if
		
		
		dotwhere2 = instr(mrdKostpris(counter), ".")
		if dotwhere2 <> "0" then
			mrdKostpris(counter) = left(mrdKostpris(counter), (dotwhere2-1))
		end if
		
		mrdTotTimer(counter) = mrdTotTimer(counter) + oRec("timer")
	end function
	
	
	
	
	
	
	
	'public function omsBprmdUM(counter)
	'	mrdOmsTUM(counter) = mrdOmsTUM(counter) + (oRec2("Timer") * oRec2("timepris"))
	'	
	'	dotwhere = instr(mrdOmsTUM(counter), ".")
	'	if dotwhere <> "0" then
	'		mrdOmsTUM(counter) = left(mrdOmsTUM(counter), (dotwhere-1))
	'	end if
	'end function
	
	
	
	public varBeregning
	public left_intMrdOms
	
	public function kommaFunc()
	
	if intBudgetTimer <> 0 then
	intBudgetTimer = intBudgetTimer
	else
	intBudgetTimer = 1
	end if
	
	varBeregning = (intJobTpris/intBudgetTimer)
	instr_intMrdOms = SQLBless(instr(varBeregning, "."))
	
		if instr_intMrdOms > 0 then
		left_intMrdOms = left(varBeregning, instr_intMrdOms + 2)
		else
		left_intMrdOms = varBeregning
		end if
	
	end function
	
	function SQLBless(s)
	dim tmp
	tmp = s
	tmp = replace(tmp, ",", ".")
	SQLBless = tmp
	end function
%>
