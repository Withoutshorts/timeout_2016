<%

				
				Redim preserve sonTimerVal(iRowLoop)
				Redim preserve manTimerVal(iRowLoop)
				Redim preserve tirTimerVal(iRowLoop)
				Redim preserve onsTimerVal(iRowLoop)
				Redim preserve torTimerVal(iRowLoop)
				Redim preserve freTimerVal(iRowLoop)
				Redim preserve lorTimerVal(iRowLoop)
				
				Redim preserve dtimeTtidspkt_son(iRowLoop)
				Redim preserve dtimeTtidspkt_man(iRowLoop)
				Redim preserve dtimeTtidspkt_tir(iRowLoop)
				Redim preserve dtimeTtidspkt_ons(iRowLoop)
				Redim preserve dtimeTtidspkt_tor(iRowLoop)
				Redim preserve dtimeTtidspkt_fre(iRowLoop)
				Redim preserve dtimeTtidspkt_lor(iRowLoop)
				
				Select case datepart("w", aTable1Values(12, iRowLoop))
				case 1
				sonTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				sonRLoop = iRowLoop
				dtimeTtidspkt_son(sonRLoop) = aTable1Values(14,sonRLoop)
				sonKomm = aTable1Values(15,iRowLoop)
				
					if aTable1Values(20,iRowLoop) = 1 then
					sonKomm_off = 1
					else
					sonKomm_off = 0
					end if
				
				case 2
				manTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				manRLoop = iRowLoop
				dtimeTtidspkt_man(manRLoop) = aTable1Values(14,manRLoop)
				manKomm = aTable1Values(15,iRowLoop)
				
					if aTable1Values(20,iRowLoop) = 1 then
					manKomm_off = 1
					else
					manKomm_off = 0
					end if
					
				case 3
				tirTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				tirRLoop = iRowLoop
				dtimeTtidspkt_tir(tirRLoop) = aTable1Values(14,tirRLoop)
				tirKomm = aTable1Values(15,iRowLoop)
					
					if aTable1Values(20,iRowLoop) = 1 then
					tirKomm_off = 1
					else
					tirKomm_off = 0
					end if
					
				case 4
				onsTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				onsRLoop = iRowLoop
				dtimeTtidspkt_ons(onsRLoop) = aTable1Values(14,onsRLoop)
				onsKomm = aTable1Values(15,iRowLoop)
				
					if aTable1Values(20,iRowLoop) = 1 then
					onsKomm_off = 1
					else
					onsKomm_off = 0
					end if
				
				case 5
				torTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				torRLoop = iRowLoop
				dtimeTtidspkt_tor(torRLoop) = aTable1Values(14,torRLoop)
				torKomm = aTable1Values(15,iRowLoop)
				
					if aTable1Values(20,iRowLoop) = 1 then
					torKomm_off = 1
					else
					torKomm_off = 0
					end if
					
				case 6
				freTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				freRLoop = iRowLoop
				dtimeTtidspkt_fre(freRLoop) = aTable1Values(14,freRLoop)
				freKomm = aTable1Values(15,iRowLoop)
				
					if aTable1Values(20,iRowLoop) = 1 then
					freKomm_off = 1
					else
					freKomm_off = 0
					end if
					
				case 7
				lorTimerVal(iRowLoop) = SQLBless(aTable1Values(10, iRowLoop))
				lorRLoop = iRowLoop
				dtimeTtidspkt_lor(lorRLoop) = aTable1Values(14,lorRLoop)
				lorKomm = aTable1Values(15,iRowLoop)
				
					if aTable1Values(20,iRowLoop) = 1 then
					lorKomm_off = 1
					else
					lorKomm_off = 0
					end if
					
				end select
				
	%>