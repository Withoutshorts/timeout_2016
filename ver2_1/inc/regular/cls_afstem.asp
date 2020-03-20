<%

'*** tjekker om uge er afsluttet / lukket / lønkørsel
public ugeerAfsl_og_autogk_smil
function tjkClosedPeriodCriteria(tjkDato, ugeNrAfsluttet, usePeriod, SmiWeekOrMonth, splithr, smilaktiv, autogk, autolukvdato, lonKorsel_lukketIO, ugegodkendt)

    ugeerAfsl_og_autogk_smil = 0

    'if session("mid") = 1 then
    'Response.write "<br>ugegodkendt: "& ugegodkendt &" varTjDatoUS_man: "& varTjDatoUS_man &" ugeNrAfsluttet: (uge=" &datepart("ww", ugeNrAfsluttet, 2, 2) &") " & cDate(ugeNrAfsluttet) & " > tjkDato: " & cDate(tjkDato) & " usePeriod: "& usePeriod & " SmiWeekOrMonth: " & SmiWeekOrMonth &""_
    '& "splithr: "& splithr &", smilaktiv: " & smilaktiv &", autogk: "& autogk &", autolukvdato: "& autolukvdato &", lonKorsel_lukketIO: "& lonKorsel_lukketIO 
    'end if

    'if ( (datepart("ww", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 0 AND splithr = 0) OR (cdate(ugeNrAfsluttet) >= cdate(tjkDato) AND cint(SmiWeekOrMonth) = 0 AND splithr = 1) _
    'OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 )) AND (cint(ugegodkendt) = 1 AND smilaktiv = 1 AND autogk = 1 AND ugeNrAfsluttet <> "1-1-2044") then

    'OR (smilaktiv = 1 AND autogk = 2 AND ugeNrAfsluttet <> "1-1-2044" AND splithr = 0)
    'AND autogk = 2

    call thisWeekNo53_fn(ugeNrAfsluttet) 

    if ( ( (thisWeekNo53 = usePeriod AND cint(SmiWeekOrMonth) = 0 AND splithr = 0) OR (cdate(ugeNrAfsluttet) >= cdate(tjkDato) AND cint(SmiWeekOrMonth) = 0 AND splithr = 1) _
    OR (datepart("m", ugeNrAfsluttet, 2, 2) = usePeriod AND cint(SmiWeekOrMonth) = 1 ))_
    AND ( ( (cint(ugegodkendt) = 1 AND autogk = 1) OR (autogk = 2) ) AND smilaktiv = 1 AND ugeNrAfsluttet <> "1-1-2044") ) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDato, 2, 2) = year(now) AND DatePart("m", tjkDato, 2, 2) < month(now)) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (day(now) > autolukvdatodato AND DatePart("yyyy", tjkDato, 2, 2) < year(now) AND DatePart("m", tjkDato, 2, 2) = 12)) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND DatePart("yyyy", tjkDato, 2, 2) < year(now) AND DatePart("m", tjkDato, 2, 2) <> 12) OR _
    (smilaktiv = 1 AND autolukvdato = 1 AND (year(now) - DatePart("yyyy", tjkDato, 2, 2) > 1))) OR _
    cint(lonKorsel_lukketIO) = 1 then

        ugeerAfsl_og_autogk_smil = 1
        
            'if session("mid") = 1 then
            'Response.write "<br>##ugeerAfsl_og_autogk_smil: "& ugeerAfsl_og_autogk_smil
            'end if

    else
        
        ugeerAfsl_og_autogk_smil = 0

            'if session("mid") = 1 then
            'Response.write "<br>#2#ugeerAfsl_og_autogk_smil: "& ugeerAfsl_og_autogk_smil
            'end if
    	
    end if

    if lto = "plan" AND (session("mid") = 268 OR session("mid") = 274) then  'Altid åben til fakturerbare timer for Clara og Ruben
        
            ugeerAfsl_og_autogk_smil = 0

    end if

    'if session("mid") = 1 then
    'Response.write "<br>##ugeerAfsl_og_autogk_smil: "& ugeerAfsl_og_autogk_smil & "godkendtstatus: "& godkendtstatus
    'end if
    
end function


public ferieBal
function feriebal_fn(ferieOptjtimer_X, ferieOptjOverforttimer_X, ferieOptjUlontimer_X, ferieAFTimer_X, ferieAFulonTimer_X, ferieUdbTimer_X)

     if len(trim(ferieOptjtimer_X)) <> 0 then
     ferieOptjtimer_X = ferieOptjtimer_X
     else
     ferieOptjtimer_X = 0
     end if

     if len(trim(ferieOptjOverforttimer_X)) <> 0 then
     ferieOptjOverforttimer_X = ferieOptjOverforttimer_X
     else
     ferieOptjOverforttimer_X = 0
     end if

     if len(trim(ferieOptjUlontimer_X)) <> 0 then
     ferieOptjUlontimer_X = ferieOptjUlontimer_X
     else
     ferieOptjUlontimer_X = 0
     end if

     if len(trim(ferieAFTimer_X)) <> 0 then
     ferieAFTimer_X = ferieAFTimer_X
     else
     ferieAFTimer_X = 0
     end if

     if len(trim(ferieAFulonTimer_X)) <> 0 then
     ferieAFulonTimer_X = ferieAFulonTimer_X
     else
     ferieAFulonTimer_X = 0
     end if

     if len(trim(ferieUdbTimer_X)) <> 0 then
     ferieUdbTimer_X = ferieUdbTimer_X
     else
     ferieUdbTimer_X = 0
     end if

     select case lto
     case "esn", "tec"
     ferieBal  = ((ferieOptjtimer_X + ferieOptjOverforttimer_X + ferieOptjUlontimer_X) - (ferieAFTimer_X + ferieAFulonTimer_X + ferieUdbTimer_X))
     case "cst", "wwf"
     ferieBal  = ((ferieOptjtimer_X + ferieOptjOverforttimer_X + ferieOptjUlontimer_X) - (ferieAFTimer_X + ferieAFulonTimer_X + ferieUdbTimer_X))
     case "xwwf"
	 ferieBal  = ((ferieOptjtimer_X + ferieOptjOverforttimer_X + ferieOptjUlontimer_X) - (ferieAFTimer_X + ferieUdbTimer_X - (ferieAFulonTimer_X)))
     case else
     ferieBal  = ((ferieOptjtimer_X + ferieOptjOverforttimer_X + ferieOptjUlontimer_X) - (ferieAFTimer_X + ferieUdbTimer_X))
     end select


end function



'********* Norm Real Graf ****************************
function normRealGrafWeekPage(medarbid, stdatoSQL)


    'Response.write "stdatoSQL "& stdatoSQL

    call normtimerPer(medarbid, stdatoSQL, 6, 0) %>
   
   <div style="position:relative; left:0px; top:5px; width:200px; border:0px #CCCCCC solid;">
   <table width=100% cellspacing=0 cellpadding=0 border=0>
   <tr>
        <td style="height:20px; border-bottom:1px #999999 solid;">Timer</td>
        <td style="border-bottom:1px #cccccc solid;"">
        <div style="position:absolute; left:30px; z-index:500000000; top:0px; width:20px; border-right:1px #cccccc solid;">
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr><td align=right style="padding:2px;" class=lille>2</td></tr>
        </table>
        </div>
        <div style="position:absolute; left:50px; z-index:500000000; top:0px; width:20px; border-right:1px #cccccc solid;">
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr><td align=right style="padding:2px;" class=lille>4</td></tr>
        </table>
        </div>
        <div style="position:absolute; left:70px; z-index:500000000; top:0px; width:34px; border-right:1px #cccccc solid;">
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr><td align=right style="padding:2px;" class=lille>7,4</td></tr>
        </table>
        </div>
        <div style="position:absolute; left:104px; z-index:500000000; top:0px; width:26px; border-right:1px #cccccc solid;">
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr><td align=right style="padding:2px;" class=lille>10</td></tr>
        </table>
        </div>
        <div style="position:absolute; left:130px; z-index:500000000; top:0px; width:20px; border-right:1px #cccccc solid;">
        <table cellpadding=0 cellspacing=0 border=0 width=100%>
        <tr><td align=right style="padding:2px;" class=lille>12</td></tr>
        </table>
        </div>
        </td>
   </tr>
  <%tdheight = 31 %>

  <%
  dtMan =  dateAdd("d", 0, stdatoSQL)
  dtTir =  dateAdd("d", 1, stdatoSQL)
  dtOns =  dateAdd("d", 2, stdatoSQL)
  dtTor =  dateAdd("d", 3, stdatoSQL)
  dtFre =  dateAdd("d", 4, stdatoSQL)
  dtLor =  dateAdd("d", 5, stdatoSQL)
  dtSon =  dateAdd("d", 6, stdatoSQL)
  %>

   <tr>
       
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid; " class=lille>Man <br /> <%=day(dtMan)&"."& month(dtMan) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, stdatoSQL, nTimMan, 30, 21) %>&nbsp;</td>
         </tr>
         <tr>
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid; " class=lille>Tir <br /> <%=day(dtTir)&"."& month(dtTir) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 1, stdatoSQL)) &"-"& month(dateadd("d", 1, stdatoSQL)) &"-"& day(dateadd("d", 1, stdatoSQL)), nTimTir, 30, 53) %>&nbsp;</td>
         </tr>
         <tr>
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Ons <br /> <%=day(dtOns)&"."& month(dtOns) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 2, stdatoSQL)) &"-"& month(dateadd("d", 2, stdatoSQL)) &"-"& day(dateadd("d", 2, stdatoSQL)), nTimOns, 30, 85) %>&nbsp;</td>
        </tr>
         <tr>
        <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Tor <br /> <%=day(dtTor)&"."& month(dtTor) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 3, stdatoSQL)) &"-"& month(dateadd("d", 3, stdatoSQL)) &"-"& day(dateadd("d", 3, stdatoSQL)), nTimTor, 30, 117) %>&nbsp;</td>
         </tr>
         <tr>
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Fre <br /> <%=day(dtFre)&"."& month(dtFre) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 4, stdatoSQL)) &"-"& month(dateadd("d", 4, stdatoSQL)) &"-"& day(dateadd("d", 4, stdatoSQL)), nTimFre, 30, 149) %>&nbsp;</td>
         </tr>
         <tr>
         <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Lør <br /> <%=day(dtLor)&"."& month(dtLor) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 5, stdatoSQL)) &"-"& month(dateadd("d", 5, stdatoSQL)) &"-"& day(dateadd("d", 5, stdatoSQL)), nTimLor, 30, 180) %>&nbsp;</td>
        </tr>
         <tr>
        <td style="height:<%=tdheight%>px; border-bottom:1px #D6Dff5 solid;" class=lille>Søn <br /> <%=day(dtSon)&"."& month(dtSon) %></td><td valign=bottom style="border-bottom:1px #EFF3FF solid;"><%call normrealgraf(medarbid, year(dateadd("d", 6, stdatoSQL)) &"-"& month(dateadd("d", 6, stdatoSQL)) &"-"& day(dateadd("d", 6, stdatoSQL)), nTimSon, 30, 211) %>&nbsp;</td>
        
        
   </tr>
</table>
   
   </div>

<%end function


'******************************************************









public ntimPer, ntimMan, ntimTir, ntimOns, ntimTor, ntimFre, ntimLor, ntimSon, antalDageMtimer, antalDageMtimerIgnHellig
public ntimManIgnHellig, ntimTirIgnHellig, ntimOnsIgnHellig, ntimTorIgnHellig, ntimFreIgnHellig, ntimLorIgnHellig, ntimSonIgnHellig, nTimerPerIgnHellig
	function normtimerPer(medid, stDato, interval, io)

    if medid <> "0" AND len(trim(medid)) <> 0 AND instr(medid, ",") = 0 then
    medid = trim(medid)
    else 
    medid = 0
    end if

    'Response.write "VAR: m: "& medid &" , stdato: "& stDato &", intervla: "& interval & "<br>"
	
	'Response.Write "Kalder normtidpper<br>"
	
	'** Maks 30 typer. **'
	dim mtyperIntvDato, mtyperIntvTyp, mtyperIntvEndDato, mtyperIntvRul 
	redim mtyperIntvDato(30), mtyperIntvTyp(30), mtyperIntvEndDato(30 mtyperIntvRul(30)
	
	'** Aktuel medarbejdertype **'
	strSQLtype = "SELECT medarbejdertype, ansatdato, opsagtdato FROM medarbejdere WHERE mid = "& medid & ""

    'response.Write strSQLtype
    'Response.Flush

	oRec2.open strSQLtype, oConn, 3
	if not oRec2.EOF then
	
	mtyperIntvTyp(0) = oRec2("medarbejdertype")
    mtyperIntvDato(0) = oRec2("ansatdato")
    mtyperIntvEndDato(0) = oRec2("opsagtdato")
	mtyperIntvRul(0) = 0

	end if
	oRec2.close
	
	
	'** Hvis medarbjeder ansat dato er før licens startdato benyttes licens startDato **'
	call licensStartDato()
	ltoStDato = startDatoDag &"/"& startDatoMd &"/"& startDatoAar
	if cdate(ltoStDato) > cDate(mtyperIntvDato(0)) then
	mtyperIntvDato(0) = ltoStDato
	end if
	
    
    'Response.Write mtyperIntvDato(0) &" "& ltoStDato 
	 
	'*** Finder medarbejdertyper_historik ***'
    strSQLmth = "SELECT mid, mtype, mtypedato, normtimer_rul FROM medarbejdertyper_historik WHERE "_
    &" mid = "& medid &" ORDER BY mtypedato, id"
    
    'Response.Write strSQLmth
    'Response.flush
    
    t = 1
    oRec2.open strSQLmth, oConn, 3
    while not oRec2.EOF 
    
    mtyperIntvTyp(t) = oRec2("mtype")
    mtyperIntvDato(t) = oRec2("mtypedato")
    mtyperIntvRul(t) = oRec2("normtimer_rul")
    
    t = t + 1
    oRec2.movenext
    wend
    oRec2.close
    
    'if session("mid") = 1 then
    'Response.Write "t: " & t & " interval = "& interval &"<br>"
    'Response.Write mtyperIntvTyp(t) & "##" & mtyperIntvDato(0) &" "& ltoStDato & "<br>"
    '*****************************************'
    'end if
	
			
			stDatoDag = weekday(stDato, 2)
			ntimPer = 0
            antalDageMtimer = 0
            nTimerPerIgnHellig = 0
            antalDageMtimerIgnHellig = 0
			
			for c = 0 to interval
			
			            if c = 0 then
			            n = stDatoDag
			            datoCount = stDato
			            else
			                if n < 7 then
			                n = n + 1
			                else
			                n = 1
			                end if
			            datoCount = dateAdd("d", 1, datoCount)
			            end if
			            

                        'if session("mid") = 1 then
			            'Response.Write "her: " & cDate(datoCount) &" >= "& cDate(mtyperIntvDato(0)) & "<br>"
                        'end if


			            '*** Tjekker ansatDato **'
			            '*** Skal først tjekke dage efter ansat dato, og før opsagt dato ***'
                        'IO = 2 skal ikke tjekke opsagdato. F.eks forecasdt op kapaciset, hvor medarbejder skal have foreste for de månder de er ansat, selvom de er sagt op.
			            'if ( cDate(datoCount) >= cDate(mtyperIntvDato(0)) ) AND (cDate(datoCount) < cDate(mtyperIntvEndDato(0)) ) then
	                    if ( cDate(datoCount) >= cDate(mtyperIntvDato(0)) AND cDate(datoCount) < cDate(mtyperIntvEndDato(0)) AND io <> 2) OR (cDate(datoCount) >= cDate(mtyperIntvDato(0)) AND io = 2) then		            

			            'if session("mid") =  1 then
			            'Response.Write "<u>cdate(datoCount) </u>"& cdate(datoCount) & " aftrædelse : "& cDate(mtyperIntvEndDato(0))  &"<br>"
	                    'end if		            

			            '*** Finder den medarbejder typer der passer til den valgte dato ***'
			            if t = 1 then
			            mtypeUse = mtyperIntvTyp(0) 
                        mtyperRulUse = 0
			            'Response.Write  " A mtypeUse: "&  mtypeUse & "<br>"
			            else
			               
			                for d = 2 to t  
			                    
			                    
			                    if cdate(datoCount) >= mtyperIntvDato(d-1) then 
    			                mtypeUse = mtyperIntvTyp(d-1)
                                mtyperRulUse = mtyperIntvRul(d-1)
    			                'Response.Write  cdate(datoCount) & " >= "& mtyperIntvDato(d-1)  &" - B mtypeUse: "&  mtypeUse & "<br>"
    			                else
    			                    if d = 2 then
    			                    mtypeUse = mtyperIntvTyp(d-1)
                                    mtyperRulUse = mtyperIntvRul(d-1)
    			                    'Response.Write  cdate(datoCount) & " - C1 mtypeUse: "&  mtypeUse & "<br>"
			                    
    			                    else
			                        mtypeUse = mtypeUse
			                        mtyperRulUse = mtyperRulUse
                                    'Response.Write  cdate(datoCount) & " - C mtypeUse: "&  mtypeUse & "<br>"
			                    
			                        end if
			                    end if
			            
			                next
			            
			            end if




                       'Er der oprettet Rul på medarbejdertypen
                        select case lto
                        case "intrant - local", "outz", "kongeaa" 'kun aktiveret for
    
                        if mtyperRulUse = 1 then
			            
                        end if
                        
                        end select
			           
			
			            select case n 
						case 7
						normdag = "normtimer_son"
						nd1 = 0
						case 1
						normdag = "normtimer_man"
						nd2 = 0
						case 2
						normdag = "normtimer_tir"
						nd3 = 0
						case 3
						normdag = "normtimer_ons"
						nd4 = 0
						case 4
						normdag = "normtimer_tor"
						nd5 = 0
						case 5
						normdag = "normtimer_fre"
						nd6 = 0
						case 6
						normdag = "normtimer_lor"
						nd7 = 0
						end select
						
					
					'** Finder normtimer på den valgte dag for den vagte type **'	
                    if len(trim(mtypeUse)) <> 0 then
                    mtypeUse = mtypeUse
                    else
                    mtypeUse = 0
                    end if

					strSQLnt = "SELECT "& normdag &" AS timer FROM "_
					&" medarbejdertyper t WHERE t.id = "& mtypeUse 
				
					
					'if session("mid") = 1 then
					'Response.write strSQLnt & " // n:"& n & "<hr>"
					'Response.flush
					'end if 
					
					oRec3.open strSQLnt, oConn, 3 
					if not oRec3.EOF then
						
						ntimPerThis = oRec3("timer")
						
						select case n 
						case 7
						ntimSon = ntimPerThis
                        ntimSonIgnHellig = ntimSon
						case 1
						ntimMan = ntimPerThis
                        ntimManIgnHellig = ntimMan
						case 2
						ntimTir = ntimPerThis
                        ntimTirIgnHellig = ntimTir
						case 3
						ntimOns = ntimPerThis
						ntimOnsIgnHellig = ntimOns
                        case 4
						ntimTor = ntimPerThis
						ntimTorIgnHellig = ntimTor
                        case 5
						ntimFre = ntimPerThis
						ntimFreIgnHellig = ntimFre
                        case 6
						ntimLor = ntimPerThis
						ntimLorIgnHellig = ntimLor
                        end select
						
                        'Response.Write "ntimPerThis:" & ntimPerThis & "<br>"
						
					end if
					oRec3.close 
					
					'Response.write "datoCount " & formatdatetime(datoCount, 2) & "<br>"
					'Response.end

                    
          
                    '**Ignorer helligdage
                    nTimerPerIgnHellig = nTimerPerIgnHellig + ntimPerThis	


                    '** Antal dage ign. helligdage
                    if ntimPerThis <> 0 then
                        antalDageMtimerIgnHellig = antalDageMtimerIgnHellig + 1
                    'Response.Write "antalDageMtimerIgnHellig:" & antalDageMtimerIgnHellig &"<br>"
                    end if 

					
				    call helligdage(datoCount, 0, lto, medid)
				        
				        if cint(erHellig) = 1 then
				            select case n 
						    case 7
						    ntimSon = 0
						    case 1
						    ntimMan = 0
						    case 2
						    ntimTir = 0
						    case 3
						    ntimOns = 0
						    case 4
						    ntimTor = 0
						    case 5
						    ntimFre = 0
						    case 6
						    ntimLor = 0
						    end select

                            ntimPerThis = 0

						end if
					
					if erHellig <> 1 then
					ntimPer = ntimPer + ntimPerThis	
                            
                            if ntimPerThis <> 0 then
                            antalDageMtimer = antalDageMtimer + 1
                            'Response.Write "antalDageMtimer:" & antalDageMtimer & " ntimPerThis: "& ntimPerThis &"<br>"
                            end if 

					else
					ntimPer = ntimPer

                             

					end if
					
				
					
					
					
					end if '** Ansat Dato **'
					
							
		    next
			
			if len(ntimPer) <> 0 then
			ntimPer = ntimPer
			else
			ntimPer = 0
			end if

            if len(nTimerPerIgnHellig) <> 0 then
			nTimerPerIgnHellig = nTimerPerIgnHellig
			else
			nTimerPerIgnHellig = 0
			end if

            
            if antalDageMtimer <> 0 then
            antalDageMtimer = antalDageMtimer
            else
            antalDageMtimer = 1
            end if

            if antalDageMtimerIgnHellig <> 0 then
            antalDageMtimerIgnHellig = antalDageMtimerIgnHellig
            else
            antalDageMtimerIgnHellig = 1
            end if

        
			
			'if session("mid") = 1 then
            'Response.write "her ntimPer: " & ntimPer & "<br>"
            'end if
            'Response.end

            if io = 1 then 'reduktion for ferie afholdt, ulon og planlagt

            stDatoSQLkri = year(stDato) &"-"& month(stDato) &"-"& day(stDato) 
            slDatoSQLkri = dateAdd("d", interval, stDatoSQLkri)
            slDatoSQLkri = year(slDatoSQLkri) &"-"& month(slDatoSQLkri) &"-"& day(slDatoSQLkri) 

             strSQLfeau = "SELECT sum(timer) AS sumtimer, tdato, month(tdato) AS month, tfaktim FROM timer WHERE tmnr = " & medid  & ""_
	         &" AND (tfaktim = 11 OR tfaktim = 13 OR tfaktim = 14 OR tfaktim = 18 OR tfaktim = 19 OR tfaktim = 22 OR tfaktim = 23 OR tfaktim = 24 OR tfaktim = 25 OR tfaktim = 115) "_
             &" AND tdato BETWEEN '"& stDatoSQLkri &"' AND '"& slDatoSQLkri &"' GROUP BY tmnr ORDER BY tdato"

            'response.write strSQLfeau & "<br>"

            oRec6.open strSQLfeau, oConn, 3
	        if not oRec6.EOF then
        
                'response.write "Sumtimer: "& oRec6("sumtimer") &"<br>"

                ntimPer = ntimPer - oRec6("sumtimer")
          
            oRec6.movenext
            end if
            oRec6.close          

            end if        


	end function


    public holidayTimer
    function holidayHoursInPer(medarbtype, stPer, slPer, medarbid)
        
        stPer = year(stPer) &"/"& month(stPer) &"/"& day(stPer)
        slPer = year(slPer) &"/"& month(slPer) &"/"& day(slPer)

        strSQLMedCal = "SELECT med_cal FROM medarbejdere WHERE mid = "& medarbid
        oRec7.open strSQLMedCal, oConn, 3
        if not oRec7.EOF then
            med_cal = oRec7("med_cal")
        end if
        oRec7.close


        call medariprogrpFn(medarbid)
        medariprogrpTxtDage = medariprogrpTxt

        holidayTimer = 0
        strSQLNormTimer = "SELECT normtimer_man, normtimer_tir, normtimer_ons, normtimer_tor, normtimer_fre, normtimer_lor, normtimer_son FROM medarbejdertyper WHERE id = "& medarbtype
        'response.Write strSQLNormTimer            
        oRec7.open strSQLNormTimer, oConn, 3
                if not oRec7.EOF then

                    nh_projgrp = 0
                    strSQLHoli = "SELECT nh_date, nh_projgrp, nh_name FROM national_holidays WHERE nh_open = 1 AND nh_date BETWEEN '"& stPer &"' AND '"& slPer &"' AND nh_country = '"& med_cal &"'"
                    'response.Write "<br> holi sql " & strSQLHoli
                    oRec8.open strSQLHoli, oConn, 3
                    while not oRec8.EOF
                        holiday = datepart("w", oRec8("nh_date"), 2,2)

                 

                        '*** PROGRP

                            prgGrpFundet = 0
                            negativTjk = 0
                            positivTjk = 0

                            if oRec8("nh_projgrp") <> "" then 'len(trim(oRec8("nh_projgrp"))) <> 0 AND 
                            nh_projgrp = 1
                            nh_projgrp_arr = split(oRec8("nh_projgrp"), ",")

                            else
                            nh_projgrp = 0
                            end if


                            if cint(nh_projgrp) = 1 then


                                
                                 for p = 0 to UBOUND(nh_projgrp_arr) AND prgGrpFundet = 0

                                        if instr(trim(nh_projgrp_arr(p)), "-") = 0 then 'POSTIVE DVS kun helligdag for disse projektgrupper


                                            'Response.Write "medariprogrpTxtDage: "& medariprogrpTxtDage &" nh_projgrp_arr(p): " & nh_projgrp_arr(p) 

                                            positivTjk = 1
                                            
                                            if instr(medariprogrpTxtDage, "#"& trim(nh_projgrp_arr(p)) &"#") <> 0 then
                                            prgGrpFundet = 1
                                            end if


                                            'Response.Write " erHellig: "& erHellig & "<br>"

                                        else

                                            'NEGATIVE DVS ikke helligdag for disse projektgrupper
                                            nh_projgrp_arr(p) = replace(nh_projgrp_arr(p), "-", "")
                                            negativTjk = 1
                                            
            
                                            if instr(medariprogrpTxtDage, "#"& trim(nh_projgrp_arr(p)) &"#") <> 0 then
                                            prgGrpFundet = 1
                                            end if


                                        end if

                                 next

                        end if

                        '*** PROGRP END

                        '** Hvis der ikke er angivet gruppe, GÆLDER ALLE, nh_projgrp = 0
                        '** Hvis der er angivet KUN for disse grupper: nh_projgrp = 1 AND prgGrpFundet = 1 AND instr(trim(nh_projgrp_arr(p)), "-") = 0
                        '** Hvis der er angivet MINUS eksluder disse grupper (Gruppen må ikke findes så): nh_projgrp = 1 AND prgGrpFundet = 0 AND instr(trim(nh_projgrp_arr(p)), "-") <> 0)

                        if (nh_projgrp = 0) OR (nh_projgrp = 1 AND prgGrpFundet = 1 AND positivTjk = 1) OR (nh_projgrp = 1 AND prgGrpFundet = 0 AND negativTjk = 1) then

                            select case holiday
                            case 1
                                holidayTimer = holidayTimer + oRec7("normtimer_man")
                            case 2
                                holidayTimer = holidayTimer + oRec7("normtimer_tir")
                            case 3
                                holidayTimer = holidayTimer + oRec7("normtimer_ons")
                            case 4
                                holidayTimer = holidayTimer + oRec7("normtimer_tor")
                            case 5
                                holidayTimer = holidayTimer + oRec7("normtimer_fre")
                            case 6       
                                holidayTimer = holidayTimer + oRec7("normtimer_lor")
                            case 7
                                holidayTimer = holidayTimer + oRec7("normtimer_son")
                            end select

                        end if

                            'response.Write "<br> new holiday " & oRec8("nh_date") & " holitimer " & holidayTimer

                        'if session("mid") = 1 then
                        'response.Write "<br>" & oRec8("nh_name") & "  holiday timer:  "  & holidayTimer & " holiday number " & holiday
                        'end if

                        oRec8.movenext
                        wend
                        oRec8.close

                       'response.Write "<br> holidayshour i per " & holidayTimer 

            end if
            oRec7.close

    end function


    public totNormtimer 
    function work_days2(BegDate, EndDate, manTimer, tirTimer, onsTimer, torTimer, freTimer, lorTimer, sonTimer)

        daysInTotal = DateDiff("d", BegDate, EndDate)
        daysInTotal = daysInTotal + 2

        'if session("mid") = 1 then
        'response.Write "<br>BegDate = "& BegDate &", EndDate = "& EndDate &" daysInTotal " & daysInTotal 
        'response.Write daysInTotal 
        'end if
    
        manTimerTot = 0
        tirTimerTot = 0
        onsTimerTot = 0
        torTimerTot = 0
        freTimerTot = 0
        lorTimerTot = 0
        sonTimerTot = 0
        totNormtimer = 0

        dagInterval = 1
        tjkDate = BegDate
        while dagInterval < daysInTotal
           

            select case DatePart("w", tjkDate, 2,2)
            case 1
            manTimerTot = manTimerTot + manTimer
            case 2
            tirTimerTot = tirTimerTot + tirTimer
            case 3
            onsTimerTot = onsTimerTot + onsTimer
            case 4
            torTimerTot = torTimerTot + torTimer
            case 5
            freTimerTot = freTimerTot + freTimer
            case 6
            lorTimerTot = lorTimerTot + lorTimer
            case 7   
            sonTimerTot = sonTimerTot + sonTimer
            end select

            tjkDate = DateAdd("w", 1, tjkDate)
            dagInterval = dagInterval + 1

        wend

            

            totNormtimer = manTimerTot + tirTimerTot + onsTimerTot + torTimerTot + freTimerTot + lorTimerTot + sonTimerTot 
            
            'if session("mid") = 1 then
            'response.Write " <br> her tottimer " & totNormtimer
            'end if
            'response.Write "End function"

    end function

    Public norm_aarstotal, arrsferie, antalUger, norm_ugetotal, antalhelligdagetimer, maanederansat, totalMedarbNorm, totalMedarbFerie
    function kapcitetsnorm(medarbid, medarbAnsdato, medarbOpsdato, startperiode, slutperiode, interval)

        'interval koder - 1 = år, 2 = halvår, 3 = kvartal
        
        norm_aarstotal = 0
        arrsferie = 0
        antalUger = 0
        norm_ugetotal = 0
        antalhelligdagetimer = 0
        maanederansat = 0
        'totalMedarbNorm = 0
        totalMedarbFerie = 0
        
        'dagsdato = Day(now) &"-"& Month(now) &"-"& Year(now)
        'response.Write "<br> Ansatdato " & medarbAnsdato
        if cdate(medarbAnsdato) < cdate(startperiode) then
            asDato = year(startperiode)&"/"&month(startperiode)&"/"&day(startperiode)
        else
            asDato = year(medarbAnsdato)&"/"&month(medarbAnsdato)&"/"&day(medarbAnsdato)
            
            if cdate(asDato) > cdate(slutperiode) then
                mtD = -1
            end if

        end if

    
        if cdate(medarbOpsdato) > cdate(startperiode) AND cdate(medarbOpsdato) < cdate(slutperiode) then
            endDtKri = year(medarbOpsdato)&"/"&month(medarbOpsdato)&"/"&day(medarbOpsdato)
            endDtKri = DateAdd("d", -1, endDtKri) 'Fordi man ikke arbejder den dag man bliver sagt op.
            endDtKri = year(endDtKri)&"/"&month(endDtKri)&"/"&day(endDtKri)            
        else            
            endDtKri = year(slutperiode)&"/"&month(slutperiode)&"/"&day(slutperiode)
        end if
        
        if cdate(medarbOpsdato) < cdate(startperiode) then
                mtD = -1
        end if
      


        maanederansat = DateDiff("m",asDato,endDtKri,2,2) + 1
        dageansat = DateDiff("d",asDato,endDtKri,2,2) 
        'response.Write "måneder ansat " & maanederansat
        if maanederansat > 12 then
            maanederansat = 12
        end if

        findafslutendeType = 1
        norm_ugetotal = 0
        nTimerPerIgnHellig = 0
        asDatoSQL = year(asDato) & "/" & month(asDato) & "/" & day(asDato)
        forsteDatoskift = asDato '"1-1-"&aar
        forstiforsteArr = asDato '"1-1-"&aar
        sidstedagiArr = endDtKri '"31-12-"&aar
        slutdatePerSQL = year(endDtKri) & "/" & month(endDtKri) & "/" & day(endDtKri)
        lastDatoskift = forsteDatoskift 
        '*** Finder medarbejdertyper_historik ***'
        '*** Hvis der er skiftet mellem typer i løbet af året skal den regne hele året ud.
        strSQLmth = "SELECT mth.mid, mth.id, mth.mtype mtypeid, mth.mtypedato FROM medarbejdertyper_historik mth "_
        &" INNER JOIN (SELECT id, mtypedato FROM medarbejdertyper_historik ORDER BY id DESC) as mth2 ON mth.id = mth2.id WHERE "_
        &" mth.mid = "& medarbid &" AND mth.mtypedato BETWEEN '"& asDatoSQL &"' AND '"& slutdatePerSQL &"' GROUP BY mth.mtypedato ORDER BY mth.mtypedato DESC"
        'response.Write "<br> start slut dato " & asDatoSQL &"-"& slutdatePerSQL
        'response.Write "<br> strSQLmth" & strSQLmth
        mt = 0
        oRec6.open strSQLmth, oConn, 3
        while not oRec6.EOF

            if mt = 0 then
                forsteDatoskift = oRec6("mtypedato")
            end if

            call normtimerPer(medarbid, oRec6("mtypedato"), 6, 2)

            if Cdate(oRec6("mtypedato")) = Cdate(asDatoSQL) then
                findafslutendeType = 0
            end if
   
            if mt = 0 then
                
                call work_days2(oRec6("mtypedato"), sidstedagiArr, ntimManIgnHellig, ntimTirIgnHellig, ntimOnsIgnHellig, ntimTorIgnHellig, ntimFreIgnHellig, ntimLorIgnHellig, ntimSonIgnHellig)
                
                'if session("mid") = 1 then
                'response.Write " fra "& mt &" "& oRec6("mtypedato") &" "& sidstedagiArr & " nTimer " & totNormtimer & " tId " & oRec6("mtypeid") 
                'end if

                'sidstedagiArr = day(sidstedagiArr) & "/" & month(sidstedagiArr) & "/" & year(sidstedagiArr)

            
                intervalWeeksArr(mt) = (dateDiff("d", oRec6("mtypedato"), sidstedagiArr, 2,2) + 1) '1 for skæve uger ved skift mm.              
                call holidayHoursInPer(oRec6("mtypeid"), oRec6("mtypedato"), sidstedagiArr, medarbid)               
                antalhelligdagetimer = antalhelligdagetimer + holidayTimer

            else

                call work_days2(oRec6("mtypedato"), DateAdd("w", - 1, lastDatoskift), ntimManIgnHellig, ntimTirIgnHellig, ntimOnsIgnHellig, ntimTorIgnHellig, ntimFreIgnHellig, ntimLorIgnHellig, ntimSonIgnHellig)
                'response.Write "<br> fra "& oRec6("mtypedato") &" "& lastDatoskift & " nTimer " & totNormtimer & " tId " & oRec6("mtypeid")
                intervalWeeksArr(mt) = (dateDiff("d", oRec6("mtypedato"), lastDatoskift, 2,2) + 1)
                call holidayHoursInPer(oRec6("mtypeid"), oRec6("mtypedato"), lastDatoskift, medarbid)               
                antalhelligdagetimer = antalhelligdagetimer + holidayTimer

            end if

            'norm_ugetotalArr(mt) = (nTimerPerIgnHellig / 5) * intervalWeeksArr(mt)
            norm_ugetotalArr(mt) = totNormtimer
            'response.Write "<br> norm_ugetotalArr " & norm_ugetotalArr(mt)
            'intervalWMnavn(mt) = oRec6("mtypenavn") & " " & oRec6("mid")

            lastDatoskift = year(oRec6("mtypedato")) &"/"& month(oRec6("mtypedato"))&"/"&day(oRec6("mtypedato"))
            mt = mt + 1
            oRec6.movenext
            wend
            oRec6.close


            '** sidste type skal altid gå til 31-12
            'intervalWeeksArr(mt) = dateDiff("w", lastDatoskift, sidstedagiArr, 2,2)

            '*** Finder den afsluttende typ før dette år***'
            if mt <> 0 AND findafslutendeType = 1 then
                'strSQLmth = "SELECT mth.mid, mth.mtype as mtypeid, mth.mtypedato, mt.type AS mtypenavn FROM medarbejdertyper_historik mth "_
                '&" LEFT JOIN medarbejdertyper mt ON (mt.id = mth.mtype) WHERE "_
                '&" mth.mid = "& medarbid &" AND mth.mtypedato < '"& asDatoSQL &"' ORDER BY mth.mtypedato DESC, mth.id DESC LIMIT 1"

                 strSQLmth = "SELECT mth.mid, mth.id, mth.mtype mtypeid, mth.mtypedato FROM medarbejdertyper_historik mth "_
                 &" INNER JOIN (SELECT id, mtypedato FROM medarbejdertyper_historik ORDER BY id DESC) as mth2 ON mth.id = mth2.id WHERE "_
                 &" mth.mid = "& medarbid &" AND mth.mtypedato < '"& asDatoSQL &"' GROUP BY mth.mtypedato ORDER BY mth.mtypedato DESC"
       

                oRec6.open strSQLmth, oConn, 3
                if not oRec6.EOF then

                    call normtimerPer(medarbid, forstiforsteArr, 6, 2)
                    
                    call work_days2(forstiforsteArr, DateAdd("w", - 1, lastDatoskift), ntimManIgnHellig, ntimTirIgnHellig, ntimOnsIgnHellig, ntimTorIgnHellig, ntimFreIgnHellig, ntimLorIgnHellig, ntimSonIgnHellig)
                    
                    intervalWeeksArr(mt) = (dateDiff("d", forstiforsteArr, lastDatoskift, 2,2) + 1)
                    norm_ugetotalArr(mt) = totNormtimer
                    'intervalWMnavn(mt) = oRec6("mtypenavn") & " " & oRec6("mid")       
                    
                    call holidayHoursInPer(oRec6("mtypeid"), forstiforsteArr, lastDatoskift, medarbid)               
                    antalhelligdagetimer = antalhelligdagetimer + holidayTimer

                mt = mt + 1
                end if
                oRec6.close
            end if


            antalUgerTot = 0


            if mtD = -1 OR maanederansat <= 0 then 'ansat efter det år der tjekkes
                                
                norm_aarstotal = 0
                arrsferie = 0
                antalUger = 0
                norm_ugetotal = 0
                antalhelligdagetimer = 0
                maanederansat = 0
                
            else

                'antalUger = dateDiff("w", forstiforsteArr, sidstedagiArr, 2,2)
                antalUger = (dateDiff("d", forstiforsteArr, sidstedagiArr, 2,2) + 1)          
                'response.Write "<Br> dage i per " & antalUger
                'antalhelligdagetimer = 0 '60

                if mt = 0 then 'samme type hele året
                    
                    call normtimerPer(medarbid, asDato, 6, 2) 'io = 2 ignorer opsagdato, da den benytter 1 dag i interval (ansat dato + 6 dage). Derfor skal medarbjedere der er opdagt have norm for m til de bliver opsagt. 
                    norm_ugetotal = (nTimerPerIgnHellig) '/ 5
                  
                    call work_days2(asDato, endDtKri, ntimManIgnHellig, ntimTirIgnHellig, ntimOnsIgnHellig, ntimTorIgnHellig, ntimFreIgnHellig, ntimLorIgnHellig, ntimSonIgnHellig)
                    norm_i_periode = totNormtimer

                    'ugeriperode = (dateDiff("d", forstiforsteArr, sidstedagiArr, 2,2) + 1) / 7
                    
                    call thisWeekNo53_fn(startperiode) 
                    startdatodag = thisWeekNo53 'datepart("ww", startperiode, 2,2)
                    call thisWeekNo53_fn(slutperiode) 
                    slutdatodag = thisWeekNo53 'datepart("ww", slutperiode, 2,2)
                    
                    
                    findestypeForAsdato = 0
                    strSQLmedarbType = "SELECT mtype FROM medarbejdertyper_historik WHERE mid = "& medarbid &" AND mtypedato < '"& asDatoSQL &"' ORDER BY mtypedato DESC, id DESC LIMIT 1"
                    'response.Write "herher " & strSQLmedarbType 
                    oRec6.open strSQLmedarbType, oConn, 3
                    if not oRec6.EOF then
                    findestypeForAsdato = 1
                    medarbejdertype = oRec6("mtype")
                    end if
                    oRec6.close

                    if findestypeForAsdato = 0 then
                        'response.Write "<br> medarb " & medarbid
                        strSQLmedarbType = "SELECT Medarbejdertype FROM medarbejdere WHERE mid = "& medarbid
                        oRec6.open strSQLmedarbType, oConn, 3
                        if not oRec6.EOF then
                            medarbejdertype = oRec6("Medarbejdertype")
                        end if
                        oRec6.close

                    end if

                    call holidayHoursInPer(medarbejdertype, asDatoSQL, slutdatePerSQL, medarbid)               
                    antalhelligdagetimer = antalhelligdagetimer + holidayTimer
            
                    'ntimMan + ntimTir + ntimOns + ntimTor + ntimFre + ntimLor + ntimSon  
                    'response.Write "<br> " & norm_ugetotal & " normer " & ntimMan &" "& ntimTir &" "& ntimOns &" "& ntimTor &" "& ntimFre &" "& ntimLor &" "& ntimSon
                else
                    'Skal vægtes
                    for mtarr = 0 to mt - 1 'UBOUND(norm_ugetotalArr)
                    norm_ugetotal = norm_ugetotal + norm_ugetotalArr(mtarr)
                    norm_i_periode = norm_i_periode + norm_ugetotalArr(mtarr)
                    antalUgerTot = antalUgerTot + intervalWeeksArr(mtarr)
                 
                    next

                    if antalUgerTot <> 0 then
                        'response.Write " 12345 nedned antalUgerTot " & antalUgerTot
                        norm_ugetotal = (norm_ugetotal / antalUgerTot) * 7 'antalUger
                        
                    else
                        norm_ugetotal = 0
                    end if

                               

                end if

                
                'response.Write "<br> workdays " & workdays
                'response.Write "<br> holidays " & antalhelligdagetimer
                if mt = 0 then 'samme type hele året
                'norm_aarstotal = (norm_ugetotal * antalUger) - (antalhelligdagetimer) 'maanederansat/12) *
                'norm_aarstotal = (norm_ugetotal * workdays) - (antalhelligdagetimer)
                else
                'norm_aarstotal = (norm_ugetotal * antalUgerTot) - (antalhelligdagetimer)
                'norm_aarstotal = (norm_ugetotal * workdays) - (antalhelligdagetimer)
                end if
                
                arrsferie = ((maanederansat/12) * norm_ugetotal * 6)
                
                norm_aarstotal = norm_i_periode - antalhelligdagetimer

                'if session("mid") = 1 then
                'response.Write "<br><br> totNormtimer " & norm_aarstotal
                'Response.Write "<br>norm_i_periode: " & norm_i_periode 
                'response.Write "<br> antalugertot " & antalUgerTot
                'response.Write "<br> norm_ugetotal " & norm_ugetotal
                'response.write "<br> arrsferie: " & arrsferie 
                'response.write "<br> antalhelligdagetimer: "& antalhelligdagetimer 
                'response.write "<br>" &  startperiode &","& slutperiode
                'end if

            end if 'mt -1


            if totalAntalMedarbs = 1 then
            totalMedarbNorm = norm_aarstotal
            totalMedarbFerie = arrsferie
            else 
            totalMedarbNorm = totalMedarbNorm + norm_aarstotal
            totalMedarbFerie = totalMedarbFerie + arrsferie  
            end if



    end function







public normtimerStDag
	function nortimerStandardDag(mtypeUse)
	
	
	ntimStdag = 0
	workdays_divider = 0
	
	    for n = 1 to 7 
	                    select case n 
						case 7
						normdag = "normtimer_son"
						case 1
						normdag = "normtimer_man"
						case 2
						normdag = "normtimer_tir"
						case 3
						normdag = "normtimer_ons"
						case 4
						normdag = "normtimer_tor"
						case 5
						normdag = "normtimer_fre"
						case 6
						normdag = "normtimer_lor"
						end select
						
					
					'** Finder normtimer på den valgte dag for den vagte type **'	
					strSQLnt = "SELECT "& normdag &" AS timer FROM "_
					&" medarbejdertyper t WHERE t.id = "& mtypeUse 
				    
		            oRec3.open strSQLnt, oConn, 3 
					if not oRec3.EOF then
						
						if oRec3("timer") <> 0 then
						ntimStdag = ntimStdag + oRec3("timer")
						workdays_divider = workdays_divider + 1
						else
						ntimStdag = ntimStdag
						end if
						
					end if
					oRec3.close 
				
				next
	            
	            '** workdays_divider bør altid være = 5
                '** Da danske overenskomster altdi er baseret på en 5 dages arb. uge
	            if workdays_divider <> 0 then
	            workdays_divider = workdays_divider
	            else
	            workdays_divider = 1
	            end if
	            
	            normtimerStDag = ntimStdag/workdays_divider
	
	end function







    
    





'**********************************************************    

dim thisWdt
redim thisWdt(7)
function normrealgraf(usemrn, realDato, nTimDag, lft, tp)



'************** Norm timer ***************
if nTimDag > 0 then

nTimDagPX = formatnumber(nTimDag, 0)
%>
   
   <span id=nt style="position:absolute; left:<%=lft%>px; top:<%=tp+1%>px; width:<%=10*nTimDagPX%>px; z-index:500; height:13px; background-color:#CCCCCC;">
   <table width=100% cellpadding=0 cellspacing=0 border=0><tr>
   <td align=right style="padding-right:5px;" class=lille><%
    if nTimDag <> 0 then
    nTimDag = nTimDag
    else
	nTimDag = 0
	end if
   
    %>
  
   <%=nTimDag %>
   
   </td></tr></table></span>
<%end if 
    
    
    
    
    '**** Real timer ****************
        strSQL = "SELECT sum(timer) AS timer_indtastet FROM timer WHERE Tmnr = "& usemrn &" AND Tdato ='"& realDato &"' AND ("& aty_sql_realhours &")"
		intDayHours = 0
		
		'Response.Write strSQL
		'Response.flush
		oRec.open strSQL, oConn, 3
		
		
		
		if not oRec.EOF then
		intDayHours = oRec("timer_indtastet")
	    end if
	    oRec.close 
	    
	    if intDayHours <> 0 then
	    intDayHours = intDayHours
	    else
	    intDayHours = 0
	    end if
	    
	    %>
   
   <%if intDayHours > 0 then 

   if intDayHours > 18 then
   intDayHoursDivval = 18
   else
   intDayHoursDivval = intDayHours
   end if

    
    thisWeekDay = datePart("w", realDato, 2,2)

   thisWdt(thisWeekDay) = 10*formatnumber(intDayHoursDivval,0)
   %>
   <span id="rt" style="position:absolute; padding:0px 0px 0px 2px; left:<%=lft%>px; top:<%=tp+15%>px; width:<%=thisWdt(thisWeekDay)%>px; z-index:1000; height:14px; background-color:yellowgreen;">
   <table callpadding=0 cellspacing=0 border=0 width=100%></tr><td class=lille style="padding-right:5px;" align=right><%=formatnumber(intDayHours, 2) %></td></table>
    
   </span>
   <%end if





    '*********** Timer_temp_timer ********'


        strSQL = "SELECT sum(timer) AS timer_temp FROM timer_import_temp WHERE medarbejderid = "& usemrn &" AND Tdato ='"& realDato &"' AND overfort = 0"
		intDayHours = 0
		
		'Response.Write strSQL
		'Response.flush
		oRec.open strSQL, oConn, 3
		
		
		
		if not oRec.EOF then
		intDayHours = oRec("timer_temp")
	    end if
	    oRec.close 
	    
	    if intDayHours <> 0 then
	    intDayHours = intDayHours
	    else
	    intDayHours = 0
	    end if
	    
	    %>
   
   <%if intDayHours > 0 then 

   if intDayHours > 18 then
   intDayHoursDivval = 18
   else
   intDayHoursDivval = intDayHours
   end if

    thisWeekDay = datePart("w", realDato, 2,2)
    thisLeft = thisWdt(thisWeekDay) + lft
   %>
   <span id="tt" style="position:absolute; padding:0px 0px 0px 2px; left:<%=thisLeft%>px; top:<%=tp+15%>px; width:<%=10*formatnumber(intDayHoursDivval,0)%>px; z-index:1000; height:14px; background-color:#FFC0CB;">
   <table callpadding=0 cellspacing=0 border=0 width=100%></tr><td class=lille style="padding-right:5px;" align=right><%=formatnumber(intDayHours, 2) %></td></table>
    
   </span>
   <%end if %>




   
 <%  
 end function 















public aktidKorrFundet
function indlasKorrektioner(mid, korrKom, korrReal, dtInd)

                '**** Ændrer status på planlagt ferie til afholdt ferie ***'
		        '*** (indlæser ny registrering så historik beholdes) ***'
		        '**** Tjekker om der findes reg. i forvejen *****'
		        '113 Korrek. Kom
		        '114 Korr. Real
		        
		       
		        for f = 1 to 2
		        
		        if f = 1 then
		        
                aktype = 113
                timerThis = split(korrKom, ":")
		        
                
                for t = 0 To UBOUND(timerThis)

                if t = 0 then
                timerThisSQL = timerThis(t)
                else
                timerThisSQL = timerThisSQL &","& formatnumber((timerThis(t) * 100) / 60, 0)
                end if

                next


                else
		        aktype = 114
		        timerThis = korrReal
                timerThisSQL = timerThis
                end if
		        
		        aktid = 0
		        editDato = year(now)&"/"& month(now)&"/"&day(now)
                
                call meStamdata(mid)

		        
		        '** Finder navn og id på Korrektions akt. ***'
		        strSQLKoafn = "SELECT a.id, a.navn, k.kkundenavn, k.kkundenr, j.jobnr, j.jobnavn FROM job j"_
		        &" LEFT JOIN aktiviteter a ON (a.fakturerbar = "& aktype &" AND a.job = j.id) "_
		        &" LEFT JOIN kunder k ON (k.kid = j.jobknr) "_
                &" WHERE j.jobstatus = 1 AND a.id <> 'NULL' GROUP BY a.id ORDER BY a.id DESC"
		        
		       'Response.write strSQLKoafn & "<br>"
               'Response.flush
		        
		        oRec3.open strSQLKoafn, oCOnn, 3
		        if not oRec3.EOF then
		        
		        if oRec3("id") <> "" then
		        aktid = oRec3("id")
		        aktnavn = oRec3("navn")

                jobnavn = oRec3("jobnavn")
                jobnr = oRec3("jobnr")

                knavn = oRec3("kkundenavn")
                knr = oRec3("kkundenr")

		        end if
		        
		        end if
		        oRec3.close
		        
		       if cdbl(aktid) <> 0 then
		        
		        
                       
                        tpThis = 0
                        kpThis = 0
                        kursThis = 100
                        kommThis = "Korrektion indlæst"
                        origin = 5
                        ekstTimer = 0

                        '*** Hent først evt. manuel indlæsning på dagen '**
                        strSQLfindesTimer = "SELECT timer FROM timer WHERE tdato = '"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"' AND tmnr = "& mid & " AND tfaktim = "& aktype & " AND taktivitetid = " & aktid
                        
                        'Response.write strSQLfindesTimer & "<br>"
                        'Response.flush
                        oRec4.open strSQLfindesTimer, oConn, 3
                        
                        while not oRec4.EOF 
                        ekstTimer = ekstTimer + (oRec4("timer"))
                        oRec4.movenext
                        wend
                        oRec4.close 

                        '** Tømmer ***
                        strSQLDelTimer = "DELETE FROM timer WHERE tdato = '"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"' AND tmnr = "& mid & " AND tfaktim = "& aktype & " AND taktivitetid = " & aktid
                        oConn.execute(strSQLDelTimer)

                        timerThisSQL = (timerThisSQL/1 + (ekstTimer/1))
                        timerThisSQL = replace(timerThisSQL, ",", ".")

                        'Response.write "(timerThisSQL/1 + ekstTimer/1) MID: "& mid &": " & timerThisSQL &"+"& ekstTimer & "<br>"


                        '*** Indlæser afholdt ferie ***'
		                strSQLKoins = "INSERT INTO timer "_
		                &"("_
		                &" timer, tfaktim, tdato, tmnavn, tmnr, tjobnavn, tjobnr, tknavn, tknr, "_
		                &" timerkom, TAktivitetId, taktivitetnavn, Taar, TimePris, TasteDato, fastpris, tidspunkt, "_
		                &" editor, kostpris, seraft, sttid, sltid, valuta, kurs, origin "_
		                &") "_
		                &" VALUES "_
		                &" (" _
		                & timerThisSQL &", "& aktype &", "_
		                & "'"& year(dtInd) &"/"& month(dtInd) & "/"& day(dtInd) &"', "_
		                & "'"& meNavn &"', "_
		                & mid &", "_
		                & "'"& jobnavn &"', "_
		                & "'"& jobnr &"', "_
		                & "'"& knavn &"', "_
		                & knr &", "_
		                & "'"& kommThis &"', "_
		                & aktid &", "_
		                & "'"& aktnavn &"', "_
		                & year(now) &", "_
		                & tpThis &", "_
		                & "'"& year(now) &"/"& month(now) & "/"& day(now) &"', 0, "_
		                & "'00:00:01', "_
		                & "'Korrektions indlæsning', "_
		                & kpThis &", 0, '00:00:00', '00:00:00', 1, "_
		                & kursThis &", "& origin &")"
		                
                       
		                
		                
		                
                        'Response.write strSQLKoins & "<br>"
                        'Response.flush

                        oConn.execute(strSQLKoins)
		            
                
               

               
                
                end if '*** aktid <> 0
                

                'response.write "<br>aktid: " & aktid 

                next
                
                if f = 1 then

                    if session("stempelur") <> 0 then 'er stempelur slået til og sder der tjekkes herfor
                    aktidKorrFundet = aktid
                    else
                    aktidKorrFundet = -1 
                    end if

                else

                aktidKorrFundet = aktid
                end if
                
                
                'Response.flush

end function



public lonKorsel_lukketPerDt, lonKorsel_lukketIO
function lonKorsel_lukketPer(tjkDay, hr_jobRisiko, thisMid)

     if isNull(hr_jobRisiko) <> true then
     hr_jobRisiko = hr_jobRisiko
     else
     hr_jobRisiko = 0
     end if

     hr_jobRisikoOpr = hr_jobRisiko

    'hr: job risiko -2 eller Stempelur altid -2
    lonKorsel_lukketPerDt = "1-1-2002"
    lk_close_projects = 0   
    lk_medarbtype = 0
    strSQL5 = "SELECT lk_dato, lk_close_projects, lk_medarbtype FROM lon_korsel WHERE lk_id <> 0 ORDER BY lk_dato DESC"
	oRec5.open strSQL5, oConn, 3 
	if not oRec5.EOF then
	    lonKorsel_lukketPerDt = oRec5("lk_dato")
        lk_close_projects = oRec5("lk_close_projects") 
        lk_medarbtype = oRec5("lk_medarbtype")
	end if
	oRec5.close 

     lk_medarbtypeLukket = 0
     if lk_medarbtype <> 0 then

         call meStamdata(thisMid)
         if cint(lk_medarbtype) = cint(meType) then
         lk_medarbtypeLukket = 1
         end if

     end if

     '*** FORCE LUK også ALLE projekter (valgt på HR listen afslutning af periode)
     select case cint(lk_close_projects) 
     case 1 'LUK også for projekter
     hr_setting = -2
     case 2 'LUK KUN projekter

        if hr_jobRisiko >= 0 then '20191011 OR hr_jobRisiko = -2 
        hr_setting = -2
        else
        hr_setting = 0
        end if
     
     case else '0
     hr_setting = hr_jobRisiko 'LUK KUN interne HR job -2
     end select

     tjkDayFormat = day(tjkDay) &"-"& month(tjkDay) &"-"& year(tjkDay)

        if cdate(lonKorsel_lukketPerDt) >= cDate(tjkDayFormat) AND cint(hr_setting) = -2 then

            if cint(lk_medarbtype) = 0 OR (cint(lk_medarbtype) <> 0 AND cint(lk_medarbtypeLukket) = 1) then
            lonKorsel_lukketIO = 1
            else
            lonKorsel_lukketIO = 0
            end if
        
     
        else
        lonKorsel_lukketIO = 0
        end if

        'if session("mid") = 1 then
        'Response.Write "lonKorsel_lukketPerDt "& cdate(lonKorsel_lukketPerDt) &" tjkDay " & cDate(tjkDay) & " cDate(tjkDayFormat): "& cDate(tjkDayFormat) &" lonKorsel_lukketIO: "& lonKorsel_lukketIO &" lk_medarbtype: "& lk_medarbtype &" lk_medarbtypeLukket: "& lk_medarbtypeLukket &" hr_jobRisiko: "& hr_jobRisiko &", hr_jobRisikoOpr = "& hr_jobRisikoOpr &" thisMid: "& thisMid &" lk_close_projects: "& lk_close_projects &" hr_setting: "& hr_setting &"<br>"
        'end if

end function


function lonKorsel_lukketPerPrDato(tjkDato, thisMid)

    
    lonKorsel_lukketPerDt = licensstdato

    tjkDatoThisSQL = year(tjkDato) &"-"& month(tjkDato) &"-"& day(tjkDato)
    call meStamdata(thisMid)

    strSQL5 = "SELECT lk_dato, lk_medarbtype FROM lon_korsel WHERE lk_id <> 0 AND ((lk_dato <= '"& tjkDatoThisSQL &"' AND (lk_medarbtype = '0') "_
    &" OR (lk_dato <= '"& tjkDatoThisSQL &"' AND lk_medarbtype = '"& meType &"'))) ORDER BY lk_dato DESC"
	
    'if session("mid") = 1 then
    'Response.write tjkDatoSQL &"<br>"& strSQL5 & "<br>"
    'end if


    oRec5.open strSQL5, oConn, 3 
	if not oRec5.EOF then
	    lonKorsel_lukketPerDt = oRec5("lk_dato") 
    end if
	oRec5.close 

    'Response.write "lonKorsel_lukketPerDt: "& lonKorsel_lukketPerDt

    'Response.end

end function


public flexSaldoFYreal_norm, slDatoNormPrDdMinus1
function fn_flexSaldoFYreal_norm(mid)

        flexSaldoFYreal_norm = 0

         '** norm til d.d FY
        call licensStartDato()
        stDatoNorm = licensstdato '"1/1/2018" '& year(now)
      

        slDatoNorm = dateAdd("d", -1, now) '"1/1/"& year(now)
        slDatoNormPrDdMinus1 = slDatoNorm
        dageDiff = dateDiff("d",stDatoNorm, slDatoNorm, 2,2) 

        call normtimerPer(mid, stDatoNorm, dageDiff, 0)

        stDatoSQL = year(stDatoNorm) &"/"& month(stDatoNorm) &"/"& day(stDatoNorm) 'year(now)&"/1/1" 
        'slDatoSQL = year(now)&"/12/31" 
        slDatoSQL = dateAdd("d", -1, now)
        slDatoSQL = year(slDatoSQL) &"/"& month(slDatoSQL) &"/"& day(slDatoSQL)

        call akttyper2009(2)

         '*** Henter korrektion inden indtastning ***
        korrektionMaxFlexTimerFYfor = 0 
        strSQLtimer = "SELECT sum(timer) AS realTimerFY FROM timer WHERE tmnr = "& mid &" AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"' AND tfaktim = 114 GROUP BY tmnr" 
        oRec9.open strSQLtimer, Oconn, 3
        if not oRec9.EOF then 

        korrektionMaxFlexTimerFYfor = oRec9("realTimerFY")

        end if
        oRec9.close

        if len(trim(korrektionMaxFlexTimerFYfor)) <> 0 then
        korrektionMaxFlexTimerFYfor = korrektionMaxFlexTimerFYfor
        else
        korrektionMaxFlexTimerFYfor = 0
        end if


        '*** tjekker flex saldo ***
        realTimerFY = 0 
        strSQLtimer = "SELECT sum(timer) AS realTimerFY FROM timer WHERE tmnr = "& mid &" AND tdato BETWEEN '"& stDatoSQL &"' AND '"& slDatoSQL &"' AND ("& aty_sql_realhours &") GROUP BY tmnr" 
        oRec9.open strSQLtimer, Oconn, 3
        if not oRec9.EOF then 

            realTimerFY = oRec9("realTimerFY")

        end if
        oRec9.close


        if len(trim(realTimerFY)) <> 0 then
        realTimerFY = realTimerFY
        else
        realTimerFY = 0
        end if

      
        
        flexSaldoFYreal_norm = ((realTimerFY + (korrektionMaxFlexTimerFYfor)) - (ntimPer*1)) 'Korrektioner ikke med?? + korrektionReal(x) globalfunc3 2277 
      
      


end function
    
%>