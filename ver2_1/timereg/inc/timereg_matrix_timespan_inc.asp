<%
public xtimerthis_mtx
function xmatrixtimespan(idag, mtrx, sTtid, sLtid, datoThis)

                                        useDate = idag '"01-01-" & year(now)
                                        sTtid_org = sTtid


                                        call helligdage(datoThis, 0, lto)

                                        select case mtrx
                                        case 1 'dag
                                                lb = "07:00:00"
                                                ub = "19:00:00"
                                        case 2,4 'aften / Aften Weekend
                                                lb = "19:00:00"
                                                ub = "23:59:59"
                                        case 3 'weekend
                                               lb = "07:00:00"
                                               ub = "19:00:00"
                                       
                                        case 5 'helligdag
                                               lb = "00:00:01"
                                               ub = "23:59:59" 

                                        case 6,7 'NAT / Morgen + Weekend nat (før mødetid)
                                                lb = "00:00:01"
                                                ub = "07:00:00"
                                        end select


                                        if (((mtrx = 1 OR mtrx = 2 OR mtrx = 6) AND (datepart("w", datoThis, 2,2) < 6) AND erHellig <> 1) OR _
                                        ((mtrx = 3 OR mtrx = 4 OR mtrx = 7) AND (datepart("w", datoThis, 2,2) = 6 OR datepart("w", datoThis, 2,2) = 7) AND erHellig <> 1) _
                                        OR (mtrx = 5 AND erHellig = 1)) then 'OR mtrx = 6

                                        

                                                        '*** Beregner timer og klokkeslet efter tarifgrænser *****
                                                        '06:00-23:00
                                                        '05:00-17:00
                                                        '15:00-03:00
                                                        '09:30-15:00
                                                        '05:00-06:00
                                                        '06:00-07:00 
                                                        '21:00-02:00
                                                        '15:00-23:30
                                                        '21:00-16:00
                                                        'START TID
                                                         if (cDate(useDate &" "& sTtid) > cDate(useDate &" "& lb)) then 
                                                            
                                                                   
                                                                    '6 7 Nat: Start tid vil altid være større en 00:00
                                                                    if (cDate(useDate &" "& sTtid) >= cDate(useDate &" "& sLtid) AND mtrx >= 6) then 'NAT                                             
                                                                    sTtid = lb
                                                                    else
                                                                    
                                                                        '21:00-16:00 DAG
                                                                        if (cDate(useDate &" "& sTtid) >= cDate(useDate &" "& sLtid) AND cDate(useDate &" "& sLtid) >= cDate(useDate &" "& lb) AND (mtrx = 1 OR mtrx = 3)) then 'DAG
                                                                        sTtid = lb
                                                                        else     
                                                                        sTtid = sTtid
                                                                        end if

                                                                    end if

                                                        
                                                        else
                                                                  
                                                                   '2 4 Aften
                                                                   if (cDate(useDate &" "& sTtid) >= cDate(useDate &" "& lb) AND (mtrx = 2 OR mtrx = 4)) then 'NAT                                             
                                                                   sTtid = sTtid
                                                                   else
                                                                   sTtid = lb
                                                                   end if
                                                                    
                                                       
                                                        end if


                                                        'SLUT TID
                                                        if ((cDate(useDate &" "& sLtid) <= cDate(useDate &" "& ub)) AND (cDate(useDate &" "& sLtid) >= cDate(useDate &" "& lb))) OR _
                                                        (cDate(useDate &" "& ub) < cDate(useDate &" 07:30:00") AND mtrx < 5) then '< 5 IKKE NAT + helligdag aktiviteter 00 - 07
                                                            
                                                                sLtid = sLtid
                                                        
                                                        else
                                                                    
                                                                   '2 4 Aften
                                                                   '15-03
                                                                   if (mtrx = 2 OR mtrx = 4) then 'AFTEN
                                                                       
                                                                       '09:00 - 15:00 < en start tid for aften
                                                                       '15:00 - 03:00
                                                                       '06:00 - 07:00
                                                                       '05:00 - 02:00
                                                                       '21:00 - 23:00
                                                                       '15:00 - 21:00
                                                                       '05:00 - 23:00
                                                                       if (cDate(useDate &" "& sLtid) <= cDate(useDate &" "& lb)) then
                                                                                
                                                                                    if (cDate(useDate &" "& sTtid_org) >= cDate(useDate &" "& sLtid)) then                                             
                                                                                    sLtid = ub 
                                                                                    else
                                                                                    sLtid = "00:01" ' = 0 timer. ALTID NUL TIMER
                                                                                    end if
                                                                       else
                                                                       
                                                                                    '** Hvis sluttid > 23:59
                                                                                    if (cDate(useDate &" "& sLtid) => cDate(useDate &" "& ub)) then
                                                                                    sLtid = ud
                                                                                    else
                                                                                    sLtid = sLtid
                                                                                    end if                                                                   

                                                                       end if

                                                                   else
                        
                                                                        if (mtrx = 1 OR mtrx = 3) then 'DAG / WEEKEND DAG

                                                                                '** Hvis sluttid er før starttid på aften 
                                                                                '06:00 - 02:00 --> 12 timer på dag USE: UB :OK
                                                                                '15:00 - 02:00 --> 4 timer på Dag  USE: UB :OK
                                                                                '21:00 - 02:00 --> 0 timer på Dag  USE: UB :OK
                                                                                '06:00 - 07:00 --> 0 timer på Dag  USE: UB/sl :OK
                                                                                '05:00 - 06:00 --> 0 timer på Dag  USE: UB: OK
                                                                                '05:00 - 15:00 --> 8 timer på Dag  USE: sl :OK
                                                                                '09:00 - 15:00 --> 6 timer på Dag  USE: sl :OK
                                                                                '09:00 - 21:00 --> 10 timer på Dag USE: UB: OK
                                                                                '21:00 - 16:00 --> 9 timer på Dag USE: 
                                                                                if (cDate(useDate &" "& sLtid) < cDate(useDate &" "& lb)) AND (cDate(useDate &" "& sTtid_org) <= cDate(useDate &" "& sLtid)) then
                                                                                sLtid = sLtid
                                                                                else

                                                                                                                                                            
                                                                                    if (cDate(useDate &" "& sLtid) < cDate(useDate &" "& ub)) then
                                                                                    sLtid = sLtid
                                                                                    else 
                                                                                    sLtid = ub
                                                                                    end if  
                                                                    
                                                                                end if
                                                                        else
                                                                        sLtid = ub
                                                                        end if 

                                                                   end if
                                                        
                                                        end if

                                                        timerthis_mtx = dateDiff("h", useDate &" "& sTtid, useDate &" "& sLtid, 2,2)
                                                        totalmin = datediff("n", idag &" "& sTtid, idag &" "& sLtid)
						                                call timerogminutberegning(totalmin)
						                                timerthis_mtx = thoursTot&"."&tminProcent 'tminTot
						
						 
					                                    if timerthis_mtx < 0 then '** Hen over kl 24:00 **'
                                                        timerthis_mtx = 0
						                                'timerthis_mtx = 24 + (replace(timerthis_mtx,".",","))
						                                'timerthis_mtx = replace(timerthis_mtx,",",".")
						                                end if


                                                        '*** Mandag morgen --> Normal dag

                                                        'response.write "<br> = mtrx: "& mtrx &" matrixAktFundet: "& matrixAktFundet & " timerthis_mtx: "& timerthis_mtx & " sTtid: "& sTtid & " sLtid: "& sLtid & " datepart: " & datepart("w", datoThis, 2,2) & "##"& cDate(useDate &" "& sTtid) &"<="& cDate(useDate &" "& sLtid)
                

                                     else        

                                        timerthis_mtx = 0

                                     end if 'mtrx 1-7 + weekend/helligdage   
    
end function   
%>