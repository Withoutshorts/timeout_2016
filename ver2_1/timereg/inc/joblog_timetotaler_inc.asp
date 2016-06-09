<%

    public strMedarbOskriftLinie, expTxt
    sub medarboSkriftlinje


             if cint(directexp) <> 1 then
    
			strMedarbOskriftLinie = strMedarbOskriftLinie & "<tr>"
			strMedarbOskriftLinie = strMedarbOskriftLinie & "<td height="& tdh &" valign=bottom style='padding:2px; border-top:1px #CCCCCC solid; width:250px;' bgcolor='#F7F7F7'>Kunde<br><b>Job</b> (aktiviteter)</td>"
				
				
				
				'*** job totaler oskrifter  **'
				
						
							'if jobid <> 0 then
							strFakbtimTxt = "<b>Budget.</b><br>(forkalkuleret)"
							'else
							'strFakbtimTxt = "Forkalk. timer&nbsp;"
							'end if
						
							
							strFakbTxt = "Brutto Oms.&nbsp;"

                            'if jobid = 0 then
                            'strFakbTxt = strFakbTxt & "<br><span style='color:#6CAE1C; font-size:11px;'>Faktisk timepris</span>&nbsp;"
                            'end if
						
						
				
				strMedarbOskriftLinie = strMedarbOskriftLinie & "<td "&tdstyleTimOms&" bgcolor=#F7F7F7>"& strFakbtimTxt 
				
				if cint(visfakbare_res) = 1 then
				strMedarbOskriftLinie = strMedarbOskriftLinie & "<br>"& strFakbTxt &"</td>"
				else
				strMedarbOskriftLinie = strMedarbOskriftLinie & "</td>"
				end if


                            '*******************************************'
                            '****** Prev Saldo *************************'
                            '*******************************************'
                            if cint(visPrevSaldo) = 1 then
                            strMedarbOskriftLinie = strMedarbOskriftLinie & "<td "&tdstyleTimOms&" bgcolor=#F7F7F7>"
                
                             if cint(vis_restimer) = 1 then
                             strMedarbOskriftLinie = strMedarbOskriftLinie &"<span style='color:#999999; font-size:9px;'>Forecasttimer</span><br>"
                             end if
                
                             strMedarbOskriftLinie = strMedarbOskriftLinie &"<b>Real. timer</b><br>"
                    
                    
                
                            strMedarbOskriftLinie = strMedarbOskriftLinie &"(før valgte periode)</td>"
               
                            end if
                

               

				
				            strMedarbOskriftLinie = strMedarbOskriftLinie & "<td "&tdstyleTimOms&" bgcolor=#F7F7F7>"
                
                             if cint(vis_restimer) = 1 then
                             strMedarbOskriftLinie = strMedarbOskriftLinie &"<span style='color:#999999; font-size:9px;'>Forecasttimer</span><br>"
                             end if

                            strMedarbOskriftLinie = strMedarbOskriftLinie &jobaktOskrift&" "
				
				                select case cint(visfakbare_res) 
                                case 1
				                strMedarbOskriftLinie = strMedarbOskriftLinie & "<b>Realiseret tim.</b>&nbsp;"
				                strMedarbOskriftLinie = strMedarbOskriftLinie & "<br>Omsætning&nbsp;<br>Balance<br>(i periode)&nbsp;"
				    
				                case 2

                                strMedarbOskriftLinie = strMedarbOskriftLinie & "<b>Realiseret tim.</b>&nbsp;"
				                strMedarbOskriftLinie = strMedarbOskriftLinie & "<br>Kost. ialt&nbsp;<br>Balance<br>(i periode)&nbsp;"

                                case else
				                strMedarbOskriftLinie = strMedarbOskriftLinie & "<b>Real. timer</b><br>(i periode)&nbsp;"
				    
				                    'if cint(vis_enheder) = 1 then
				                    'strMedarbOskriftLinie = strMedarbOskriftLinie &"<br>Enheder&nbsp;"
				                    'end if
				    
				                end select
				
				            strMedarbOskriftLinie = strMedarbOskriftLinie & "</td>"


                    '*****************************************'
                    '********** Timer total uasent periode ***'
                    '*****************************************'

                    if cint(visPrevSaldo) = 1 then 
                 
                    strMedarbOskriftLinie = strMedarbOskriftLinie & "<td "&tdstyleTimOms&" bgcolor=#F7F7F7>"

                     if cint(vis_restimer) = 1 then
                     strMedarbOskriftLinie = strMedarbOskriftLinie &"<span style='color:#999999; font-size:9px;'>Forecasttimer</span><br>"
                     end if

                     strMedarbOskriftLinie = strMedarbOskriftLinie & "<b>Real. timer</b><br>(ialt)</td>"

                    end if

            
                end if 'if cint(directexp) <> 1 then 
				

                
                expTxt = expTxt &"Kontakt;Jobnavn;Jobnr;Fase;Aktivitet;Jobtype;"
			    expTxt = expTxt &"Jobansvarlig;Init;Jobejer;Init;"
                expTxt = expTxt &"Timer forkalkuleret;"

                if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
				expTxt = expTxt &"Bruttooms (Budget);"
                end if

                

               'Kun til CSV spar denne på det almindelige udtræk 
               call exportptOskrifter
                

                if cint(csv_pivot) = 1 then
                 expTxt = expTxt &"Medarb. [init]/type;" 'Kun PIVOT
                end if


			

                '*** Overskrifter og medarb ***'
				for v = 0 to v - 1
						
                        
                        medabTottimer(v) = 0
						restimerTot(v) = 0
						omsTot(v) = 0
						fakbareTimerTot(v) = 0
						normTimerTot(v) = 0
						medabTotenh(v) = 0 
						medabRestimer(v) = 0

                        subMedabRestimer(v) = 0
                        subMedabTottimer(v) = 0
                        subMedabTotenh(v) = 0
                        omsSubTot(v) = 0

                
                        if cint(directexp) <> 1 then 

                                select case right(v, 1)
                                case 0,2,4,6,8
                                bgthis = "#F7F7F7"
                                case else
                                bgthis = "#F7F7F7"
                                end select
						
						'** Medarb 1 Row ***

						strMedarbOskriftLinie = strMedarbOskriftLinie & "<td colspan="& md_split_cspan &" "&tdstyleTimOms1&" style='width:50px;'  bgcolor='"& bgthis &"'><img src='../ill/blank.gif' height=1 width=50><br><span style='color:#000000; font-size:9px;'>"& medarbnavnognr_short(v) &"</span>"
						
						end if 'if cint(directexp) <> 1 then 	

                            for c = 0 to etal
                            
                                if cint(csv_pivot) = 0 OR (cint(csv_pivot) = 1 AND v = 0 AND c = 0) then
                                expTxt = expTxt &"Real. timer;"
                            

                                    if cint(csv_pivot) <> 1 then

                                        if cint(vis_restimer) = 1 then
                                        expTxt = expTxt &"Res. timer;"
                                        end if

							            if cint(vis_enheder) = 1 then
							            expTxt = expTxt &"Enheder;" 
							            end if
							
							            if cint(visfakbare_res) = 1 then
							            expTxt = expTxt &"Omsætning;"
							            expTxt = expTxt &"Timepris;"
							            end if


                                        if cint(visfakbare_res) = 2 then
							            expTxt = expTxt &"Kostpris;"
							            expTxt = expTxt &"Kostpris pr. time;"
							            end if

                                    end if

                                end if 

                            next

                         
						
						    
						if cint(directexp) <> 1 then 
						strMedarbOskriftLinie = strMedarbOskriftLinie & "</td>"
                        end if 'if cint(directexp) <> 1 then 
						
						
				next
				
				
                expTxt = expTxt &"xx99123sy#z"


                

                '*** Medarb linier på eksport ***'
                if cint(csv_pivot) <> 1 then 'IKKE PIVOT

                for v = 0 to v - 1

                        if v = 0 then
                        '*** Kun export ****
                        call tommeCSVfelter
                        end if

                   

                            expTxt = expTxt & replace(medarbnavnognr(v), "<br>", " ") &";"
							

                            
                            for c = 0 to etal
                            
                            if c <> 0 then
                            expTxt = expTxt &";"
                            end if

                            if cint(vis_restimer) = 1 then
                            expTxt = expTxt &";"
                            end if

							if cint(vis_enheder) = 1 then
							expTxt = expTxt &";" 
							end if
							
							if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
							expTxt = expTxt &";"
							expTxt = expTxt &";"
							end if


                            next

                            

                next



                '**** Måned ***'
                if cint(md_split) = 1 OR cint(md_split) = 2 then

				for v = 0 to v - 1
                                                

                                               

                                            

                                              

                                                if v = 0 then
                            
                                                if cint(directexp) <> 1 then 

                                                strMedarbOskriftLinie = strMedarbOskriftLinie &"</tr><tr bgcolor=#EFf3ff><td colspan=3>&nbsp;</td>"
                                                
                                                    if cint(visPrevSaldo) = 1 then
                                                    strMedarbOskriftLinie = strMedarbOskriftLinie &"<td>&nbsp;</td><td>&nbsp;</td>"
                                                    end if

                                                end if 'if cint(directexp) <> 1 then 

                                                expTxt = expTxt &"xx99123sy#z"

				                                call tommeCSVfelter
                                               
                                                end if




                                                  if cint(md_split) = 1 then '** 3 md

                                                  if cint(directexp) <> 1 then 
                                                 strMedarbOskriftLinie = strMedarbOskriftLinie &"<td align=center style='font-size:8px; border-left:1px #CCCCCC solid'>"& mdThis1 &"<br><img src='../ill/blank.gif' width=45 height=1></td>"_
                                                 &"<td align=center style='font-size:8px;'>"& mdThis2 &"<br><img src='../ill/blank.gif' width=45 height=1></td>"_
                                                 &"<td align=center style='font-size:8px;'>"& mdThis3 &"<br><img src='../ill/blank.gif' width=45 height=1></td>"
                                                 end if 'if cint(directexp) <> 1 then 


                                                  expTxt = expTxt & replace(mdThis1, "<br>", "") &";"



                                                  if cint(vis_restimer) = 1 then
                                                 expTxt = expTxt &";"
                                                 end if

                                                if cint(vis_enheder) = 1 then
							                    expTxt = expTxt &";" 
							                    end if
							
							                    if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
							                    expTxt = expTxt &";"
							                    expTxt = expTxt &";"
							                    end if    

                                                 expTxt = expTxt & replace(mdThis2, "<br>", "") &";"

                                                  if cint(vis_restimer) = 1 then
                                                 expTxt = expTxt &";"
                                                 end if

                                                if cint(vis_enheder) = 1 then
							                    expTxt = expTxt &";" 
							                    end if
							
							                    if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
							                    expTxt = expTxt &";"
							                    expTxt = expTxt &";"
							                    end if    


                                                 expTxt = expTxt & replace(mdThis3, "<br>", "") &";"


                                                  if cint(vis_restimer) = 1 then
                                                 expTxt = expTxt &";"
                                                 end if

                                                if cint(vis_enheder) = 1 then
							                    expTxt = expTxt &";" 
							                    end if
							
							                    if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
							                    expTxt = expTxt &";"
							                    expTxt = expTxt &";"
							                    end if    


                                                 else

                                                   for m = 1 to 12
                                                         
                                                         mdThis = dateadd("m", -(12-m), datoSlut)
                                                         mdThis = left(monthname(month(mdThis)), 3) &"<br>"& year(mdThis)
                                            
                                                         if cint(directexp) <> 1 then 
                                                         strMedarbOskriftLinie = strMedarbOskriftLinie &"<td align=center style='font-size:8px; border-left:1px #CCCCCC solid'>"& mdThis &"<br><img src='../ill/blank.gif' width=45 height=1></td>"
                                                         end if 'if cint(directexp) <> 1 then                                                      

                                                         expTxt = expTxt & replace(mdThis, "<br>", "") &";" 

                                                         
                                                         if cint(vis_restimer) = 1 then
                                                         expTxt = expTxt &";"
                                                         end if

                                                        if cint(vis_enheder) = 1 then
							                            expTxt = expTxt &";" 
							                            end if

                                                        if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
                                                        expTxt = expTxt &";;"
                                                        end if

                                                   next
   
                                                     
                                                      


                                                 end if


                                                
                                                 
                                                 

							




                next
                antalV = v

            
                end if 'PIVOT csv_pivot = 1

                end if									

				'strMedarbOskriftLinie = strMedarbOskriftLinie & "</tr>"	


    end sub






sub subTotaler_gt


            '***************************'
			'*** Sub totaler i bunden af hvert job ved udspec ******'
			'***************************'
			
			

               
				
			    strJobLinie_Subtotal = "<tr>"
				strJobLinie_Subtotal = strJobLinie_Subtotal & "<td style='padding:4px; border-top:1px #CCCCCC solid;' valign=bottom bgcolor=snow><b>Jobtotal:</b></td>"
						
						strJobLinie_Subtotal = strJobLinie_Subtotal & "<td class=lille valign=bottom align=right "&tdstyleTimOms2&" bgcolor=snow>" 
						
						
						

                        strJobLinie_Subtotal = strJobLinie_Subtotal & formatnumber(subbudgettimer, 2) & " t."

                        if cint(vis_enheder) = 1 then
					    strJobLinie_Subtotal = strJobLinie_Subtotal & "<br>"
						end if

						
					    if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
						strJobLinie_Subtotal = strJobLinie_Subtotal & "<br><span style='color=#000000; font-size:8px;'>"& basisValISO &" "& formatnumber(subbudget, 2)&"</span><br>"
                        
                            if jobid = 0 then
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "<span style='color:green; font-size:8px;'>"& basisValISO &" "& formatnumber(TimeprisFaktiskSub, 2) &"</span><br></td>"
                            else
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "</td>"
                            end if
							
								
							
						else
						strJobLinie_Subtotal = strJobLinie_Subtotal & "</td>"
						end if
						

                        '**** Saldo før ***'
                        if cint(visPrevSaldo) = 1 then
                        strJobLinie_Subtotal = strJobLinie_Subtotal & "<td class=lille valign=bottom align=right "&tdstyleTimOms2&" bgcolor=snow>"
                        
                            if cint(vis_restimer) = 1 then
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "<span style='color:#999999;'>"& formatnumber(saldoRestimerSub,0) &"</span><br>"
                            end if


                        strJobLinie_Subtotal = strJobLinie_Subtotal & formatnumber(saldoJobSub, 2) 

                        '*** Enheder ***'
						if cint(vis_enheder) = 1 then
                        strJobLinie_Subtotal = strJobLinie_Subtotal &"<br><span style='color:#5c75AA; font-size:9px;'> enh. "& formatnumber(enhederPrevSaldoSub,2)  & "</span>"
                        end if
                        
                        strJobLinie_Subtotal = strJobLinie_Subtotal &"</td>"

                     
                        end if

                        

						'*** Jobtotaler IALT i periode ***'
						strJobLinie_Subtotal = strJobLinie_Subtotal & "<td class=lille valign=bottom align=right "&tdstyleTimOms10&" bgcolor=snow>" 
						

                        '**** Res timer ***'
                         if cint(vis_restimer) = 1 then
                         strJobLinie_Subtotal = strJobLinie_Subtotal & "<span style='color:#999999;'>"& formatnumber(restimerSubJob,0) &"</span><br>"
                         end if

                         strJobLinie_Subtotal = strJobLinie_Subtotal & formatnumber(subtotaljboTimerIalt,2)

						'*** Enheder ***'
						if cint(vis_enheder) = 1 then
                        strJobLinie_Subtotal = strJobLinie_Subtotal &"<br><span style='color:#5c75AA; font-size:9px;'> enh. "& formatnumber(subJobEnh,2)  & "</span>"
                        end if
                                

                               
						if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
						strJobLinie_Subtotal = strJobLinie_Subtotal & "<br><span style='color=#000000; font-size:8px;'>"& basisValISO &" "&formatnumber(subtotaljobOmsIalt, 2)& "</span>" 
						strJobLinie_Subtotal = strJobLinie_Subtotal & "<br><font class=megetlillesilver>bal.: "&formatnumber(subdbialt, 2)&"</td>"
						else
						strJobLinie_Subtotal = strJobLinie_Subtotal & "</td>"
						end if
						
						

                        '********** Grandtotal uanset periode ******'
                        if cint(visPrevSaldo) = 1 then
                        
                        strJobLinie_Subtotal = strJobLinie_Subtotal & "<td class=lille valign=bottom align=right "&tdstyleTimOms20&" bgcolor=snow>"

                        '**** Res timer ***'
                        if cint(vis_restimer) = 1 then
                        strJobLinie_Subtotal = strJobLinie_Subtotal & "<span style='color:#999999;'>"& formatnumber(restimerSubGtotalJob,0) &"</span><br>"
                        end if
                           
                        
                       

                       

                        if formatnumber(timerTotSaldoSubGtotal) <> formatnumber(saldoJobSub/1+subtotaljboTimerIalt) then
                        fntCol = "darkred"
                        tSign = "~"
                        else
                        fntCol = "#000000"
                        tSign = "="
                        end if

                         strJobLinie_Subtotal = strJobLinie_Subtotal &"<span style='color:"& fntCol &";'> "&tSign&" <b>"& formatnumber(timerTotSaldoSubGtotal, 2) &"</b></span>"


                         '*** Enheder ***'
						if cint(vis_enheder) = 1 then
                           if formatnumber(enhederGSub) <> formatnumber(subJobEnh/1+enhederPrevSaldoSub)  then
                                efntCol = "#3b5998"
                                eSign = "~"
                                else
                                efntCol = "#5c75AA"
                                eSign = "="
                                end if
                         strJobLinie_Subtotal = strJobLinie_Subtotal &"<br><span style='color:"&efntCol&"; font-size:9px;'> "& eSign &" enh. "& formatnumber(enhederGSub,2)  & "</span>"
                         end if

                         strJobLinie_Subtotal = strJobLinie_Subtotal &"</td>"

						end if 
						
				
				
				
				'if cint(jobid) = 0 then
				
						for v = 0 to v - 1
						strJobLinie_Subtotal = strJobLinie_Subtotal & "<td class=lille colspan="& md_split_cspan &" valign=top align=right "&tdstyleTimOms3&" bgcolor=snow>"
                        
                         if cint(vis_restimer) = 1 then
                            if subMedabRestimer(v) <> 0 then
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "<span style='color:#999999;'>"& formatnumber(subMedabRestimer(v),0) &"</span><br>"
                            else
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "&nbsp;<br>"
                            end if
                         subMedabRestimer(v) = 0
                         end if
                        
                        if subMedabTottimer(v) <> 0 then
                        strJobLinie_Subtotal = strJobLinie_Subtotal & formatnumber(subMedabTottimer(v), 2)
                        else
                        strJobLinie_Subtotal = strJobLinie_Subtotal & "&nbsp;<br>"
                        end if
                          
                        subMedabTottimer(v) = 0

                        'strJobLinie_Subtotal = strJobLinie_Subtotal & " % "& formatnumber(subMedabTottimer(v), 2)  
						
						if cint(vis_enheder) = 1 then
						    if subMedabTotenh(v) <> 0 then
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "<br><span style='color:#5c75AA; font-size:9px;'>enh. " & formatnumber(subMedabTotenh(v), 2) & "</span>"
                            else
                            strJobLinie_Subtotal = strJobLinie_Subtotal & "&nbsp;<br>"
                            end if
                        subMedabTotenh(v) = 0
						end if
						        
                               
                            if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
							
                                if omsSubTot(v) <> 0 then
                                strJobLinie_Subtotal = strJobLinie_Subtotal & "<br><span style='color=#000000; font-size:8px;'>"& basisValISO &" "&formatnumber(omsSubTot(v), 2)&"</span></td>"
                                else
                                strJobLinie_Subtotal = strJobLinie_Subtotal & "<br>&nbsp;</td>"
                                end if
                            omsSubTot(v) = 0
								
							else
							strJobLinie_Subtotal = strJobLinie_Subtotal & "</td>"
							end if
						
							
						next
				 ' end if 'if visMedarb = 1 then
			strJobLinie_Subtotal = strJobLinie_Subtotal & "</tr>"
			

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
          
            

			
			'Response.write strJobLinie_Subtotal
            strJobLinie = strJobLinie & strJobLinie_Subtotal

            

       end sub


   sub tommeCSVfelter


                   
				    expTxt = expTxt &";;;;;;;;;;;;"
				   

                    if cint(visPrevSaldo) = 1 then
                    expTxt = expTxt &";;"

                            if cint(vis_restimer) = 1 then
                            expTxt = expTxt &";;;"
                            end if

                            if cint(vis_enheder) = 1 then
                            expTxt = expTxt &";;;"
                            end if
                    
                    else

                            if cint(vis_restimer) = 1 then
                            expTxt = expTxt &";"
                            end if

                            if cint(vis_enheder) = 1 then
                            expTxt = expTxt &";"
                            end if

                    end if


                  
							
					if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
					expTxt = expTxt &";;;"
                    end if    

       end sub





sub exportptOskrifter

                if cint(visPrevSaldo) = 1 then
                expTxt = expTxt &"Real. timer Ialt (Før valgte periode);" 
                
                    
                    if cint(vis_restimer) = 1 then
                    expTxt = expTxt &"Ressource timer (Før valgte periode);"
                    end if
                    

                    if cint(vis_enheder) = 1 then
                    expTxt = expTxt &"Enheder (Før valgte periode);"
                    end if

                end if

                
           


                expTxt = expTxt &"Real. timer Ialt (I periode);" 


                    
                    if cint(vis_restimer) = 1 then
                    expTxt = expTxt &"Ressource timer (I periode);"
                    end if
                    

                    if cint(vis_enheder) = 1 then
                    expTxt = expTxt &"Enheder (I periode);"
                    end if


                if cint(visfakbare_res) = 1 then
				expTxt = expTxt &"Omsætning ialt (I periode);Balance;"
				end if
				
                if cint(visfakbare_res) = 2 then
				expTxt = expTxt &"Kostpris ialt (I periode);Balance;"
				end if

                if cint(visPrevSaldo) = 1 then
                expTxt = expTxt &"Real. timer (Ialt, uanset per.);" 

                    
                    if cint(vis_restimer) = 1 then
                    expTxt = expTxt &"Ressource timer (Ialt, uanset per.);"
                    end if
                    

                    if cint(vis_enheder) = 1 then
                    expTxt = expTxt &"Enheder (Ialt, uanset per.);"
                    end if
                
                
                end if

end sub




 sub tomtdfelt_a

            if lastWrtMd = -1 OR (m > lastWrtMd AND m < md) then
                    
                    if cint(directexp) <> 1 then 
                    strJobLinie = strJobLinie & "<td class=lille align=right "& tdstyleTimOms3 &" bgcolor='"& bgthis &"'>&nbsp;"
                    'strJobLinie = strJobLinie & "<br>lastWrtMd:"& lastWrtMd &" cnt: "& cnt &" m:"& m & "<br> md: "& md & " md_md:" & md_md
                    end if ' if cint(directexp) <> 1 then

                    expTxt = expTxt &";"


                            '** Res timer ***'
                            if cint(vis_restimer) = 1 then
                                    
                                li = "a"
                            
                                    if cint(md_split_cspan) = 12 then
                                    
                                            '**find lasMD Wrt
                                            if  lastWrtMd = "-1" then
                                            md_yearA = dateAdd("m", m-1, datoStart)
                                            else
                                            md_yearA = dateAdd("m", m, datoStart)
                                            end if
                                    else
                                    md_yearA = dateAdd("m", 0, datoStart)
                                    end if

                                md_mdA = month(md_yearA)
                                md_yearA = year(md_yearA)
                                  

                                call resTimer(jobmedtimer(x,0), jobmedtimer(x,12), jobmedtimer(x,4), md_mdA, md_yearA, md_split_cspan, 0)
                              
                                if cint(directexp) <> 1 then  
                                strJobLinie = strJobLinie & "<span style='color:#999999;'>"& restimerThis &"</span><br>&nbsp;"
                                end if 'if cint(directexp) <> 1 then 
                        
                                        if len(trim(restimerThis)) <> 0 then
                                        restimerThis = restimerThis
                                        restimerThisExp = restimerThis
                                        else
                                        restimerThis = 0
                                        restimerThisExp = ""
                                        end if

                                         medabTotRestimerprMd(v, m-1) = (medabTotRestimerprMd(v, m-1)/1 + restimerThis/1)

                                        medabRestimer(v) = medabRestimer(v) + restimerThis
                                        subMedabRestimer(v) = subMedabRestimer(v) + restimerThis
                                     
                                    expTxt = expTxt & restimerThisExp &";"
                                
                                            
                            end if

                                    
                                    
                            if cint(vis_enheder) = 1 then
							expTxt = expTxt &";" 
							end if
							
							if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
							expTxt = expTxt &";"
							expTxt = expTxt &";"
							end if    


                       
                end if 

                    if cint(directexp) <> 1 then 
                    strJobLinie = strJobLinie & "</td>"
                    end if' if cint(directexp) <> 1 then 
                    'loopsFO = loopsFO + 1


    end sub
	


   







	sub timerTdfelt
	

                                        lastWrtM = jobmedtimer(x, 4)
                                        lastWrtMd = md 'datepart("m", jobmedtimer(x, 36), 2,2)
                                
                                        if jobmedtimer(x, 38) = 0 then
                                        lastWrtJ = jobmedtimer(x, 0)
                                        else
                                        lastWrtJ = jobmedtimer(x, 12)
                                        end if

                                        lastWrtX = x

                                      
                                        if cint(directexp) <> 1 then 
                                        '**** Timeantal ****'
								        strJobLinie = strJobLinie & "<td class=lille align=right "& tdstyleTimOms3 &" bgcolor='"& bgthis &"'>"
                                        'strJobLinie = strJobLinie & "&nbsp;cnt:"& cnt &" md: " & md & "<br> mdmd: "& md_md & " m:" & m & " <br>" & jobmedtimer(x, 36) & " medarb(v) : "& medarb(v) &"<br>"
                                        '".." & v &"'//12: "& jobmedtimer(x, 12) &" Udsepc: "& upSpec &" - jobid: "& jobid & " 4: "& jobmedtimer(x, 4) &" mv: "& medarb(v) &" .. "'udSpec = "& upSpec &" a: "& jobmedtimer(x, 12) &" j:" & jobmedtimer(x,0) & " m:"& jobmedtimer(x,4) & " dt:" & jobmedtimer(x, 36) &" 38:" & jobmedtimer(x, 38) & "<br>"
								        end if
                                 


                                         if cint(vis_restimer) = 1 then
                                    
                                   

                                            li = "v"

                                            if cint(md_split_cspan) = 12 then
                                            mdThisV = md_md 'dateAdd("m", m, datoStart)
                                            else
                                            mdThisV = md_md
                                            end if

                                            call resTimer(jobmedtimer(x,0), jobmedtimer(x,12), jobmedtimer(x,4), mdThisV, md_year, md_split_cspan, 0)
                                 
                                            if cint(directexp) <> 1 then 
                                            strJobLinie = strJobLinie & "<span style='color:#999999;'>"& restimerThis &"</span><br>"
                                            end if

                                            vCntVal = vCntVal + 1 

                                            if len(trim(restimerThis)) <> 0 then
                                            restimerThis = restimerThis
                                            restimerThisExp = restimerThis
                                            else
                                            restimerThis = 0
                                            restimerThisExp = ""
                                            end if

                                            medabRestimer(v) = medabRestimer(v) + restimerThis
                                            subMedabRestimer(v) = subMedabRestimer(v) + restimerThis
                                            medabTotRestimerprMd(v, md) = (medabTotRestimerprMd(v, md)/1 + restimerThis/1)
                                            
                                  
                                         
                                         end if

                                          
                
                                        if cint(directexp) <> 1 then 


								            '*** timer ***
								            if jobmedtimer(x,3) <> 0 then
                                    
                                   

									            strJobLinie = strJobLinie & formatnumber(jobmedtimer(x,3), 2)  

                                                if cint(vis_enheder) = 1 then
                                                    if formatnumber(jobmedtimer(x, 25)) <> 0 then
						                            strJobLinie = strJobLinie & " <br><span style='color:#5c75AA; font-size:9px;'> enh. "& formatnumber(jobmedtimer(x, 25), 2) & "</span>" 
                                                    else
                                                    strJobLinie = strJobLinie & "<br>&nbsp;"
                                                    end if
						                        end if
							                else
									            strJobLinie = strJobLinie & "<br>&nbsp;"
								            end if 



                                       end if 'if cint(directexp) <> 1 then 











								
								        if jobmedtimer(x,3) <> 0 then
                
                                            if cint(csv_pivot) = 1 then 'PIVOT
                                            expTxt = expTxt & replace(medarbnavnognr(v), "<br>", "") &";" 'KUN PIVOT
                                            end if

                                        'lastMidstrId = jobmedtimer(x,4)
                                        'Real Timer
								        expTxt = expTxt & formatnumber(jobmedtimer(x,3), 2)
								        expTxt = expTxt &";"
								        else

								            if cint(csv_pivot) <> 1 then
                                            expTxt = expTxt &";"
                                            end if
								        
                                        end if
								

                                        '*** Res timer **'
                                        if cint(vis_restimer) = 1 then
                                                 if cint(csv_pivot) <> 1 then
                                                 expTxt = expTxt & restimerThisExp &";"
                                                 end if
                                        end if


								        '*** Enheder ***
								        if cint(vis_enheder) = 1 then
								        
								                if len(jobmedtimer(x, 25)) <> 0 then
                                                enhThisTxt = formatnumber(jobmedtimer(x, 25), 2)
								                else
								                enhThisTxt = ""
                                                end if
                                        
                                        expTxt = expTxt & enhThisTxt &";"

								        
								
								        end if
								
                                        
                                


								        '*** Omsætning ***'
								        if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
									    

                                           if cint(directexp) <> 1 then 
									         

                                                if jobmedtimer(x,3) <> 0 then
									                strJobLinie = strJobLinie & "<br><span style='color=#000000; font-size:8px;'>"& basisValISO &" " & formatnumber(jobmedtimer(x,7), 2) & "</span>"
                                                else
                                                    strJobLinie = strJobLinie & ""
                                                end if
									  
								
								                    '*** Timepris ***'
								    
								                        if jobmedtimer(x,7) <> 0 AND jobmedtimer(x,3) <> 0 then
								                        medTpris = formatnumber((jobmedtimer(x,7) / jobmedtimer(x,3)), 2)
								                        else
								                        medTpris = 0
								                        end if
								    
								        
                                                    if vis_redaktor = 1 AND cint(visfakbare_res) = 1 then
                                                    
                                                   
                                                    strJobLinie = strJobLinie &"<br><input type='text' class='f_tp_"&jobmedtimer(x, 4)&"_"&jobmedtimer(x, 0)&"_"&jobmedtimer(x, 12)&"' id='f_tp_"&x&"' value='"& formatnumber(medTpris, 2) & "' name='FM_tp_t' style='color=#6CAE1C; padding:0px; width:40px; height:14px; line-height:9px; font-size:9px;'> /t. <span class='s_tp' id='s_tp_"&jobmedtimer(x, 4)&"_"&jobmedtimer(x, 0)&"_"&jobmedtimer(x, 12)&"__"&x&"' style='font-size:9px; color:#3B5998;'><=></span>"
                                                    strJobLinie = strJobLinie &"<input type='hidden' value='#' name='FM_tp_t'>"
                                                    strJobLinie = strJobLinie &"<input type='hidden' value='"& jobmedtimer(x, 6) & "' name='FM_jobnr_t'>"
                                                    strJobLinie = strJobLinie &"<input type='hidden' value='"& jobmedtimer(x, 0) & "' name='FM_jobid_t'>"
                                                    strJobLinie = strJobLinie &"<input type='hidden' value='"& jobmedtimer(x, 12) & "' name='FM_aktid_t'>"
                                                    strJobLinie = strJobLinie &"<input type='hidden' value='"& jobmedtimer(x, 4) & "' name='FM_medid_t'>"
                                                    strJobLinie = strJobLinie &"<input type='hidden' value='"& jobmedtimer(x, 36) & "' name='FM_mth_t'>"
                                                  
                                                    else

                                                        if jobmedtimer(x,3) <> 0 then
                                                            if cint(visfakbare_res) = 1 then
                                                            strJobLinie = strJobLinie &"<br><span style='color=#6CAE1C; font-size:9px;'>"& formatnumber(medTpris, 2) & " /t.</span>"
                                                            else
                                                            strJobLinie = strJobLinie &"<br><span style='color=#FF0000; font-size:9px;'>"& formatnumber(medTpris, 2) & " /t.</span>"
                                                            end if
                                                        else
                                                            strJobLinie = strJobLinie &""
                                                        end if


                                                    end if
								                


                                            end if 'if cint(directexp) <> 1 then 
								
										
										        if jobmedtimer(x,7) > 0 then
										        expTxt = expTxt &formatnumber(jobmedtimer(x,7), 2)
										        expTxt = expTxt &";"&formatnumber(medTpris, 2)&";"
										        else
										        expTxt = expTxt &"0;"&formatnumber(medTpris, 2)&";"
										        end if
										
								        end if
								
                                        loopsFO = md 'loopsFO + 1

                                        
                                        if cint(directexp) <> 1 then 
								        strJobLinie = strJobLinie &"</td>"
	                                    end if 'if cint(directexp) <> 1 then 							

							 	        medabTottimer(v) = medabTottimer(v) + jobmedtimer(x,3)
                                        subMedabTottimer(v) = subMedabTottimer(v) + jobmedtimer(x,3)

                                        '*** tot timer pr. medar. pr. md
                                        medabTottimerprMd(v, md) = (medabTottimerprMd(v, md)/1 + jobmedtimer(x,3)/1)

                                     
							 	
							 	        if cint(vis_enheder) = 1 then
							 	        medabTotenh(v) = medabTotenh(v) + jobmedtimer(x,25)
                                        subMedabTotenh(v) = subMedabTotenh(v) + jobmedtimer(x,25) 
                                        medabTotEnhprMd(v, md) = (medabTotEnhprMd(v, md)/1 + jobmedtimer(x,25)/1) 
							 	        end if

                                         if cint(visfakbare_res) = 1 OR cint(visfakbare_res) = 2 then
                                         medabTotOmsprMd(v, md) = (medabTotOmsprMd(v, md)/1 + jobmedtimer(x,7)/1)  
                                         end if
							 	
								        omsTot(v) = omsTot(v) + jobmedtimer(x,7)
                                        omsSubTot(v) = omsSubTot(v) + jobmedtimer(x,7)
								
								        'fakbareTimerTot(v) = fakbareTimerTot(v) + jobmedtimer(x,9)
								
                                
                                        'LastMd = datepart("m", jobmedtimer(x, 36), 2,2)



    end sub

 %>