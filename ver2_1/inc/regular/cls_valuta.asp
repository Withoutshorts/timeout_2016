 
 <%

function opdaterValutaAktiveJob(lto, io, jobids, valutaid, intKurs)

                    if cint(io) = 1 then 'Henter aktuelle kurs 

                        intKurs = 100
                        strSQLv = "SELECT kurs FROM valutaer WHERE id = " & valutaid
	                    oRec3.open strSQLv, oConn, 3
	
	                    if not oRec3.EOF then
	
	                                                         
	                    intKurs = oRec3("kurs")/100
	                                                            
	
	                    end if
	                    oRec3.close

                        'intKurs = replace(intKurs, ".", "")
                        intKurs = replace(intKurs, ",", ".")

                    else

                        intKurs = intKurs

                    end if
                

                    '*** Henter aktive job i valuta og opdaterer timer
                    '*** SPECIAL LTO SETTINGS *****'
                    select case lto
                    case "nt", "intranet - local"
                    
                    if jobids <> 0 then
                    jobidSQL = " AND id = "& jobids
                    else
                    jobidSQL = ""
                    end if
     
                    strSQLopdajobtimer = "SELECT id, jobnr, orderqty, tax_pc, cost_price_pc, sales_price_pc, freight_pc, sales_price_pc_valuta, "_
                    &" cost_price_pc_valuta, comm_pc, fastpris FROM job WHERE jobstatus = 1 AND cost_price_pc_valuta = "& valutaid & jobidSQL
                    
     
                    case else
                    strSQLopdajobtimer = "SELECT id, jobnr FROM job WHERE jobstatus = 1 AND valuta = "& valutaid &""
                    end select
                   
                    
                    'response.write strSQLopdajobtimer & "<br>"
                 
                        j = 0
                    	oRec.open strSQLopdajobtimer, oConn, 3
        				while not oRec.EOF
        				            
                


                                    'if cint(j) = 0 then

                        
                                        '*** SPECIAL LTO SETTINGS *****'
                                        select case lto
                                        case "nt", "intranet - local"


                                        if oRec("tax_pc") <> 0 then
                                        tax = 1 + (oRec("tax_pc")/100)
                                        else
                                        tax = 1
                                        end if

                                        orderqty = oRec("orderqty")
                                        
                                        freight_pc = oRec("freight_pc")
                                            
                                        cost_price_pc = oRec("cost_price_pc")
                                        
                                        sales_price_pc = oRec("sales_price_pc")

                                        comm_pc = oRec("comm_pc")
                                        
                                        intKursUse = intKurs/100
                                        
                                        'response.write "<br><br>((qty: "&  orderqty &" * cost "& cost_price_pc &" * "& tax &") + ("& freight_pc &" * "& orderqty & ")) * kurs: "& (intKursUse)  & " # comm_pc: "& comm_pc
                                        'response.write "<br>sales_price_pc: "& sales_price_pc
                                       
                                        'if oRec("fastpris") = 2 then 'COMI

                                        jo_udgifter_intern_ny = (orderqty * cost_price_pc * tax)
                                        'response.write "jo_udgifter_intern_ny 1: "& jo_udgifter_intern_ny & "<br>"
                                        jo_udgifter_intern_ny = jo_udgifter_intern_ny  + (freight_pc * orderqty)
                                        'response.write "jo_udgifter_intern_ny 2: "& jo_udgifter_intern_ny & "<br>"
                                        jo_udgifter_intern_ny = jo_udgifter_intern_ny  * (intKursUse)
                                        'jo_udgifter_intern_ny = formatnumber(jo_udgifter_intern_ny, 2)
                                        'response.write "jo_udgifter_intern_ny 3: "& jo_udgifter_intern_ny & "<br>"

                                        jo_udgifter_intern_ny = replace(jo_udgifter_intern_ny, ".", "")
                                        jo_udgifter_intern_ny = replace(jo_udgifter_intern_ny, ",", ".")
                                        'response.write "jo_udgifter_intern_ny 4: "& jo_udgifter_intern_ny & "<br>"

                                        'else 'salesprice

                                     

                                        'end if                        

                                        jo_bruttooms_ny = (orderqty * sales_price_pc)

                                        if cint(oRec("sales_price_pc_valuta")) = cint(oRec("cost_price_pc_valuta")) then 'salgspris har samme valuta som kost
                                        
                                        jo_bruttooms_ny = jo_bruttooms_ny * (intKursUse)
                                        
                                        else
                                                            
                                                                intSalgsPcKurs = 100
                                                                strSQLv = "SELECT kurs FROM valutaer WHERE id = " & oRec("sales_price_pc_valuta")
	                                                            oRec3.open strSQLv, oConn, 3
	
	                                                            if not oRec3.EOF then
	
	                                                         
	                                                            intSalgsPcKurs = oRec3("kurs")/100
	                                                            
	
	                                                            end if
	                                                            oRec3.close

                                        jo_bruttooms_ny = jo_bruttooms_ny * (intSalgsPcKurs)                                    
                                        end if
                                        
                                        'jo_bruttooms_ny = formatnumber(jo_bruttooms_ny, 2)
                                        
                                        jo_bruttooms_ny = replace(jo_bruttooms_ny, ".", "")
                                        jo_bruttooms_ny = replace(jo_bruttooms_ny, ",", ".")

                                        strSQLopdajob = "UPDATE job SET jo_udgifter_intern = "& jo_udgifter_intern_ny &", jo_bruttooms = "& jo_bruttooms_ny &" WHERE id = "& oRec("id")
                    
                                      
                                        'response.write "<br>"& strSQLopdajob
                                        'response.flush
                                        'response.end
                                        oConn.execute(strSQLopdajob)

                                        case else

                                         strSQLopdajob = "UPDATE job SET jo_valuta_kurs = "& intKurs &" WHERE jo_valuta = "& valutaid
                    
                                      
                                        'response.write "<br>"& strSQLopdajob
                                        'response.end
                                        oConn.execute(strSQLopdajob)

                                        end select
                    


                                    'end if


                                    select case lto
                                    case "nt", "intranet - local"


                                    case else
                                    '*** Opdaterer alle valutaer på timeregistreringer 
                                    strSQLopdtimer = "UPDATE timer SET kurs = "& intKurs &" WHERE tjobnr = '"& oRec("jobnr") &"' AND valuta = "& valutaid 

                                    'response.write "<br>" & strSQLopdtimer
                                    oConn.execute(strSQLopdtimer)            
                                    end select

                                    
                        j = j + 1
        				oRec.movenext
        				wend
        				oRec.close


end function



public valutaKode_CCC, valutaKode_CCC_f8, valutaKurs_CCC 
function valutakode_fn(valid)

     valutaKode_CCC = ""
      strSQL3 = "SELECT id, valutakode, grundvaluta, kurs FROM valutaer WHERE id = "& valid
    		
    		
		    oRec3.open strSQL3, oConn, 3 
		    if not oRec3.EOF then 

            valutaKode_CCC = oRec3("valutaKode")
            valutaKode_CCC_f8 = "<span style='font-size:7px;'>"& oRec3("valutaKode") &"</span>"
            valutaKurs_CCC = oRec3("kurs")
            
            end if
            oRec3.close


end function


'** Valutakoder Liste simpel
function valutaList(valuta, felt)

     %>
      <select name="<%=felt %>" id="<%=felt%>" class="s_valuta form-control input-small" style="width:60px;">
            <%strSQL = "SELECT id, valutakode FROM valutaer WHERE id <> 0 ORDER BY valutakode"
        	oRec6.open strSQL, oConn, 3
        	while not oRec6.EOF

                if cint(valuta) = cint(oRec6("id")) then
                valSel = "SELECTED"
                else
                valSel = ""
                end if
        				
            %><option value="<%=oRec6("id")%>" <%=valSel %>><%=oRec6("valutakode") %></option><%
        				       
            oRec6.movenext
        	wend
        	oRec6.close %>
                                


            </select>

    <%
end function

'** Valutakoder Liste med kurs
 function valutakoder(i, valuta, layout)
	
     select case cint(layout) 
     case 0
     cssThis = "width:55px; font-size:9px;" 
     cls = ""
     case 1
     cssThis = ""
     cls = "form-control input-small"
     case else
     cssThis = ""
     cls = ""
     end select

     if instr(i, "freight_price_pc_valuta") <> 0 then
     dsabl = "disabled"
     else
     dsabl = ""
     end if
     
     %>
	<select name="FM_valuta_<%=i %>" id="FM_valuta_<%=i %>" style="<%=cssThis%>" class="<%=cls%>" <%=dsabl %>>
		    
		    <%
		    strSQL3 = "SELECT id, valutakode, grundvaluta, kurs FROM valutaer ORDER BY valutakode"
    		
    		
		    oRec3.open strSQL3, oConn, 3 
		    while not oRec3.EOF 
    		
		    if cint(valuta) = cint(oRec3("id")) then
		    valGrpCHK = "SELECTED"

               'KURS VED OPRETTELSE
               if thisfile = "job_nt.asp" AND func = "red" then

                'if sales_price_pc <> 0 AND orderqty <> 0 AND i = "sales_price_pc_valuta" then
                'oprKurs = formatnumber((bruttooms / (orderqty * sales_price_pc)), 2)
                'else
                'oprKurs = 0
                'end if 

                oprKurs = 0

                if i = "cost_price_pc_valuta" OR i = "freight_price_pc_valuta" then
                oprKurs = cost_price_kurs_used

                        if cdbl(oprKurs) = 1 then 'gl ordrer
                        'oprKurs = formatnumber((bruttooms / (orderqty * sales_price_pc)), 2)
                        

                        if cost_price_pc <> 0 AND orderqty <> 0 then 'AND i = "sales_price_pc_valuta"
                        oprKurs = formatnumber((bruttooms / (orderqty * cost_price_pc)), 2)
                        else
                        oprKurs = 0
                        end if 

                        end if

                end if

                if i = "sales_price_pc_valuta" then
                oprKurs = sales_price_kurs_used

                        if cdbl(oprKurs) = 1 then 'gl ordrer
                        'oprKurs = formatnumber((bruttooms / (orderqty * sales_price_pc)), 2)
                        if sales_price_pc <> 0 AND orderqty <> 0 then 'AND i = "sales_price_pc_valuta"
                        oprKurs = formatnumber((bruttooms / (orderqty * sales_price_pc)), 2)
                        else
                        oprKurs = 0
                        end if 

                        end if

                end if

             

               end if

		    else
		    valGrpCHK = ""
		    end if
		    
                if cint(layout) = 0 then
                kursTxt = ""
                else
                    if oRec3("kurs") <> 100 then
                    kursTxt = "("& formatnumber(oRec3("kurs")/100, 2) &")"
                    else
                    kursTxt = ""
                    end if
                end if

               
               'KURS VED OPRETTELSE
               if thisfile = "job_nt.asp" AND func = "red" AND valGrpCHK = "SELECTED" then

               if cdbl(oprKurs) <> 0 then
               kursTxt = " - "& oprKurs & " "& kursTxt  
               else
               kursTxt = " - "& kursTxt
               end if

               end if
		   
		    %>
		    <option value="<%=oRec3("id")%>" <%=valGrpCHK %>><%=oRec3("valutakode") &" "& kursTxt%></option>
		    <%
		    oRec3.movenext
		    wend
		    oRec3.close %>
		    </select>
	<%
	end function
	
	
	public dblKurs
	function valutaKurs(intValuta)
	    '**** Finder aktuel kurs ***'
       dblKurs = 100

        if cint(intValuta) <> 0 then
        intValuta = intValuta
        else
        intValuta = 1
        end if

       strSQL = "SELECT kurs FROM valutaer WHERE id = " & intValuta
       oRec.open strSQL, oConn, 3
       if not oRec.EOF then
       dblKurs = replace(oRec("kurs"), ",", ".")
       end if 
       oRec.close
	end function

    public dblKurs_fakhist
	function valutaKurs_fakhist(intValuta)
	    '**** Finder aktuel kurs ***'
       dblKurs_fakhist = 100

       strSQLdblKurs_fakhist = "SELECT kurs FROM valutaer WHERE id = " & intValuta
       oRec6.open strSQLdblKurs_fakhist, oConn, 3
       if not oRec6.EOF then
       dblKurs_fakhist = oRec6("kurs") 'replace(oRec6("kurs"), ",", ".")
       end if 
       oRec6.close
	end function


    public dblKurs_job
	function valutaKurs_job(intValuta)
	    '**** Finder aktuel kurs ***'
       dblKurs_job = 100

        if len(trim(intValuta)) <> 0 then
        intValuta = intValuta
        else
        intValuta = 1
        end if 

       strSQLdblKurs_job = "SELECT kurs FROM valutaer WHERE id = " & intValuta
       oRec6.open strSQLdblKurs_job, oConn, 3
       if not oRec6.EOF then
       dblKurs_job = replace(oRec6("kurs"), ",", ".")
       end if 
       oRec6.close
	end function


    public dblKursOR3
	function valutaKursOR3(intValuta)
	    '**** Finder aktuel kurs ***'
       dblKursOR3 = 100
       strSQL = "SELECT kurs FROM valutaer WHERE id = " & intValuta
       oRec3.open strSQL, oConn, 3
       if not oRec3.EOF then
       dblKursOR3 = oRec3("kurs")
       end if 
       oRec3.close
	end function

    public basisValId, basisValISO, basisValKurs, basisValISO_f8
	function basisValutaFN()
	    '**** Finder basisValuta ***'
       dblKurs = 100
       basisValId = 1
        
       strSQL = "SELECT id, valutakode, kurs FROM valutaer WHERE grundvaluta = 1"
       
       'oConn.execute(strSQL) 
       oRec.open strSQL, oConn, 3
       if not oRec.EOF then
       basisValKurs = replace(oRec("kurs"), ",", ".")

       basisValId = oRec("id")
       basisValISO = oRec("valutakode")
       basisValISO_f8 = "<span style='font-size:7px;'>"& basisValISO &"</span>"

       end if 
       oRec.close
	end function
	
	
	public valBelobBeregnet
	function beregnValuta(belob,frakurs,tilkurs)
	        
        'Response.write lastFaknr & " belob: "& belob & "<br>"
        'Response.flush

        if isNull(belob) <> true AND belob <> "-" AND isNull(frakurs) <> true AND isNull(tilkurs) <> true AND tilkurs <> 0 AND len(trim(belob)) <> 0 then
        valBelobBeregnet = belob * (frakurs/tilkurs)
        else
        valBelobBeregnet = 0
        end if
    
    if len(valBelobBeregnet) <> 0 then
    valBelobBeregnet = valBelobBeregnet
    else
    valBelobBeregnet = 0
    end if
    
    
    valBelobBeregnet = formatnumber(valBelobBeregnet, 2)
    valBelobBeregnet = replace(valBelobBeregnet, ".", "")
    valBelobBeregnet = valBelobBeregnet/1
    'Response.Write valBelobBeregnet & "<br>"
    'Response.flush
	end function

    %>