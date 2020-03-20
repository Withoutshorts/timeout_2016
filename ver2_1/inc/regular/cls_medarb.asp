 
 <%
 
  

 public meNavn, meNr, meInit, meTxt, meEmail, meType, meAnsatDato, meOpsagtdato, meforecaststamp, meBrugergruppe, meVisskiftversion, meMansat, timer_ststop, create_newemployee, meCPR, meCreate_newemployee, meMed_lincensindehaver, meCal
     'Public Shared Function meStamdata(medid)
    Function meStamdata(medid)  

       if len(trim(medid)) <> 0 then
       medid = medid
       else
       medid = 0
       end if

       meVisskiftversion = 0
       
        strSQLmnavn = "SELECT mnavn, init, mnr, email, medarbejdertype, ansatdato, opsagtdato, forecaststamp, brugergruppe, visskiftversion, mansat, timer_ststop, create_newemployee, mcpr, med_lincensindehaver, med_cal "_
        &" FROM medarbejdere WHERE mid = "& medid

	    oRec3.open strSQLmnavn, oConn, 3
	    if not oRec3.EOF then
    	
	    meNavn = oRec3("mnavn")
	    meNr = oRec3("mnr")
	    meInit = oRec3("init")
    	meEmail = oRec3("email")
    	meType = oRec3("medarbejdertype")
    	meAnsatDato = oRec3("ansatdato")
    	meOpsagtdato = oRec3("opsagtdato")
        meBrugergruppe = oRec3("brugergruppe")
        meVisskiftversion = oRec3("visskiftversion")'
        meMansat = oRec3("mansat")
        meCPR = oRec3("mcpr")
    	
	    meTxt = meNavn ' & " ("& meNr & ")"
    	
	    if len(trim(meInit)) <> 0 then
	    meTxt = meTxt &" ["& meInit & "]"
	    end if
    	
        if IsNull(oRec3("forecaststamp")) <> true then
        meforecaststamp = formatdatetime(oRec3("forecaststamp"), 1)
        else
        meforecaststamp = ""
        end if

        timer_ststop = oRec3("timer_ststop")

        create_newemployee = oRec3("create_newemployee")
        meCreate_newemployee = create_newemployee

        meMed_lincensindehaver = oRec3("med_lincensindehaver")
        meCal = oRec3("med_cal")

	    end if
	    oRec3.close
       
   end function
   


'****** Medarbejder timepris og kostpris ********************
public strMnavn, dblkostpris, tprisGen, valutaGen, mkostpristarif_A, mkostpristarif_B, mkostpristarif_C, mkostpristarif_D, intKpValuta 
function mNavnogKostpris(strMnr)
'** Henter navn og kostpris ***'

SQLmedtpris = "SELECT medarbejdertype, timepris, tp0_valuta, kostpris, mnavn, "_
&" kostpristarif_A, kostpristarif_B, kostpristarif_C, kostpristarif_D, kp1_valuta FROM medarbejdere, medarbejdertyper "_
&" WHERE Mid = "& strMnr &" AND medarbejdertyper.id = medarbejdertype"

'Response.Write SQLmedtpris
'Response.flush
oRec.Open SQLmedtpris, oConn, 3

		if Not oRec.EOF then
		 	
		 	if oRec("kostpris") <> "" then
			dblkostpris = oRec("kostpris")
			else
			dblkostpris = 0
			end if
		
		strMnavn = oRec("mnavn")
		tprisGen = oRec("timepris")
		valutaGen = oRec("tp0_valuta")

        mkostpristarif_A = oRec("kostpristarif_A")
        mkostpristarif_B = oRec("kostpristarif_B")
        mkostpristarif_C = oRec("kostpristarif_C") 
        mkostpristarif_D = oRec("kostpristarif_D")

        intKpValuta = oRec("kp1_valuta")
		
		end if

oRec.close

'** Slut timepris **


end function




   public antalmedarb
   function antalmedarb_fn(io)

   '1 = aktiv
   '2 = de-aktiv
   '3 = passiv
   '4 = Gæst

   if io = 2 then
   medStatKri = "mansat <> 2 AND mansat <> 4" '** Aktiv + passive, minus gæster
   else
   medStatKri = "mansat <> 0 AND mansat <> 4" 'Alle
   end if


            antalmedarb = 0
	        strSQL = "SELECT count(Mid) AS antusr FROM medarbejdere WHERE " & medStatKri
			oRec2.Open strSQL, oConn, 3
			if not oRec2.EOF then
			antalmedarb = oRec2("antusr")
			end if
			oRec2.close

   end function


     antMtypT = 150 'antalmtyper 
     '************* Medarbejdertyper i medarbejderhovedgrupper ****'
    public mtypgrpids, mtypgrpnavn, mtypgrpid, medarbimType, medarbSQlKriMtypGrp, mtypid, mtypnavn, kunMtypgrp, kunMtypgrpNavn, thisMtypGrpBelIg, thisMtypTimerBelob
    redim mtypgrpids(antMtypT), mtypgrpnavn(antMtypT), mtypgrpid(antMtypT), medarbimType(antMtypT), mtypid(antMtypT), mtypnavn(antMtypT)
    redim kunMtypgrp(10), kunMtypgrpNavn(10)
    redim thisMtypGrpBelIg(50), thisMtypTimerBelob(50)

    
     
   function mtyperIGrp_fn(vlgtmtypgrp,visning) 
     
     if visning = 0 then 'budget:0 / Realiseret:1
     sostergrpSQL = "mt.sostergp = 0" 
     else
     sostergrpSQL = "mt.sostergp <> -1"
     end if

     'Response.write "HER: "& vlgtmtypgrp
     'vlgtmtypgrp = 0
     if vlgtmtypgrp <> "0" then
    
        vlgtmtypgrpIds = split(vlgtmtypgrp, ", ")
        
        vlgtmtypgrpSQLkri = " (mtg.id = 0 "
        for v = 0 to UBOUND(vlgtmtypgrpIds) 
        vlgtmtypgrpSQLkri = vlgtmtypgrpSQLkri & " OR mtg.id = "& trim(vlgtmtypgrpIds(v))
        next
        
     vlgtmtypgrpSQLkri = vlgtmtypgrpSQLkri & ")"   

     'Response.write vlgtmtypgrpSQLkri
     'Response.end 

     else
     vlgtmtypgrpSQLkri = "mtg.id > 0" 'så vi undgår -1 gruppen, dvs dem der i søstergruppe er sat til ikke at skulle medregnes overhoved.
     end if

    'Response.write "THISFILE: "& thisfile & " FM_medarb. "& request("FM_medarb") & "<br>"
    'if thisfile = "joblog_timetotaler" AND 

    strSQLmtyp = "SELECT mtg.id AS mtypGrpId, mt.id AS mtypId, mtg.navn AS mtypgrpnavn, mt.type AS mtypenavn FROM medarbtyper_grp AS mtg "_
    &" LEFT JOIN medarbejdertyper AS mt ON (mt.mgruppe = mtg.id AND "& sostergrpSQL &")"_
    &" WHERE "& vlgtmtypgrpSQLkri &" AND ("& sostergrpSQL &") ORDER BY mtg.id, mt.id "
    
    'if session("mid") = 1 then
    'Response.write "xx" &  strSQLmtyp
    'Response.flush
    'end if
    
     tcnt = 0
     kmyp = 0


     mtypgrpidStr = "#0#"
     mtypgrpids(0) =  "typeid = 0 "
     LastMtypGrpId = 0
     LastMtypId = 0
     thisMtypid = 0
     medarbSQlKriMtypGrp = "(m.mid = 0 "
     medarbimType(tcnt) = " tmnr = 0 "
     
    
     oRec6.open strSQLmtyp, oConn, 3
     while not oRec6.EOF
    
    if isNull(oRec6("mtypId")) <> true then

     tcnt = tcnt + 1

            'if oRec6("mtypGrpId") <> LastMtypGrpId then
             if oRec6("mtypId") <> LastMtypId then

             'tcnt = tcnt + 1
             mtypgrpids(tcnt) = "typeid = 0 "
             medarbimType(tcnt) = " tmnr = 0 "
     
             end if
    
    
    mtypgrpids(tcnt) = mtypgrpids(tcnt) & " OR typeid = "& oRec6("mtypId")
    mtypgrpnavn(tcnt) = oRec6("mtypgrpnavn")
    mtypgrpid(tcnt) = oRec6("mtypGrpId")
    mtypid(tcnt) = oRec6("mtypId")
    mtypnavn(tcnt) = oRec6("mtypenavn")
   
    if instr(mtypgrpidStr, "#"&oRec6("mtypGrpId")&"#") = 0 then
    kmyp = kmyp + 1
    kunMtypgrp(kmyp) = oRec6("mtypGrpId")
    kunMtypgrpNavn(kmyp) = oRec6("mtypgrpnavn")
    mtypgrpidStr = mtypgrpidStr & ",#"& oRec6("mtypGrpId")&"#"
    end if

    if isNull(oRec6("mtypId")) <> true then
    thisMtypid = oRec6("mtypId")
    else
    thisMtypid = 0
    end if
            
            '** Medarbjedere af typen
            '*** Kan optimeres da medarbejerrtype_grp er kommet med på medarbejder niveau
            strSQLmedityp = "SELECT m.mid FROM medarbejdere AS m WHERE m.medarbejdertype = "& thisMtypid & " AND m.medarbejdertype_grp <> -1" 
            oRec4.open strSQLmedityp, oConn, 3
            while not oRec4.EOF

             medarbimType(tcnt) =  medarbimType(tcnt) & " OR tmnr = "& oRec4("mid")
             medarbSQlKriMtypGrp = medarbSQlKriMtypGrp & " OR m.mid = "& oRec4("mid")



             oRec4.movenext
             wend 
             oRec4.close

       
     
     
     end if 'mtypeID Null

     LastMtypGrpId = oRec6("mtypGrpId") 
     LastMtypId = oRec6("mtypId")
     

     oRec6.movenext
    Wend
    oRec6.close 

     medarbSQlKriMtypGrp = medarbSQlKriMtypGrp & ")"

    end function



     public mtypeids, mtypenavne, mtypesostergp, mtypeGp
     redim mtypeids(1), mtypenavne(1), mtypesostergp(1), mtypeGp(1)
     function fn_medarbtyper()

     t = 0
     '** kun dem der skal vises under stat. SosterGp > 0 tilføejs til dens søster
     strSQLmtyp = "SELECT id, type, sostergp, mgruppe FROM medarbejdertyper WHERE id <> 0 AND sostergp = 0 ORDER BY mtsortorder, type"

     oRec4.open strSQLmtyp, oConn, 3
     while not oRec4.EOF
     
     redim preserve mtypeids(t), mtypenavne(t), mtypesostergp(t), mtypeGp(t)
     
    
     mtypeids(t) = oRec4("id")
     mtypenavne(t) = oRec4("type")
     mtypesostergp(t) = ""
     mtypeGp(t) = oRec4("mgruppe") 
                
        strSQLSostyp = "SELECT id, type, sostergp FROM medarbejdertyper WHERE sostergp = "& oRec4("id") &" ORDER BY id"
        oRec5.open strSQLSostyp, oConn, 3
        while not oRec5.EOF
        
              mtypesostergp(t) = mtypesostergp(t) & " OR mtype = "& oRec5("id")

        oRec5.movenext
        wend
        oRec5.close
   
   
   
     
     t = t + 1
     oRec4.movenext
     wend 
     oRec4.close


     end function




     sub medarb_vaelgandre


                    if cint(visAlleMedarb_pas) = 1 then
                        strSQLmpasKri  = " mansat <> 2 AND mansat <> 4 "
                    else
                        strSQLmpasKri  = " mansat = 1 "
                    end if

     
					strSQL = "SELECT Mid, Mnavn, Mnr, Brugergruppe, init, mansat FROM medarbejdere WHERE "& strSQLmpasKri &" "& strSQLmids &" GROUP BY mid ORDER BY Mnavn"
					
                    if thisfile = "ugeseddel_2011.asp" OR thisfile = "logindhist_2011.asp" then
                    mSelWdth = "100%"
                    mSelcls = "form-control input-small"
                    else
                    mSelWdth = "250"
                    mSelcls = ""
					end if%>

                    
					<select name="usemrn" id="usemrn" style="width:<%=mSelWdth%>px;" class="<%=mSelcls %>">
                        <!-- onchange="submit(); -->
					<%
					
					oRec.open strSQL, oConn, 3
					while not oRec.EOF 
					
					if cdbl(oRec("mid")) = cdbl(usemrn) then
					rchk = "SELECTED"
					else
					rchk = ""
					end if
                    
                        if len(trim(oRec("init"))) <> 0 then
                        medTxt = oRec("mnavn") & " ["& oRec("init") &"]"
                        else
                        medTxt = oRec("mnavn") 
                        end if

                        if oRec("mansat") = 3 then
                        mStatusTxt = " - Passiv"
                        else
                        mStatusTxt = ""
                        end if
                    %>
					<option value="<%=oRec("Mid")%>" <%=rchk%>><%=medTxt & " " & mStatusTxt%></option>
					<%
					
					
					oRec.movenext
					wend
					oRec.close
				%></select>


    <%

     end sub


    public mtypNoflex, mtypMtypeId 
    function mtyp_fn(mid)

        mtypMtypeId = 0
        mtypNoflex = 0
        strSQLmtyp = "SELECT m.medarbejdertype, mt.noflex, mt.id AS medarbtype FROM medarbejdere AS m "_
        &" LEFT JOIN medarbejdertyper AS mt ON (mt.id = m.medarbejdertype) WHERE m.mid = "& mid &" ORDER BY id"
        
        'response.write strSQLmtyp
        'response.flush
        
        oRec5.open strSQLmtyp, oConn, 3
        while not oRec5.EOF
        
         mtypNoflex = oRec5("noflex")
         mtypMtypeId = oRec5("medarbtype")

        oRec5.movenext
        wend
        oRec5.close


    end function






public lastLoginDateFm, mstatus
sub mstatus_lastlogin                                                
      
        
        
        
        
            select case oRec("mansat") 
            case "1"
            mstatus = medarb_txt_029
            case "2"
            mstatus = medarb_txt_030
            case "3"
            mstatus = medarb_txt_031
            case "4"
            mstatus = medarb_txt_149
            case else
            mstatus = ""
            end select

                                            
                lastLoginDateFm = ""
                if len(trim(oRec("lastlogin"))) <> 0 AND isNull(oRec("lastlogin")) <> true then
                lastLoginDateFm = oRec("lastlogin")
                                                            
                lastLoginDateFm = replace(lastLoginDateFm, "Jan.", "-01-") 
                lastLoginDateFm = replace(lastLoginDateFm, "Feb.", "-02-")
                lastLoginDateFm = replace(lastLoginDateFm, "Mar.", "-03-")
                lastLoginDateFm = replace(lastLoginDateFm, "Apr.", "-04-")
                lastLoginDateFm = replace(lastLoginDateFm, "Maj.", "-05-")
                lastLoginDateFm = replace(lastLoginDateFm, "Jun.", "-06-")
                lastLoginDateFm = replace(lastLoginDateFm, "Jul.", "-07-")
                lastLoginDateFm = replace(lastLoginDateFm, "Aug.", "-08-")
                lastLoginDateFm = replace(lastLoginDateFm, "Sep.", "-09-")
                lastLoginDateFm = replace(lastLoginDateFm, "Okt.", "-10-")
                lastLoginDateFm = replace(lastLoginDateFm, "Nov.", "-11-")
                lastLoginDateFm = replace(lastLoginDateFm, "Dec.", "-12-")
                lastLoginDateFm = replace(lastLoginDateFm, " ", "") 


                if mid(lastLoginDateFm, 2,1) = "-" then
                lastLoginDateFm = "0" & lastLoginDateFm
                end if

                else
                lastLoginDateFm = "01-01-2002"
                end if

                lastLoginDateFm = cDate(lastLoginDateFm)

                'lastLoginDateFm = lastLoginDateFm & ": " & mid(lastLoginDateFm, 2,1)
                


end sub




Public alleMedIJurEnhedSQL
function alleMedIJurEnhed(med_lincensindehaver, mansat)

        if cint(mansat) = 1 then 'ALLE
        strSQLmansat = " AND mansat <> - 1 AND mansat <> 4"
        else
        strSQLmansat = " AND (mansat = 1 OR mansat = 3)" 'Aktive og passvie
        end if

        alleMedIJurEnhedSQL = " AND (tmnr = 0 "

        strSQLalleMedIJurEnhedSQL = "SELECT mid FROM medarbejdere WHERE med_lincensindehaver = "& med_lincensindehaver & " "& strSQLmansat

        oRec5.open strSQLalleMedIJurEnhedSQL, oConn, 3
        while not oRec5.EOF
        
         alleMedIJurEnhedSQL = alleMedIJurEnhedSQL & " OR tmnr = "& oRec5("mid")
      
        oRec5.movenext
        wend
        oRec5.close

        alleMedIJurEnhedSQL = alleMedIJurEnhedSQL & ")"

                   

end function




function selectMedarbPrProgrp(showAll, prgids, sel_name, onchangesubmit, io)

        'if len(trim(prgids)) <> 0 then
        'prgids = prgids
        'else
        'prgids = 0
        'end if

        select case cint(io)
            case 0

                projektgruppeKri = ""

            case 1

                if prgids <> "alle" AND prgids <> "ingen" then
                    prgids_aar = split(prgids, ",")
            
                    p = 0
                    for p = 0 TO UBOUND(prgids_aar)

                        'response.Write "<br> p her " & prgids(p) 
                        if p = 0 then
                           prgidKri = prgids_aar(p)
                        else
                           prgidKri = prgidKri &" OR pg.id = "& prgids_aar(p)
                        end if

                    next

                    projektgruppeKri = " AND pg.id = "& prgidKri
                    'response.Write " projektgruppeKri " & projektgruppeKri
                else
                    'if prgids <> "ingen" then 'ADMIN
                    projektgruppeKri = ""
                    'else
                    'projektgruppeKri = " mid = " & session("mid")
                    'end if
                end if

        end select

        if cint(onchangesubmit) = 1 then
            STRonchangesubmit = "onchange='submit();'"
        else
            STRonchangesubmit = ""
        end if

        if thisfile = "stade" AND lto = "epi2017" then
        orgvirKri = " AND (orgvir = 1)" 'OR orgvir = 2
        else
        orgvirKri = ""
        end if

                                            if prgids <> "ingen" then 'ADMIN / Teamleder
                                            strSQLm = "SELECT pg.navn AS pgnavn, pg.id AS pgid, mnavn, mid, init FROM projektgrupper pg " 
                                            strSQLm = strSQLm & " LEFT JOIN progrupperelationer pgl ON (pgl.projektgruppeid = pg.id)"
                                            strSQLm = strSQLm & " LEFT JOIN medarbejdere m ON (m.mid = pgl.medarbejderid)"
                                            strSQLm = strSQLm & " WHERE mansat = 1 "& orgvirKri &" AND pg.id <> 10 "& projektgruppeKri &" ORDER BY pg.navn, mnavn"
                                            else                                    
                                            strSQLm = "SELECT mnavn, mid, init FROM medarbejdere m WHERE mid = "& session("mid") 
                                            end if
                                        
                                            'if session("mid") = 1 then
                                            'response.Write strSQLm & "<br>prgids: " & prgids & "#"
                                            'response.flush
                                            'end if


                                            if prgids <> "ingen" then
                                            %>
                                           

                                                <%if cint(showAll) = 1 then %>
                                                <select name="<%=sel_name %>" class="form-control input-small" <%=STRonchangesubmit %>>
                                                <option value="0">Vis alle</option>
                                                <% 
                                                 else
                                                    %>
                                                      <select name="<%=sel_name %>" class="form-control input-small" <%=STRonchangesubmit %>>

                                                    <%
                                                 end if

                                            else
                                             %>
                                                      <select name="<%=sel_name %>" class="form-control input-small" <%=STRonchangesubmit %>>

                                                    <%

                                            end if

                                            oRec.open strSQLm, oConn, 3
                                            While not oRec.EOF 
                                            

                                            if prgids <> "ingen" then

                                            if lastpgid <> oRec("pgid") then %>
                                                 <option DISABLED></option>
                                                <option DISABLED><%=oRec("pgnavn") %></option>
                                            <%end if

                                            end if

                                            if len(trim(oRec("init"))) <> 0 then
                                            init = "["& oRec("init") &"]"
                                            else
                                            init = ""
                                            end if

                                            if cdbl(usemrn) = cdbl(oRec("mid")) then
                                            mSEL = "SELECTED"
                                            else
                                            mSEL = ""
                                            end if
                                            %>
                                          <option value="<%=oRec("mid") %>" <%=mSEL %>><%=oRec("mnavn") & " "& init %></option>
                                           <%

                                            
                                           if prgids <> "ingen" then
                                           lastpgid = oRec("pgid") 
                                           end if



                                            oRec.movenext
                                            Wend
                                            oRec.close
                                            
                                         %>

                                    


                                                 </select>

<%
end function

   %>