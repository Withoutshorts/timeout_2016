<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->


   
   
<%'GIT 20160811 - SK
if len(session("user")) = 0 then
%>
<!--#include file="../inc/regular/header_inc.asp"-->
<% 

	errortype = 5
	call showError(errortype)
	
	else
	
	'tjek for aktive moduler
	call erSDSKaktiv()
	if cint(dsksOnOff) = 1 then
	Strsdskstat = "checked"
	end if
	call erCRMaktiv()
	if cint(crmOnOff) = 1 then
	Strcrmstat = "checked"
	end if
	call erERPaktiv()
	if cint(erpOnOff) = 1 then
	Strerpstat = "checked"
	end if

	
	
	func = request("func")
	select case func
	case "exch"
			
			exchangeserverURL = request("FM_exchange")
			exchangeserverDOM = request("FM_exdom")
			exchangeserverBruger = replace(request("FM_exbruger"), "\", "#")
			exchangeserverPW = request("FM_expw")
			
			if len(request("FM_smiley")) <> 0 then
			smiley = request("FM_smiley")
			else
			smiley = 0
			end if
			
			if len(request("FM_stempelur")) <> 0 then
			stempelur = request("FM_stempelur")
			else
			stempelur = 0
			end if
			
            
            if len(request("FM_stempelur_hidelogin")) <> 0 then
			stempelur_hidelogin = request("FM_stempelur_hidelogin")
			else
			stempelur_hidelogin = 0
			end if

        
            if len(request("FM_stempelur_igno_komkrav")) <> 0 then
			stempelur_igno_komkrav = request("FM_stempelur_igno_komkrav")
			else
			stempelur_igno_komkrav = 0
			end if

            if len(trim(request("FM_visAktlinjerSimpel"))) <> 0 then
            visAktlinjerSimpel = 1
            else
            visAktlinjerSimpel = 0
            end if

           

            if len(trim(request("FM_visAktlinjerSimpel_datoer"))) <> 0 then
            visAktlinjerSimpel_datoer = 1
            else
            visAktlinjerSimpel_datoer = 0
            end if


            if len(trim(request("FM_visAktlinjerSimpel_timebudget"))) <> 0 then
            visAktlinjerSimpel_timebudget = 1
            else
            visAktlinjerSimpel_timebudget = 0
            end if

            if len(trim(request("FM_visAktlinjerSimpel_realtimer"))) <> 0 then
            visAktlinjerSimpel_realtimer = 1
            else
            visAktlinjerSimpel_realtimer = 0
            end if

            if len(trim(request("FM_visAktlinjerSimpel_restimer"))) <> 0 then
            visAktlinjerSimpel_restimer = 1
            else
            visAktlinjerSimpel_restimer = 0
            end if


            if len(trim(request("FM_visAktlinjerSimpel_medarbtimepriser"))) <> 0 then
            visAktlinjerSimpel_medarbtimepriser = 1
            else
            visAktlinjerSimpel_medarbtimepriser = 0
            end if

            if len(trim(request("FM_visAktlinjerSimpel_medarbrealtimer"))) <> 0 then
            visAktlinjerSimpel_medarbrealtimer = 1
            else
            visAktlinjerSimpel_medarbrealtimer = 0
            end if

            if len(trim(request("FM_visAktlinjerSimpel_akttype"))) <> 0 then
            visAktlinjerSimpel_akttype = 1
            else
            visAktlinjerSimpel_akttype = 0
            end if

    


              
    
    



            
			
			'if len(request("FM_lukafm")) <> 0 then
			'lukafm = request("FM_lukafm")
			'else
			'lukafm = 0
			'end if
			
            lukafm = 1

			if len(trim(request("FM_kmDialog"))) <> 0 then  
			kmDialog = 1
			else
			kmDialog = 0
			end if
			
            
                    if len(trim(request("FM_timerround"))) <> 0 then
                    timerround = 1 
                    else
			        timerround = 0
                    end if
           
            
            
            if len(trim(request("FM_globalfaktor_ja"))) <> 0 then
                    
                    if len(trim(request("FM_globalfaktor"))) <> 0 then
                    globalfaktor = replace(request("FM_globalfaktor"), ",", ".")
                    else
			        globalfaktor = 1
                    end if


                            call erDetInt(globalfaktor)
			                if cint(isInt) > 0 then
			                errThis = 11
			                else
			                errThis = errThis
			                end if
			                isInt = 0 
                    
                    
                            if cint(errThis) = 11 then  %>
					     <!--include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 162
					call showError(errortype)
		            Response.End
			        
			        end if


                    strSQLopdAFaktor = "UPDATE aktiviteter SET faktor = " & globalfaktor & " WHERE id <> 0 AND aktstatus = 1 AND fakturerbar = 1" 
                    oConn.execute(strSQLopdAFaktor)
                    
            else

            globalfaktor = replace(request("FM_globalfaktor_old"), ",", ".")  
          
            end if


			if len(request("FM_autogk")) <> 0 then
			autogk = request("FM_autogk")
			else
			autogk = 0
			end if
			
			if len(request("FM_autogktimer")) <> 0 then
			autogktimer = request("FM_autogktimer")
			else
			autogktimer = 0
			end if
			
			if len(trim(request("budgetakt"))) <> 0 then
			budgetakt = request("budgetakt")
			else
			budgetakt = 0
			end if 
             

			if len(request("FM_autolukvdato")) <> 0 then
			autolukvdato = 1
			else
			autolukvdato = 0
			end if 
			
			if len(request("FM_autolukvdatodato")) <> 0 then
			autolukvdatodato = request("FM_autolukvdatodato")
			else
			autolukvdatodato = 28
			end if 
				
			
			
			'** Normal �bningstider ***
			if request("FM_brug_abningstid") = 1 then
			
			
			
			for x = 1 to 7
			
			select case x
			case 1
			    thisDay_t = request("FM_abn_man_t")
			    thisDay_m = request("FM_abn_man_m")
			    thisDay_t2 = request("FM_abn_man_t2")
			    thisDay_m2 = request("FM_abn_man_m2")
			case 2
			    thisDay_t = request("FM_abn_tir_t")
			    thisDay_m = request("FM_abn_tir_m")
			    thisDay_t2 = request("FM_abn_tir_t2")
			    thisDay_m2 = request("FM_abn_tir_m2")
			case 3
			    thisDay_t = request("FM_abn_ons_t")
			    thisDay_m = request("FM_abn_ons_m")
			    thisDay_t2 = request("FM_abn_ons_t2")
			    thisDay_m2 = request("FM_abn_ons_m2")
			case 4
			    thisDay_t = request("FM_abn_tor_t")
			    thisDay_m = request("FM_abn_tor_m")
			    thisDay_t2 = request("FM_abn_tor_t2")
			    thisDay_m2 = request("FM_abn_tor_m2")
			case 5
			    thisDay_t = request("FM_abn_fre_t")
			    thisDay_m = request("FM_abn_fre_m")
			    thisDay_t2 = request("FM_abn_fre_t2")
			    thisDay_m2 = request("FM_abn_fre_m2")
			 case 6
			    thisDay_t = request("FM_abn_lor_t")
			    thisDay_m = request("FM_abn_lor_m")
			    thisDay_t2 = request("FM_abn_lor_t2")
			    thisDay_m2 = request("FM_abn_lor_m2")
			 case 7
			    thisDay_t = request("FM_abn_son_t")
			    thisDay_m = request("FM_abn_son_m")
			    thisDay_t2 = request("FM_abn_son_t2")
			    thisDay_m2 = request("FM_abn_son_m2")
			end select
			        
			    %>
			    <!-- Validering -->
			    <%
			    if len(thisDay_t) <> 0 AND len(thisDay_m) <> 0_
			    AND len(thisDay_t2) <> 0 AND len(thisDay_m2) <> 0 then
			     
			        
			        
			        %><!-- Er det et gyldigt tidsformat --><%       
			        errThis = 0
			       
			        if IsDate(thisDay_t &":"& thisDay_m &":00") then
			        
			        else
			        errThis = 1
			        end if
			        
			        if IsDate(thisDay_t2 &":"& thisDay_m2 &":00") then
			        else
			        errThis = 1
			        end if
			        
			        
			        if cint(errThis) = 0 then
			        
			            if cdate((day(now) & "/" & month(now) & " / " & year(now) &" "& thisDay_t &":"& thisDay_m &":00")) > cdate(day(now) & "/" & month(now) & " / " & year(now)  &" "& thisDay_t2 &":"& thisDay_m2 &":00") then
			            errThis = 1
			            end if
    			        
    			        
			               if cint(errThis) = 0 then
					                
					            
					
                                select case x
			                    case 1
			                    normtid_st_man = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_man = thisDay_t2 &":"& thisDay_m2 &":00"
			                    case 2
			                    normtid_st_tir = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_tir = thisDay_t2 &":"& thisDay_m2 &":00"
			                    case 3
			                    normtid_st_ons = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_ons = thisDay_t2 &":"& thisDay_m2 &":00"
			                    case 4
			                    normtid_st_tor = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_tor = thisDay_t2 &":"& thisDay_m2 &":00"
			                    case 5
			                    normtid_st_fre = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_fre = thisDay_t2 &":"& thisDay_m2 &":00"
			                    case 6
			                    normtid_st_lor = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_lor = thisDay_t2 &":"& thisDay_m2 &":00"
			                    case 7
			                    normtid_st_son = thisDay_t &":"& thisDay_m &":00"
			                    normtid_sl_son = thisDay_t2 &":"& thisDay_m2 &":00"
                			    
			                    end select
			            
			                
			                 else
			             
			          %>
					  <!--include file="../inc/regular/header_lysblaa_inc.asp"-->
					  <!--include file="../inc/regular/topmenu_inc.asp"-->
			          <%
					    errortype = 79
					    call showError(errortype)
					    Response.End
					  end if
			                    
			                    
			          else
			             
			          %>
					  <!--include file="../inc/regular/header_lysblaa_inc.asp"-->
					  <!--include file="../inc/regular/topmenu_inc.asp"-->
			          <%
					    errortype = 78
					    call showError(errortype)
					    Response.End
					  end if
			    
			    else
			            %>
					     <!--include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 78
					call showError(errortype)
					isInt = 0
			        Response.End
			    end if
			
			
			next 
		            
		            
		    strSQLat = ", brugabningstid = 1, normtid_st_man = '"& normtid_st_man &"', "_
			&" normtid_sl_man = '"& normtid_sl_man &"', "_
			&" normtid_st_tir = '"& normtid_st_tir &"', "_
			&" normtid_sl_tir = '"& normtid_sl_tir &"', "_
			&" normtid_st_ons = '"& normtid_st_ons &"', "_
			&" normtid_sl_ons = '"& normtid_sl_ons &"', "_
			&" normtid_st_tor = '"& normtid_st_tor &"', "_
			&" normtid_sl_tor = '"& normtid_sl_tor &"', "_
			&" normtid_st_fre = '"& normtid_st_fre &"', "_
			&" normtid_sl_fre = '"& normtid_sl_fre &"', "_
			&" normtid_st_lor = '"& normtid_st_lor &"', "_
			&" normtid_sl_lor = '"& normtid_sl_lor &"', "_
			&" normtid_st_son = '"& normtid_st_son &"', "_
			&" normtid_sl_son = '"& normtid_sl_son &"' "
		            
		            
		        
            else 'Arb tid sl�et til
			
			    strSQLat = ", brugabningstid = 0"
			end if
			
			
			'*** Ignorer Stempelur Periode ***
			ignorer_st = request("FM_stempel_ignorerper_st_t") &":"& request("FM_stempel_ignorerper_st_m") &":00"
			ignorer_sl = request("FM_stempel_ignorerper_sl_t") &":"& request("FM_stempel_ignorerper_sl_m") &":00"
			
			        if IsDate(ignorer_st) then
			        else
			        errThis = 2
			        end if
			        
			        if IsDate(ignorer_sl) then
			        else
			        errThis = 2
			        end if
			 
			          
			          
			        if errThis = 2 then  %>
					     <!--include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 80
					call showError(errortype)
					isInt = 0
			        Response.End
			        
			        end if
			
			
			'** Standard Pause ***
			if len(request("FM_stempel_standardpause_A")) <> 0 then
			stpause = request("FM_stempel_standardpause_A")
			else
			stpause = 0
			end if
			
			if len(request("FM_stempel_standardpause_B")) <> 0 then
			stpause2 = request("FM_stempel_standardpause_B")
			else
			stpause2 = 0
			end if
			
			if len(request("CHKcrm")) <> 0 then
			crm = 1
			else
			crm = 0
			end if
			
			if len(request("CHKerp")) <> 0 then
			erp = 1
			else
			erp = 0
			end if
			
			if len(request("CHKsdsk")) <> 0 then
			sdsk = 1
			else
			sdsk = 0
			end if
			
			if len(request("p1_man")) <> 0 then
			p1_man = 1
			else
			p1_man = 0
			end if
			
			if len(request("p1_tir")) <> 0 then
			p1_tir = 1
			else
			p1_tir = 0
			end if
			
			if len(request("p1_ons")) <> 0 then
			p1_ons = 1
			else
			p1_ons = 0
			end if
			
			if len(request("p1_tor")) <> 0 then
			p1_tor = 1
			else
			p1_tor = 0
			end if
			
			if len(request("p1_fre")) <> 0 then
			p1_fre = 1
			else
			p1_fre = 0
			end if
			
			if len(request("p1_lor")) <> 0 then
			p1_lor = 1
			else
			p1_lor = 0
			end if
			
			if len(request("p1_son")) <> 0 then
			p1_son = 1
			else
			p1_son = 0
			end if
			
			if len(request("p2_man")) <> 0 then
			p2_man = 1
			else
			p2_man = 0
			end if
			
			if len(request("p2_tir")) <> 0 then
			p2_tir = 1
			else
			p2_tir = 0
			end if
			
			if len(request("p2_ons")) <> 0 then
			p2_ons = 1
			else
			p2_ons = 0
			end if
			
			if len(request("p2_tor")) <> 0 then
			p2_tor = 1
			else
			p2_tor = 0
			end if
			
			if len(request("p2_fre")) <> 0 then
			p2_fre = 1
			else
			p2_fre = 0
			end if
			
			if len(request("p2_lor")) <> 0 then
			p2_lor = 1
			else
			p2_lor = 0
			end if
			
			if len(request("p2_son")) <> 0 then
			p2_son = 1
			else
			p2_son = 0
			end if
			

            p2_grp = replace(request("FM_p2_grp"), "''", "")
            p1_grp = replace(request("FM_p1_grp"), "''", "")
			
			%>
			<!--#include file="inc/isint_func.asp"-->
			<%
			 
			call erDetInt(stpause)
			if isInt > 0 then
			errThis = 3
			else
			errThis = 0
			end if
			isInt = 0 
			
			call erDetInt(stpause2)
			if isInt > 0 then
			errThis = 3
			else
			errThis = 0
			end if
			isInt = 0 
			
			        if errThis = 3 then  %>
					     <!--include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 81
					call showError(errortype)
					isInt = 0
			        Response.End
			        
			        end if
			
			
			
			'**** fakturnr r�kkef�lge ****
			if len(request("FM_erp_fakturanr")) <> 0 then
			fakturanr = request("FM_erp_fakturanr")
			        
			        call erDetInt(fakturanr)
			        if isInt > 0 then
			        errThis = 4
			        else
			        errThis = errThis
			        end if
			        isInt = 0 

                    fakturanr = replace(fakturanr, ",", "")
			        fakturanr = replace(fakturanr, ".", "")
			end if
			

            '**** fakturnr r�kkef�lge ****
			if len(request("FM_erp_fakturanr_kladde")) <> 0 then
			fakturanr_kladde = request("FM_erp_fakturanr_kladde")
			        
			        call erDetInt(fakturanr_kladde)
			        if isInt > 0 then
			        errThis = 4
			        else
			        errThis = errThis
			        end if
			        isInt = 0 

                    fakturanr_kladde = replace(fakturanr_kladde, ",", "")
			        fakturanr_kladde = replace(fakturanr_kladde, ".", "")
			        
			end if

           

			rykkernr = 0
			
			if len(request("FM_erp_kreditnr")) <> 0 then
			kreditnr = request("FM_erp_kreditnr")
			        
			        call erDetInt(kreditnr)
			        if isInt > 0 then
			        errThis = 4
			        else
			        errThis = errThis
			        end if
			        isInt = 0 

                    kreditnr = replace(kreditnr, ",", "")
			        kreditnr = replace(kreditnr, ".", "")
			
			end if
			
		
			
			        
			        
			        if cint(errThis) = 4 then  %>
					     <!--include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 84
					call showError(errortype)
					Response.End
			        
			        end if
			        
			
            '***** Multible licensejere *************
            if len(trim(request("FM_multible_licensindehavere"))) <> 0 then
            multible_licensindehavere = 1
            else
            multible_licensindehavere = 0
            end if

            fakturanr_2 = request("FM_erp_fakturanr_2")
            kreditnr_2 = request("FM_erp_kreditnr_2")
            fakturanr_kladde_2 = request("FM_erp_fakturanr_kladde_2")

            fakturanr_3 = request("FM_erp_fakturanr_3")
            kreditnr_3 = request("FM_erp_kreditnr_3")
            fakturanr_kladde_3 = request("FM_erp_fakturanr_kladde_3")

            fakturanr_4 = request("FM_erp_fakturanr_4")
            kreditnr_4 = request("FM_erp_kreditnr_4")
            fakturanr_kladde_4 = request("FM_erp_fakturanr_kladde_4")
                           
            fakturanr_5 = request("FM_erp_fakturanr_5")
            kreditnr_5 = request("FM_erp_kreditnr_5")
            fakturanr_kladde_5 = request("FM_erp_fakturanr_kladde_5")
                           
                                   
			
			if len(request("FM_tsa_jobnr")) <> 0 then
			jobnr = request("FM_tsa_jobnr")
			        
			        call erDetInt(jobnr)
			        if cint(isInt) > 0 then
			        errThis = 5
			        else
			        errThis = errThis
			        end if
			        isInt = 0 
			
			end if
			
			
			if len(request("FM_tsa_tilbudsnr")) <> 0 then
			tilbudsnr = request("FM_tsa_tilbudsnr")
			        
			        call erDetInt(tilbudsnr)
			        if cint(isInt) > 0 then
			        errThis = 5
			        else
			        errThis = errThis
			        end if
			        isInt = 0 
			
			end if
			

            if len(trim(request("FM_regnskabsaar_start"))) then
                           
                            regnskabsaar_start = request("FM_regnskabsaar_start") & "-2001"
                            if isDate(regnskabsaar_start) = false then
                            regnskabsaar_start = "2001/1/1"
                            else
                            regnskabsaar_start = year(regnskabsaar_start) &"/"& month(regnskabsaar_start) & "/"& day(regnskabsaar_start) 
                            end if

            else
                             regnskabsaar_start  = "2001-01-01"
            end if


            if len(trim(request("FM_kmpris_ja"))) <> 0 then 

                    if len(request("FM_kmpris")) <> 0 then
			        kmpris = request("FM_kmpris")
			        
			                call erDetInt(kmpris)
			                if cint(isInt) > 0 then
			                errThis = 10
			                else
			                errThis = errThis
			                end if
			                isInt = 0 
			

                    kmpris = replace(kmpris, ",", ".")

                    
                                                fraDato = request("FM_kmpris_fra") 
                                                if isDate(fraDato) = false then
                                                fraDato = year(now) &"/"& month(now) & "/"& day(now)
                                                else
                                                fraDato = year(fraDato) &"/"& month(fraDato) & "/"& day(fraDato) 
                                                end if

                                      
                                          call basisValutaFN()
                                    
                                     
                                     strSQLopdAKM = "UPDATE aktiviteter SET brug_fasttp = 1, fasttp = "& kmpris &", fasttp_val = "& basisValId &" WHERE id <> 0 AND aktstatus = 1 AND fakturerbar = 5" 
                                     oConn.execute(strSQLopdAKM)

                                 

                                     strSQLopdTKM = "UPDATE timer SET timepris = "& kmpris &", valuta = "& basisValId &", kurs = "& basisValKurs &"  WHERE tdato >= '"& fraDato &"' AND Tfaktim = 5" 
                                     oConn.execute(strSQLopdTKM)


                    else

                    kmpris = "3.56"

			        end if
            
            else

            kmpris = request("FM_kmpris_old")
            kmpris = replace(kmpris, ",", ".")

            end if

             
                   
			        
			        if cint(errThis) = 5 then  %>
					     <!--include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 91
					call showError(errortype)
		            Response.End
			        
			        end if
			        
			        
                        if cint(errThis) = 10 then  %>
					     <!--include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 157
					call showError(errortype)
		            Response.End
			        
			        end if

			'if len(request("FM_erp_fakprocent")) <> 0 then
			'fakprocent = request("FM_erp_fakprocent")
			'        
			'        call erDetInt(fakprocent)
			'        if cint(isInt) > 0 then
			'        errThis = 6
			'        else
			'        errThis = errThis
			'        end if
			'        isInt = 0 
			
			'end if
			
			'*** Bruges ikke ***
			fakprocent = 0
			
			        if cint(errThis) = 6 then  %>
					     <!--include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 96
					call showError(errortype)
		            Response.End
			        
			        end if
			        
			 
			 if len(trim(request("licensstdato"))) <> 0 then      
			 licensstdato = request("licensstdato")
			 end if
			 
			        if isDate(licensstdato) = false then  %>
					     <!--include file="../inc/regular/header_lysblaa_inc.asp"-->
					     <!--include file="../inc/regular/topmenu_inc.asp"-->
			           <%
					errortype = 137
					call showError(errortype)
		            Response.End
			        
			        end if
			 
			 
			 licensstdato = year(licensstdato) & "/" & month(licensstdato) &"/"& day(licensstdato)
			
            
            '*** Opdaterer vis funktion *****'
            if len(trim(request("jobasnvigv"))) <> 0 then
            jobasnvigv = 1
            else
            jobasnvigv = 0
            end if   
                           
            if len(trim(request("akt_maksbudget_treg"))) <> 0 then
            akt_maksbudget_treg = 1
            else
            akt_maksbudget_treg = 0
            end if   
                             
			
			if len(trim(request("positiv_aktivering_akt"))) <> 0 then
            positiv_aktivering_akt = 1
            else
            positiv_aktivering_akt = 0
            end if

            
			if len(trim(request("pa_aktlist"))) <> 0 then
            pa_aktlist = 1
            else
            pa_aktlist = 0
            end if

            

            
            if len(trim(request("FM_lukaktvdato"))) <> 0 then
            lukaktvdato = request("FM_lukaktvdato")
            else
            lukaktvdato = 0
            'response.Cookies("tsa")("ignJobogAktper") = "" 20160916
            end if


            if len(trim(request("FM_salgsans"))) <> 0 then
            salgsans = 1
            else
            salgsans = 0
            end if
                           
            
            if len(trim(request("showeasyreg"))) <> 0 then
            showeasyreg = 1
            else
            showeasyreg = 0
            end if


            if len(trim(request("forcebudget_onakttreg"))) <> 0 then
            forcebudget_onakttreg = 1
            else
            forcebudget_onakttreg = 0
            end if


            if len(trim(request("akt_maksforecast_treg"))) <> 0 then
            akt_maksforecast_treg = 1
            else
            akt_maksforecast_treg = 0
            end if

                           

            if len(trim(request("forcebudget_onakttreg_filt_viskunmbgt"))) <> 0 then
            forcebudget_onakttreg_filt_viskunmbgt = 1
            else
            forcebudget_onakttreg_filt_viskunmbgt = 0
            end if

            
            

            if len(trim(request("FM_smileyAggressiv"))) <> 0 then
            smileyAggressiv = 1
            else
            smileyAggressiv = 0
            end if

            if len(trim(request("FM_smiley_agg_lukhard"))) <> 0 then
            smiley_agg_lukhard = 1
            else
            smiley_agg_lukhard = 0
            end if

                           

            if len(trim(request("FM_hidesmileyicon"))) <> 0 then
            hidesmileyicon = 1
            else
            hidesmileyicon = 0
            end if
                           
            if len(trim(request("FM_minimumslageremail"))) <> 0 then
            minimumslageremail = 1
            else
            minimumslageremail = 0
            end if
            

            
            if len(trim(request("forcebudget_onakttreg_afgr"))) <> 0 then
            forcebudget_onakttreg_afgr = request("forcebudget_onakttreg_afgr")
            else
            forcebudget_onakttreg_afgr = 0
            end if

            
            if len(trim(request("fomr_account"))) <> 0 then
            fomr_account = 1
            else
            fomr_account = 0
            end if
            
            



            if len(trim(request("showupload"))) <> 0 then
            showupload = 1
            else
            showupload = 0
            end if


              if len(trim(request("FM_week_showbase_norm_kommegaa"))) <> 0 then
              week_showbase_norm_kommegaa = 1
              else
              week_showbase_norm_kommegaa = 0
              end if

              if len(trim(request("FM_mobil_week_reg_job_dd"))) <> 0 then
              mobil_week_reg_job_dd = 1
              else
              mobil_week_reg_job_dd = 0
              end if

              if len(trim(request("FM_mobil_week_reg_akt_dd"))) <> 0 then
              mobil_week_reg_akt_dd = 1
              else
              mobil_week_reg_akt_dd = 0
              end if

              if len(trim(request("FM_mobil_week_reg_akt_dd_forvalgt"))) <> 0 then
              mobil_week_reg_akt_dd_forvalgt = 1
              else
              mobil_week_reg_akt_dd_forvalgt = 0
              end if
          

       
     
            if len(trim(request("bdgmtypon"))) <> 0 then
            bdgmtypon = 1
            else
            bdgmtypon = 0
            end if


            if len(trim(request("medarbtypligmedarb"))) <> 0 then
            medarbtypligmedarb = 1
            else
            medarbtypligmedarb = 0
            end if

        


            if len(trim(request("FM_teamleder_flad"))) <> 0 then
            teamleder_flad = 1
            else
            teamleder_flad = 0
            end if

            if len(trim(request("timesimh1h2"))) <> 0 then
            timesimh1h2 = 1
            else
            timesimh1h2 = 0
            end if

            if len(trim(request("timesimon"))) <> 0 then
            timesimon = 1
            else
            timesimon = 0
            end if

            if len(trim(request("timesimtp"))) <> 0 then
            timesimtp = 1
            else
            timesimtp = 0
            end if

            if len(trim(request("traveldietexp_on"))) <> 0 then
            traveldietexp_on = 1
            else
            traveldietexp_on = 0
            end if

            if len(trim(request("traveldietexp_maxhours"))) <> 0 then
            traveldietexp_maxhours = request("traveldietexp_maxhours")

                     call erDetInt(traveldietexp_maxhours)
			         if cint(isInt) > 0 then
                           traveldietexp_maxhours = 0
                     end if

            traveldietexp_maxhours = replace(traveldietexp_maxhours, ",", ".")
            else
            traveldietexp_maxhours = 0
            end if
                            
            
                           


            SmiWeekOrMonth = request("FM_SmiWeekOrMonth")

            SmiantaldageCount = request("FM_SmiantaldageCount")


                           select case SmiWeekOrMonth 
                           case 1 'm�nedaafslutning
                           
                           if SmiantaldageCount < 8 then
                           SmiantaldageCount = 8 
                           else
                           SmiantaldageCount = SmiantaldageCount
                           end if

                           case 2 'dagligt
                           SmiantaldageCount = 11

                           case else 'uge affslutning

                           if SmiantaldageCount > 7 then
                           SmiantaldageCount = 1 
                           else
                           SmiantaldageCount = SmiantaldageCount
                           end if

                           end select


            SmiantaldageCountClock = request("FM_SmiantaldageCountClock")
            SmiTeamlederCount = request("FM_SmiTeamlederCount")

            if len(trim(request("fomr_mandatory"))) <> 0 then
            fomr_mandatory = 1
            else
            fomr_mandatory = 0
            end if


            if len(trim(request("budget_mandatory"))) <> 0 then
            budget_mandatory = 1
            else
            budget_mandatory = 0
            end if

            if len(trim(request("tilbud_mandatory"))) <> 0 then
            tilbud_mandatory = 1
            else
            tilbud_mandatory = 0
            end if

            if len(trim(request("show_salgsomk_mandatory"))) <> 0 then
            show_salgsomk_mandatory = 1
            else
            show_salgsomk_mandatory = 0
            end if

            
                           
            
            '**** T�mmer alle aktive jovblister ved �ndring af Positiv indstilling ***'
            strSQLSel = "SELECT positiv_aktivering_akt FROM licens WHERE id = 1"
            oldPositiv_aktivering_akt = 0
            oRec.open strSQLSel, oConn, 3
            if not oRec.EOF then
            oldPositiv_aktivering_akt = oRec("positiv_aktivering_akt")
            end if 
            oRec.close

            if cint(oldPositiv_aktivering_akt) <> cint(positiv_aktivering_akt) then
            'Response.write "SLETEER"
            'Response.end
            strSQldel = "DELETE FROM timereg_usejob WHERE id <> 0"
            oConn.execute(strSQldel)
            end if

			%>
			<!-- Opdater db efter validering -->
			<%
			
                 

			'ignorertid_st, ignorertid_sl 
			strSQL = "UPDATE licens SET owa = '" & exchangeserverURL & "', dom = '"& exchangeserverDOM &"', "_
			&" kontonavn = '"& exchangeserverBruger &"',  kontopw = '"& exchangeserverPW &"', "_
			&" smiley = '"& smiley &"', stempelur = '"& stempelur &"', "_
			&" lukafm = "& lukafm &", autogk = "& autogk &", autogktimer = "& autogktimer &", "_
			&" autolukvdato = " & autolukvdato & ", autolukvdatodato = " & autolukvdatodato &", "_
			&" ignorertid_st = '"& ignorer_st &"', ignorertid_sl = '"& ignorer_sl &"', "_
			&" stpause = "& stpause &", stpause2 = "& stpause2 &", "_
			&" crm = "& crm &", "_
			&" erp = "& erp &", "_
			&" sdsk = "& sdsk &", "_
			&" p1_man = "& p1_man &", "_
			&" p1_tir = "& p1_tir &", "_
			&" p1_ons = "& p1_ons &", "_
			&" p1_tor = "& p1_tor &", "_
			&" p1_fre = "& p1_fre &", "_
			&" p1_lor = "& p1_lor &", "_
			&" p1_son = "& p1_son &", "_
			&" p2_man = "& p2_man &", "_
			&" p2_tir = "& p2_tir &", "_
			&" p2_ons = "& p2_ons &", "_
			&" p2_tor = "& p2_tor &", "_
			&" p2_fre = "& p2_fre &", "_
			&" p2_lor = "& p2_lor &", "_
			&" p2_son = "& p2_son &", fakturanr = "& abs(fakturanr) &", "_
			&" rykkernr = "& abs(rykkernr) &", kreditnr = "& kreditnr &", "_
			&" jobnr = "& jobnr &", tilbudsnr = "& tilbudsnr &", "_
			&" fakprocent = "& fakprocent &", kmdialog = "& kmDialog &", licensstdato = '"& licensstdato &"'"_
            &", fakturanr_kladde = "& abs(fakturanr_kladde) &", "_
            &" kmpris = "& kmpris &", p1_grp = '"& p1_grp &"', p2_grp = '"& p2_grp &"', jobasnvigv = "& jobasnvigv &" , positiv_aktivering_akt = " & positiv_aktivering_akt & ", "_
            &" showeasyreg = "& showeasyreg & ", forcebudget_onakttreg = "& forcebudget_onakttreg &", showupload = "& showupload & ", globalfaktor = "& globalfaktor & ", "_
            &" bdgmtypon = " & bdgmtypon & ", regnskabsaar_start = '"& regnskabsaar_start &"', "_
            &" forcebudget_onakttreg_afgr = "& forcebudget_onakttreg_afgr &", lukaktvdato = "& lukaktvdato &", salgsans = "& salgsans &","_
            &" smileyaggressiv = "& smileyAggressiv &", timerround = "& timerround &", teamleder_flad = "& teamleder_flad &", "_
            &" forcebudget_onakttreg_filt_viskunmbgt = "& forcebudget_onakttreg_filt_viskunmbgt &", "_
            &" stempelur_hidelogin = "& stempelur_hidelogin &", stempelur_igno_komkrav = "& stempelur_igno_komkrav &", "_
            &" SmiWeekOrMonth = "& SmiWeekOrMonth &", SmiantaldageCount = "& SmiantaldageCount &", "_
            &" SmiantaldageCountClock = "& SmiantaldageCountClock &", SmiTeamlederCount = "& SmiTeamlederCount & ", "_
            &" hidesmileyicon = "& hidesmileyicon &", visAktlinjerSimpel = "& visAktlinjerSimpel &", "_
            &" fomr_mandatory = "& fomr_mandatory &", akt_maksbudget_treg = "& akt_maksbudget_treg &", minimumslageremail = "& minimumslageremail &", fomr_account = "& fomr_account &", "_
            &" visAktlinjerSimpel_datoer = "& visAktlinjerSimpel_datoer &", visAktlinjerSimpel_timebudget = "& visAktlinjerSimpel_timebudget &", visAktlinjerSimpel_realtimer = "& visAktlinjerSimpel_realtimer &", visAktlinjerSimpel_restimer = "& visAktlinjerSimpel_restimer &", "_
            &" visAktlinjerSimpel_medarbtimepriser = "& visAktlinjerSimpel_medarbtimepriser &", visAktlinjerSimpel_medarbrealtimer = "& visAktlinjerSimpel_medarbrealtimer &", "_
            &" visAktlinjerSimpel_akttype = "& visAktlinjerSimpel_akttype &", timesimon = "& timesimon &", timesimh1h2 = "& timesimh1h2 & ", "_
            &" timesimtp = "& timesimtp &", budgetakt = " & budgetakt &", akt_maksforecast_treg = "& akt_maksforecast_treg &", "_
            &" traveldietexp_on = "& traveldietexp_on &", traveldietexp_maxhours = "& traveldietexp_maxhours & ", "_
            &" medarbtypligmedarb = " & medarbtypligmedarb &", pa_aktlist = " & pa_aktlist & ", "_
            &" smiley_agg_lukhard = " & smiley_agg_lukhard & ","_
            &" week_showbase_norm_kommegaa = "& week_showbase_norm_kommegaa &", mobil_week_reg_job_dd = "& mobil_week_reg_job_dd &", "_
            &" mobil_week_reg_akt_dd = " & mobil_week_reg_akt_dd & ", mobil_week_reg_akt_dd_forvalgt = " & mobil_week_reg_akt_dd_forvalgt & ","_
            &" budget_mandatory = "& budget_mandatory &", tilbud_mandatory = "& tilbud_mandatory &", show_salgsomk_mandatory = "& show_salgsomk_mandatory &", "_
            &" multible_licensindehavere = "& multible_licensindehavere &", "_
            &" fakturanr_2 = "& fakturanr_2 &", "_
            &" kreditnr_2 = "& kreditnr_2 &", "_
            &" fakturanr_kladde_2  = "& fakturanr_kladde_2  &","_
                &" fakturanr_3 = "& fakturanr_3 &", "_
            &" kreditnr_3 = "& kreditnr_3 &", "_
            &" fakturanr_kladde_3  = "& fakturanr_kladde_3  &","_
                &" fakturanr_4 = "& fakturanr_4 &", "_
            &" kreditnr_4 = "& kreditnr_4 &", "_
            &" fakturanr_kladde_4  = "& fakturanr_kladde_4  &","_
                &" fakturanr_5 = "& fakturanr_5 &", "_
            &" kreditnr_5 = "& kreditnr_5 &", "_
            &" fakturanr_kladde_5  = "& fakturanr_kladde_5  

			strSQL = strSQL & strSQLat & " WHERE id = 1"
			
			'Response.Write strSQL
			'Response.Flush
			

			oConn.execute(strSQL)
			
			Response.redirect "kontrolpanel.asp?menu=tok&func=exchopd"
			
	case "dwldb"			
	
					
					strSQL = "SELECT Mid, email, mnavn FROM medarbejdere WHERE Mid=" & session("mid")
					oRec.open strSQL, oConn, 3
					if not oRec.EOF then
					strEmail = oRec("email")
					strEditor = oRec("mnavn")
					end if
					oRec.close
					
	
					Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
					' S�tter Charsettet til ISO-8859-1 
					Mailer.CharSet = 2
					' Afsenderens navn 
					Mailer.FromName = "TimeOut DB Backup"
					' Afsenderens e-mail 
					Mailer.FromAddress = "support@outzource.dk"
					Mailer.RemoteHost = "smtp.tiscali.dk"
					' Modtagerens navn og e-mail
					Mailer.AddRecipient "Support OutZourCE", "support@outzource.dk"
					'Mailer.AddBCC "SK", "sk@outzource.dk" 
					' Mailens emne
					Mailer.Subject = "DB backup er bestilt af: "& lto &" !"
					
					' Selve teksten
					Mailer.BodyText = "" & "Backup er bestilt af "& strNavn & vbCrLf _ 
					& "Sendes til: " &strEmail & "" & vbCrLf 
					
					Mailer.SendMail
					
					strDato = year(now)&"/"&month(now)&"/"&day(now)
					
						oConn.execute("INSERT INTO dbdownload (dato, editor, email) VALUES ("_
						&"'"& strDato &"',"_
						&"'"& strEditor &"',"_
						&"'"& strEmail &"')")
						
					Response.redirect "kontrolpanel.asp?menu=tok"
				
	case else
%>
<!--#include file="../inc/regular/header_lysblaa_inc.asp"-->
<!--#include file="inc/dato.asp"-->
<!--include file="../inc/regular/topmenu_inc.asp"-->

<script language="javascript">
    var ns, ns6, ie, newlayer;

    ns4 = (document.layers) ? true : false;
    ie4 = (document.all) ? true : false
    ie5 = (document.getElementById) ? true : false
    ns6 = (document.getElementById && !document.all) ? true : false;

    function getLayerStyle(lyr) {
        if (ns4) {
            return document.layers[lyr];
        } else if (ie4) {
            return document.all[lyr].style;
        } else if (ie5) {
            return document.all[lyr].style;
        } else if (ns6) {
            return document.getElementById(lyr).style;
        }
    }

    function ShowHide(layer) {
        newlayer = getLayerStyle(layer)

        var styleObj = (ns4) ? document.layers[layer] : (ie4) ? document.all[layer].style : document.getElementById(layer).style;

        if (newlayer.visibility == "hidden") {
            newlayer.visibility = "visible";
            styleObj.display = ""
        }
        else if (newlayer.visibility == "visible") {
            newlayer.visibility = "hidden";
            styleObj.display = "none"
        }
    }
</script>

<%call menu_2014() %>

<div id="sindhold" style="position: absolute; left:90px; top:102px; visibility: visible; width: 900px">
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr>
            
   <td valign="bottom">
                <h3>TimeOut Kontrolpanel</h3>
                
               </td>
        </tr>
    </table>
    <br>
    <table cellspacing="0" cellpadding="0" border="0" width="80%">
        <form method="post" action="kontrolpanel.asp?menu=tok&func=exch">
        
        <%
		strSQL = "SELECT owa, kontonavn, kontopw, dom, smiley, stempelur, "_
		&" lukafm, autogk, brugabningstid, "_
		&" normtid_st_man, "_ 
	    &" normtid_sl_man, "_
	    &" normtid_st_tir, "_
		&" normtid_sl_tir, "_
		&" normtid_st_ons, "_
		&" normtid_sl_ons, "_
		&" normtid_st_tor, "_
		&" normtid_sl_tor, "_
		&" normtid_st_fre, "_
		&" normtid_sl_fre, "_
		&" normtid_st_lor, "_
		&" normtid_sl_lor, "_
		&" normtid_st_son, "_
		&" normtid_sl_son, autolukvdato, "_
		&" autolukvdatodato, ignorertid_st, ignorertid_sl, stpause, stpause2, "_
		&" p1_man, "_
		&" p1_tir, "_
		&" p1_ons, "_
		&" p1_tor, "_
		&" p1_fre, "_
		&" p1_lor, "_
		&" p1_son, "_
		&" p2_man, "_
		&" p2_tir, "_
		&" p2_ons, "_
		&" p2_tor, "_
		&" p2_fre, "_
		&" p2_lor, "_
		&" p2_son, fakturanr, kreditnr, rykkernr, jobnr, tilbudsnr, fakprocent, "_
		&" autogktimer, kmdialog, licensstdato, fakturanr_kladde, kmpris, p1_grp, p2_grp, jobasnvigv, "_
        &" positiv_aktivering_akt, showeasyreg, showupload, forcebudget_onakttreg, globalfaktor, bdgmtypon, "_
        &" regnskabsaar_start, forcebudget_onakttreg_afgr, lukaktvdato, salgsans, smileyaggressiv, timerround, teamleder_flad, "_
        &" forcebudget_onakttreg_filt_viskunmbgt, stempelur_hidelogin, stempelur_igno_komkrav, "_
        &" SmiWeekOrMonth, SmiantaldageCount, SmiantaldageCountClock, SmiTeamlederCount, hidesmileyicon, visAktlinjerSimpel, fomr_mandatory, akt_maksbudget_treg, minimumslageremail, fomr_account, "_
        &" visAktlinjerSimpel_datoer, visAktlinjerSimpel_timebudget, visAktlinjerSimpel_realtimer, visAktlinjerSimpel_restimer, "_
        &" visAktlinjerSimpel_medarbtimepriser, visAktlinjerSimpel_medarbrealtimer, visAktlinjerSimpel_akttype, timesimon, timesimh1h2, timesimtp, budgetakt, akt_maksforecast_treg, "_
        &" traveldietexp_on, traveldietexp_maxhours, medarbtypligmedarb, pa_aktlist, smiley_agg_lukhard, week_showbase_norm_kommegaa, mobil_week_reg_job_dd, mobil_week_reg_akt_dd, mobil_week_reg_akt_dd_forvalgt, "_
        &" budget_mandatory, tilbud_mandatory, show_salgsomk_mandatory, "_
        &" multible_licensindehavere, "_
        &" fakturanr_2, "_
        &" kreditnr_2, "_
        &" fakturanr_kladde_2,"_
        &" fakturanr_3, "_
        &" kreditnr_3, "_
        &" fakturanr_kladde_3,"_
        &" fakturanr_4, "_
        &" kreditnr_4, "_
        &" fakturanr_kladde_4, "_
        &" fakturanr_5, "_
        &" kreditnr_5, "_
        &" fakturanr_kladde_5"_
        &" FROM licens WHERE id = 1"
		
		'Response.Write strSQL
		'Response.Flush
		
		oRec.open strSQL, oConn, 3 
		if not oRec.EOF then

			ExchangeServerURL = oRec("owa")
			ExchangeServerDOM = oRec("dom")
			ExchangeServerBruger = replace(oRec("kontonavn"), "#", "\")
			ExchangeServerPW = oRec("kontopw")
			smiley = oRec("smiley")
			stempelur = oRec("stempelur")
			lukafm = oRec("lukafm")
			autogk = oRec("autogk")
			autogktimer = oRec("autogktimer")
			
			normtid_st_man = left(formatdatetime(oRec("normtid_st_man"), 3), 5) 
	        normtid_sl_man = left(formatdatetime(oRec("normtid_sl_man"), 3), 5)
	        normtid_st_tir = left(formatdatetime(oRec("normtid_st_tir"), 3), 5)
		    normtid_sl_tir = left(formatdatetime(oRec("normtid_sl_tir"), 3), 5)
		    normtid_st_ons = left(formatdatetime(oRec("normtid_st_ons"), 3), 5)
		    normtid_sl_ons = left(formatdatetime(oRec("normtid_sl_ons"), 3), 5) 
		    normtid_st_tor = left(formatdatetime(oRec("normtid_st_tor"), 3), 5)
		    normtid_sl_tor = left(formatdatetime(oRec("normtid_sl_tor"), 3), 5) 
		    normtid_st_fre = left(formatdatetime(oRec("normtid_st_fre"), 3), 5) 
		    normtid_sl_fre = left(formatdatetime(oRec("normtid_sl_fre"), 3), 5) 
		    normtid_st_lor = left(formatdatetime(oRec("normtid_st_lor"), 3), 5) 
		    normtid_sl_lor = left(formatdatetime(oRec("normtid_sl_lor"), 3), 5) 
		    normtid_st_son = left(formatdatetime(oRec("normtid_st_son"), 3), 5) 
		    normtid_sl_son = left(formatdatetime(oRec("normtid_sl_son"), 3), 5)
			
			brugabningstid = oRec("brugabningstid")
			
			strDag = oRec("autolukvdatodato")
			autolukvdato = oRec("autolukvdato") 
			
			ignorertid_st = left(formatdatetime(oRec("ignorertid_st"), 3), 5)
			ignorertid_sl = left(formatdatetime(oRec("ignorertid_sl"), 3), 5)
			
			stpauseA = oRec("stpause")
			stpauseB = oRec("stpause2")
			
			
			if oRec("p1_man") <> 0 then
			p1_manChk = "CHECKED"
			else
			p1_manChk = ""
			end if
			
			if oRec("p1_tir") <> 0 then
			p1_tirChk = "CHECKED"
			else
			p1_tirChk = ""
			end if
			
			if oRec("p1_ons") <> 0 then
			p1_onsChk = "CHECKED"
			else
			p1_onsChk = ""
			end if
			
			if oRec("p1_tor") <> 0 then
			p1_torChk = "CHECKED"
			else
			p1_torChk = ""
			end if
			
			if oRec("p1_fre") <> 0 then
			p1_freChk = "CHECKED"
			else
			p1_freChk = ""
			end if
			
			if oRec("p1_lor") <> 0 then
			p1_lorChk = "CHECKED"
			else
			p1_lorChk = ""
			end if
			
			if oRec("p1_son") <> 0 then
			p1_sonChk = "CHECKED"
			else
			p1_sonChk = ""
			end if
			
			if oRec("p2_man") <> 0 then
			p2_manChk = "CHECKED"
			else
			p2_manChk = ""
			end if
			
			if oRec("p2_tir") <> 0 then
			p2_tirChk = "CHECKED"
			else
			p2_tirChk = ""
			end if
			
			if oRec("p2_ons") <> 0 then
			p2_onsChk = "CHECKED"
			else
			p2_onsChk = ""
			end if
			
			if oRec("p2_tor") <> 0 then
			p2_torChk = "CHECKED"
			else
			p2_torChk = ""
			end if
			
			if oRec("p2_fre") <> 0 then
			p2_freChk = "CHECKED"
			else
			p2_freChk = ""
			end if
			
			if oRec("p2_lor") <> 0 then
			p2_lorChk = "CHECKED"
			else
			p2_lorChk = ""
			end if
			
			if oRec("p2_son") <> 0 then
			p2_sonChk = "CHECKED"
			else
			p2_sonChk = ""
			end if
			

            p1_grp = oRec("p1_grp")
            p2_grp = oRec("p2_grp")


			fakturanr = oRec("fakturanr")
			rykkernr = oRec("rykkernr")
			kreditnr = oRec("kreditnr")
            fakturanr_kladde = oRec("fakturanr_kladde")
            
			jobnr = oRec("jobnr")
			tilbudsnr = oRec("tilbudsnr")
			fakprocent = oRec("fakprocent")
			
			kmDialog = oRec("kmdialog")
			
			licensstdato = oRec("licensstdato")

            kmpris = replace(oRec("kmpris"), ".", ",")

            jobasnvigv = oRec("jobasnvigv")

            positiv_aktivering_akt = oRec("positiv_aktivering_akt")
            pa_aktlist = oRec("pa_aktlist")

			
            showeasyreg = oRec("showeasyreg")
            forcebudget_onakttreg = oRec("forcebudget_onakttreg")
            showupload = oRec("showupload")

            globalfaktor = oRec("globalfaktor")

            bdgmtypon = oRec("bdgmtypon")


            regnskabsaar_start = oRec("regnskabsaar_start")

            forcebudget_onakttreg_afgr = oRec("forcebudget_onakttreg_afgr")
            akt_maksforecast_treg = oRec("akt_maksforecast_treg")

            lukaktvdato = oRec("lukaktvdato")

            salgsans = oRec("salgsans")

            smileyaggressiv = oRec("smileyaggressiv")

            timerround = oRec("timerround")

            teamleder_flad = oRec("teamleder_flad")

            forcebudget_onakttreg_filt_viskunmbgt = oRec("forcebudget_onakttreg_filt_viskunmbgt")

            stempelur_hidelogin = oRec("stempelur_hidelogin")
            stempelur_igno_komkrav = oRec("stempelur_igno_komkrav")


            SmiWeekOrMonth = oRec("SmiWeekOrMonth")
            SmiantaldageCount = oRec("SmiantaldageCount")
            SmiantaldageCountClock = oRec("SmiantaldageCountClock")
            SmiTeamlederCount = oRec("SmiTeamlederCount") 


            hidesmileyicon = oRec("hidesmileyicon")
            visAktlinjerSimpel = oRec("visAktlinjerSimpel")

            fomr_mandatory = oRec("fomr_mandatory")

            akt_maksbudget_treg = oRec("akt_maksbudget_treg")

            minimumslageremail = oRec("minimumslageremail")

            fomr_account = oRec("fomr_account")


            visAktlinjerSimpel_datoer = oRec("visAktlinjerSimpel_datoer")
            visAktlinjerSimpel_timebudget = oRec("visAktlinjerSimpel_timebudget")
            visAktlinjerSimpel_realtimer = oRec("visAktlinjerSimpel_realtimer")
            visAktlinjerSimpel_restimer = oRec("visAktlinjerSimpel_restimer")

            visAktlinjerSimpel_medarbtimepriser = oRec("visAktlinjerSimpel_medarbtimepriser") 
            visAktlinjerSimpel_medarbrealtimer = oRec("visAktlinjerSimpel_medarbrealtimer")
            visAktlinjerSimpel_akttype = oRec("visAktlinjerSimpel_akttype")


            timesimh1h2 = oRec("timesimh1h2")
            timesimon = oRec("timesimon")
            timesimtp = oRec("timesimtp")

            budgetakt = oRec("budgetakt")

            traveldietexp_on = oRec("traveldietexp_on")
            traveldietexp_maxhours = oRec("traveldietexp_maxhours")

            medarbtypligmedarb = oRec("medarbtypligmedarb")

            smiley_agg_lukhard = oRec("smiley_agg_lukhard")

            week_showbase_norm_kommegaa = oRec("week_showbase_norm_kommegaa")
            mobil_week_reg_job_dd = oRec("mobil_week_reg_job_dd")
            mobil_week_reg_akt_dd = oRec("mobil_week_reg_akt_dd") 
            mobil_week_reg_akt_dd_forvalgt = oRec("mobil_week_reg_akt_dd_forvalgt")

            budget_mandatory = oRec("budget_mandatory")
            tilbud_mandatory = oRec("tilbud_mandatory")
            show_salgsomk_mandatory = oRec("show_salgsomk_mandatory")


            multible_licensindehavere = oRec("multible_licensindehavere")
            fakturanr_2 = oRec("fakturanr_2") 
            kreditnr_2 = oRec("kreditnr_2")
            fakturanr_kladde_2 = oRec("fakturanr_kladde_2")
            fakturanr_3 = oRec("fakturanr_3")
            kreditnr_3 = oRec("kreditnr_3")
            fakturanr_kladde_3 = oRec("fakturanr_kladde_3")
            fakturanr_4 = oRec("fakturanr_4")
            kreditnr_4 = oRec("kreditnr_4")
            fakturanr_kladde_4 = oRec("fakturanr_kladde_4")
            fakturanr_5 = oRec("fakturanr_5")
            kreditnr_5 = oRec("kreditnr_5")
            fakturanr_kladde_5 = oRec("fakturanr_kladde_5")

		end if
		oRec.close 


            
            if len(strDag) <> 0 Then
            strDag = strDag 
            else
            strDag = 28
            end if
            
            
             
                       
            man_t = left(normtid_st_man, 2)
            man_m = right(normtid_st_man, 2)
            man_t2 = left(normtid_sl_man, 2)
            man_m2 = right(normtid_sl_man, 2)
            
            tir_t = left(normtid_st_tir, 2)
            tir_m = right(normtid_st_tir, 2)
            tir_t2 = left(normtid_sl_tir, 2)
            tir_m2 = right(normtid_sl_tir, 2)
            
            ons_t = left(normtid_st_ons, 2)
            ons_m = right(normtid_st_ons, 2)
            ons_t2 = left(normtid_sl_ons, 2)
            ons_m2 = right(normtid_sl_ons, 2)
            
            tor_t = left(normtid_st_tor, 2)
            tor_m = right(normtid_st_tor, 2)
            tor_t2 = left(normtid_sl_tor, 2)
            tor_m2 = right(normtid_sl_tor, 2)
            
            fre_t = left(normtid_st_fre, 2)
            fre_m = right(normtid_st_fre, 2)
            fre_t2 = left(normtid_sl_fre, 2)
            fre_m2 = right(normtid_sl_fre, 2)
            
            
            lor_t = left(normtid_st_lor, 2)
            lor_m = right(normtid_st_lor, 2)
            lor_t2 = left(normtid_sl_lor, 2)
            lor_m2 = right(normtid_sl_lor, 2)
            
            son_t = left(normtid_st_son, 2)
            son_m = right(normtid_st_son, 2)
            son_t2 = left(normtid_sl_son, 2)
            son_m2 = right(normtid_sl_son, 2)
            
            
            
            ignorertid_st_t = left(ignorertid_st, 2)
            ignorertid_st_m = right(ignorertid_st, 2)
            
            ignorertid_sl_t = left(ignorertid_sl, 2)
            ignorertid_sl_m = right(ignorertid_sl, 2)

            if cint(timerround) = 1 then
            timerroundCHK = "CHECKED"
            else
            timerroundCHK = ""
            end if
            
           %>
           <tr>
                   <td bgcolor="#ffffff" style="border:1px #5582d2 solid; padding:10px;">
                <a href="javascript:ShowHide('kgnr');"><b>Generelt:</b></a></td>
                   </tr>
                  
                   
                   <tr><td style="padding:10px;"><div id="kgnr" name="kgnr" style="visibility: hidden; display: none">
                    
                    <b>Licens startdato:</b> <input id="licensstdato" name="licensstdato" type="text" style="width:100px;" value="<%=licensstdato%>"/> (dd-mm-����)
                    <br />
                    Bruges til at l�se hvor langt tilbage TimeOut skal kigge n�r det skal finde <b>norm. timer</b> p� en medarbejder.<br />
                    Hvis medarbejderen er ansat p� et senere tidspunkt bruges medarbejder <b>ansatdato</b> p� denne medarbejder.
                    
                    <br /><br />
                    <a href="kontrolpanel.asp?menu=tok&func=dwldb">Bestil database backup (zip fil)</a><br>
                    <%
		                strSQL = "SELECT dato, editor FROM dbdownload ORDER BY id DESC"
		                oRec.open strSQL, oConn, 3 
		                if not oRec.EOF then
		                ald = "n"
		                strDato = oRec("dato")
		                strEditor = oRec("editor")
		                else
		                ald = "j"
		                strDato = "--"
		                strEditor = ""
		                end if
		                oRec.close
                		
		                if ald = "n" then
		                Response.write "Sidst bestilt d. "& strDato &" af "& strEditor &"<br>"
		                end if
                    %>
                    Ved at bestille jeres egen lokale backup af jeres database, sikres det at I altid
                    har en version af jeres data, liggende p� jeres eget netv�rk. Backup'en bliver sendt
                    dig via email, og I modtager den inden for et d�gn. <b>Det koster 495 kr.</b>
                    at bestille en lokal kopi af jeres database. (SQL format)<br>
                    <br>
                    <br>
                    &nbsp;<img src="../ill/blank.gif" width="1" height="1" alt="" border="0"></div></td>
            </tr>
            </table>     
            <table cellpadding=0 cellspacing=0 border=0 width=720>
            <tr>
                <td width=520 bgcolor="#ffffff" style="border:1px #5582d2 solid; border-right:0px; padding:10px;">
                <a href="javascript:ShowHide('ktsa');"><b>TSA: (Time/sag)</b></a></td>
                  <td width=200 align="right" bgcolor="#ffffff" style="border:1px #5582d2 solid; border-left:0px; padding:10px;">
              Aktiver modul <b>*</b>:<input type="checkbox" name="CHKtsa" align="right" checked disabled/> </tr>
           <tr><td style="padding:10px;" colspan="2"><div id="ktsa" name="ktsa" style="visibility: hidden; display: none"> <table>
            <td style="padding:10px;">
                        
                    <ul>
                         <li><a href="akt_typer.asp?menu=tok" target="_blank" class='vmenu'>
                            Aktivitets typer</a>
               
                                    
                                   
                                   
            
            </td>
           </tr>
           
            <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Masker til jobnr og tilbudsnr:</h3>
            
            <%
                    if func = "exchopd" then
                    Response.write "<font color=red><b>Opdateret!</b><br><br></font>"
                    end if
                    %>            
                        
            Seneste <b>Job</b> havde nr: 
            
            <input id="FM_tsa_jobnr" name="FM_tsa_jobnr" style="width:100px;" type="text" value="<%=jobnr%>" /> (hel tal 0-1000000)
            <br />Kun ved joboprettelse tilf�jes der et jobnummer til denne maske. <br />
            Ved redigering af jobnummer p� et specifikt job, bliver denne hovedmaske ikke �ndret.<br />
            Jobnr er unikt.
         
            <br /><br>
            Seneste <b>Tilbud</b> havde nr: 
            <input id="FM_tsa_tilbudsnr" name="FM_tsa_tilbudsnr" style="width:100px;" type="text" value="<%=tilbudsnr%>" /> (hel tal 0-1000000)
            <br />
            Tilbuds nummer er unikt. 
            <br />
                &nbsp;
            
           
            </td>
            </tr>
           <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Stade-indmelding</h3>
           
          <%if jobasnvigv = 1 then
          jobasnvigvCHKL = "CHECKED"
          else
          jobasnvigvCHKL = ""
          end if
          %>

           <input type="checkbox" name="jobasnvigv" value="1" <%=jobasnvigvCHKL %> /> Stade-indmeldings funktion p� timereg. siden aktiv.
             </td>
            </tr>


       


              <tr>
                     <td style="padding:15px; border-top:1px #5582d2 dashed;">
                        <h3>Indstilllinger for joboprettesle:</h3></td>


              </tr>
        
               <tr>
            <td style="padding:15px;">
            <h3>Konto p� budget p� aktivitetslinjer og salgsomkostninger (job oprettelse)</h3>
           
          <%select case cint(budgetakt)
          case 0
          budgetAktCHK0 = "CHECKED"
          budgetAktCHK1 = ""
          budgetAktCHK2 = ""
          case 1
          budgetAktCHK0 = ""
          budgetAktCHK1 = "CHECKED"
          budgetAktCHK2 = ""
          case 2
          budgetAktCHK0 = ""
          budgetAktCHK1 = ""
          budgetAktCHK2 = "CHECKED"
          end select
          %>

           <input type="radio" name="budgetAkt" value="0" <%=budgetAktCHK0 %> /> Vis mulighed for at angive budget pr. aktivitet som Stk. / ingen konto p� salgsomkostninger (Produktion)<br />
            <input type="radio" name="budgetAkt" value="1" <%=budgetAktCHK1 %> /> Vis kontonr. pr. aktivitet og salgsomk. - dropdown kontoplan (Prof. services / NAV integration)<br />
              <input type="radio" name="budgetAkt" value="2" <%=budgetAktCHK2 %> /> Vis kontonr. pr. aktivitet og salgsomk. - �bent textfelt. (dynamisk kontoplan)  
             </td>
            </tr>
              
            <tr>
            <td style="padding:15px;">
            <h3>Budget p� medarbejdertype</h3>
           
          <%if cint(bdgmtypon) = 1 then
          bdgmtyponCHK = "CHECKED"
          else
          bdgmtyponCHK = ""
          end if
          %>

           <input type="checkbox" name="bdgmtypon" value="1" <%=bdgmtyponCHK %> /> Aktiver mulighed for at angive budget p� medarbejdertyper. (p� job)
             </td>
            </tr>




                    <tr>
            <td style="padding:15px;">
            <h3>Forretningsomr�de obligatorisk</h3>
           
          <%if cint(fomr_mandatory) = 1 then
          fomr_mandatoryCHK = "CHECKED"
          else
          fomr_mandatoryCHK = ""
          end if
          %>

           <input type="checkbox" name="fomr_mandatory" value="1" <%=fomr_mandatoryCHK %> /> Forretningsomr�de skal angives ved joboprettelse
             </td>
            </tr>


              <tr>
            <td style="padding:15px;">
            <h3>Budget obligatorisk</h3>
           
          <%if cint(budget_mandatory) = 1 then
          budget_mandatoryCHK = "CHECKED"
          else
          budget_mandatoryCHK = ""
          end if
          %>

           <input type="checkbox" name="budget_mandatory" value="1" <%=budget_mandatoryCHK %> /> Budget (nettooms�tning) skal angives ved joboprettelse. (skal v�re h�jere end 999,-)
             </td>
            </tr>

              <tr>
            <td style="padding:15px;">
            <h3>Salgsomkostninger / ulev. som popup ved joboprettelse</h3>
           
          <%if cint(show_salgsomk_mandatory) = 1 then
          show_salgsomk_mandatoryCHK = "CHECKED"
          else
          show_salgsomk_mandatoryCHK = ""
          end if
          %>

           <input type="checkbox" name="show_salgsomk_mandatory" value="1" <%=show_salgsomk_mandatoryCHK %> /> Vis <b>IKKE</b> Salgsomkostninger / ulev. som popup ved joboprettelse. (som stamaktiviteter)
             </td>
            </tr>

            

              <tr>
            <td style="padding:15px;">
            <h3>Job starter altid som et tilbud</h3>
           
          <%if cint(tilbud_mandatory) = 1 then
          tilbud_mandatoryCHK = "CHECKED"
          else
          tilbud_mandatoryCHK = ""
          end if
          %>

           <input type="checkbox" name="tilbud_mandatory" value="1" <%=tilbud_mandatoryCHK %> /> Ja, job starter altid som et tilbud ved joboprettelse.
             </td>
            </tr>

           <tr>
            <td style="padding:15px;">
            <h3>Salgsansvarlige</h3>
           
          <%if cint(salgsans) = 1 then
          salgsansCHK = "CHECKED"
          else
          salgsansCHK = ""
          end if
          %>

           <input type="checkbox" name="FM_salgsans" value="1" <%=salgsansCHK %> /> Aktiver mulighed for at angive salgsansvarlige p� job.
             </td>
            </tr>



               <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Kontonr. f�lger forrretningsomr�de eller s�t Kontonr p� aktivitet</h3>
            Standard indstillingen er at der kan knyttes en konto til et forretningsomr�de. (under forretningsomr�der i administrationen) <br />
            Herunder kan det v�lges at det bliver muligt at tilknytte et kontonummer pr. aktivitet (s�ttes i stam-aktivitets skabelonen) <br /><br />

           
          <%if cint(fomr_account) = 1 then
          fomr_accountCHK = "CHECKED"
          else
          fomr_accountCHK = ""
          end if
          %>

           <input type="checkbox" name="fomr_account" value="1" <%=fomr_accountCHK %> /> Tilknyt et kontonummer pr. aktivitet 
             </td>
            </tr>

             <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Rejseafregning - Di�ter</h3>
            Mulighed for at angive rejsedage og antal di�ter pr. dag. 
            <br /><br />

           
          <%if cint(traveldietexp_on) = 1 then
          traveldietexp_onCHK = "CHECKED"
          else
          traveldietexp_onCHK = ""
          end if
          %>

           <input type="checkbox" name="traveldietexp_on" value="1" <%=traveldietexp_onCHK %> /> Tilf�j mulighed for at angive rejsedage og di�ter.<br /><br />
           <input type="text" name="traveldietexp_maxhours" value="<%=traveldietexp_maxhours %>" style="width:60px;" /> Maks antal timer pr. dag p� dage med di�ter/rejsedage. <br /> (-1 = uendelig, kun type: der t�ller med i daglig timereg. se akt. typer)
                <br />G�lder IKKE Mobil og Ugeseddel.
             </td>
            </tr>


             <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Medarbejdertype f�lger medarbejder 1:1</h3>
           
          <%if cint(medarbtypligmedarb) = 1 then
          medarbtypligmedarbCHK = "CHECKED"
          else
          medarbtypligmedarbCHK = ""
          end if
          %>

           <input type="checkbox" name="medarbtypligmedarb" value="0" <%=medarbtypligmedarbCHK %> /> Medarbejdertype f�lger medarbejder 1:1.<br />
                 Medarbejdertype bliver step 2 i medarbejderoprettelse.
   
             </td>
            </tr>



             <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Teamleder / Flad struktur</h3>
           
          <%if cint(teamleder_flad) = 1 then
          teamlederCHK = "CHECKED"
          else
          teamlederCHK = ""
          end if
          %>

           <input type="checkbox" name="FM_teamleder_flad" value="1" <%=teamlederCHK %> /> Benyt flad struktur (alle kan se alle, ellers kan man kun se dem man er teamleder for, samt admin. kan se alle.)
             </td>
            </tr>

            <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Benyt positiv aktivitering af aktiviteter</h3>
           
          <%if cint(positiv_aktivering_akt) = 1 then
          positiv_aktivering_aktCHK = "CHECKED"
          else
          positiv_aktivering_aktCHK = ""
          end if
          %>

           <input type="checkbox" name="positiv_aktivering_akt" value="1" <%=positiv_aktivering_aktCHK %> /> Brug positiv aktivitering af aktiviteter p� timereg. siden.<br />

           <%
           'uTxt = "<b>Bem�rk:</b> ved �ndring af denne indstilling t�mmes alle personlige aktive joblister p� timereg. siden."
           'uWdt = 500
           'call infoUnisport(uWdt, uTxt) %>
           

           Der skal gives adgang til hver enkelt aktivitet for hver medarbejder f�r denne kan registrere timer p� aktiviteten. Der gives adgang fra de personlige Job-bank indstillinger fra timereg. siden. Alle der har adgang til medarbejderen via timereg. siden kan give adgang.
           <br /><br />Velegnet til virksomheder, med l�ngerevarende job og mange medarbejdere, hvor arbejdsomr�der er meget opdelt.


          <%if cint(pa_aktlist) = 1 then
          pa_aktlistCHK = "CHECKED"
          else
          pa_aktlistCHK = ""
          end if
          %>
            <br /><br />
           <input type="checkbox" name="pa_aktlist" value="1" <%=pa_aktlistCHK %> />Mobil (timetag_web), samt ugeseddel og ressourceforecast kan KUN at s�ge i <b>Personlig aktivjobliste</b> PA=1.
           Ellers kan der s�ges i hele jobbanken PA=0. Foruds�ttes at man har adgang via sine projektgrupper. (TimeTag s�ttes i timetag config)<br />

            </td>
            </tr>


             <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Upload/importer timer via excel funktion</h3>
           
          <%if cint(showupload) = 1 then
          showuploadCHK = "CHECKED"
          else
          showuploadCHK = ""
          end if
          %>

           <input type="checkbox" name="showupload" value="1" <%=showuploadCHK %> /> Vis upload / importer timer via excel ark. fra timereg. siden.
          </td>
            </tr>

          <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Benyt Komme/g� som basiskriterie</h3>
           
          <%if cint(week_showbase_norm_kommegaa) = 1 then
          week_showbase_norm_kommegaaCHK = "CHECKED"
          else
          week_showbase_norm_kommegaaCHK = ""
          end if
          %>

            <input type="checkbox" name="FM_week_showbase_norm_kommegaa" value="1" <%=week_showbase_norm_kommegaaCHK %> />  Benyt Komme/g� som basiskriterie for udregning af % afsluttet pr. dag. <br />
                Vises p� ugeseddel, Smiley afslut kriterie mm.
          </td>
            </tr>

          <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Mobil, Ugeseddel indstilling for dropdown/select felter</h3>
           
               
          <%if cint(mobil_week_reg_job_dd) = 1 then
          mobil_week_reg_job_ddCHK = "CHECKED"
          else
          mobil_week_reg_job_ddCHK = ""
          end if
          %>

         <%if cint(mobil_week_reg_akt_dd) = 1 then
          mobil_week_reg_akt_ddCHK = "CHECKED"
          else
          mobil_week_reg_akt_ddCHK = ""
          end if
          %>

          <%if cint(mobil_week_reg_akt_dd_forvalgt) = 1 then
          mobil_week_reg_akt_dd_forvalgtCHK = "CHECKED"
          else
          mobil_week_reg_akt_dd_forvalgtCHK = ""
          end if
          %>

         

           <input type="checkbox" name="FM_mobil_week_reg_job_dd" value="1" <%=mobil_week_reg_job_ddCHK %> />  Vis job s�gefelt som dropdown<br />
            <input type="checkbox" name="FM_mobil_week_reg_akt_dd" value="1" <%=mobil_week_reg_akt_ddCHK %> />  Vis aktivitet s�gefelt som dropdown<br />
            <input type="checkbox" name="FM_mobil_week_reg_akt_dd_forvalgt" value="1" <%=mobil_week_reg_akt_dd_forvalgtCHK %> />  1. Aktivitet forvalgt i dropdown<br />
             
            <input type="checkbox" name="FM_mobil_week_reg_dato" value="1" <%=mobil_week_reg_datoCHK %> DISABLED />  Vis datov�lger<br />
            <input type="checkbox" name="FM_mobil_week_reg_mat" value="1" <%=mobil_week_reg_matCHK %> DISABLED />  Vis materiale/udl�gs reg.<br />
          </td>
            </tr>


            
                 

             <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Timebudget simulering</h3>
           
          <%if cint(timesimon) = 1 then
          timesimonCHK = "CHECKED"
          else
          timesimonCHK = ""
          end if
          %>

          <%if cint(timesimh1h2) = 1 then
          timesimh1h2CHK = "CHECKED"
          else
          timesimh1h2CHK = ""
          end if
          %>

                <%if cint(timesimtp) = 1 then
          timesimtpCHK = "CHECKED"
          else
          timesimtpCHK = ""
          end if
          %>

           <input type="checkbox" name="timesimon" value="1" <%=timesimonCHK %> /> Aktiver timesimulering<br />
           <input type="checkbox" name="timesimh1h2" value="1" <%=timesimh1h2CHK %> /> Opdel finans�r i simulering i H1 og H2<br />
           <input type="checkbox" name="timesimtp" value="1" <%=timesimtpCHK %> /> Vis medarb. timepriser p� medarb. forecast 


             </td>
            </tr>
            

              <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Easyreg. funktion</h3>
           
          <%if cint(showeasyreg) = 1 then
          showeasyregCHK = "CHECKED"
          else
          showeasyregCHK = ""
          end if
          %>

           <input type="checkbox" name="showeasyreg" value="1" <%=showeasyregCHK %> /> Brug Easyreg. registrering p� timereg. siden. <br />
           Giver mulighed for at TimeOut selv fordeler dagens arbejde ud p� de forvalgte Easyreg. aktiviteter.
          </td>
            </tr>
            


                 <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Vis lager</h3>
           
          <%if cint(vis_lager) = 1 then
          vis_lagerCHK = "CHECKED"
          else
          vis_lagerCHK = ""
          end if
          %>

           <input type="checkbox" name="vis_lager" value="1" <%=vis_lagerCHK %> disabled /> Vis materialelager i menu<br />


                   <%if cint(minimumslageremail) = 1 then
          minimumslageremailCHK = "CHECKED"
          else
          minimumslageremailCHK = ""
          end if
          %>


                <input type="checkbox" name="FM_minimumslageremail" value="1" <%=minimumslageremailCHK %> /> Adviser materialelager-ansvarlige n�r lager er under minimumslager. (udsendes hver torsdag)
          </td>
            </tr>



              <tr>
            <td>
            
            
            <table cellpadding=0 cellspacing=0 border=0><tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h3>Aktiver Smileyordning (afslutning af dag / uger / md)</h3>  
            <%if smiley = 1 then
            smChk = "CHECKED"
            else
            smChk = ""
            end if %>
            <input type="checkbox" name="FM_smiley" id="FM_smiley" value="1" <%=smchk%>>
                <b>Aktiver Smiley ordning</b>
                
                <br>
                Hvis Smiley ordningen aktiveres bliver det muligt for en medarbjeder at afslutte
                uge, samt at modtage glade / sure Smiley's alt efter om man er "up to date" med
                sine timeregistreringer.


                <%
                    SmiWeekOrMonth0SEL = ""
                    SmiWeekOrMonth1SEL = ""
                    SmiWeekOrMonth2SEL = ""
                    
                select case SmiWeekOrMonth
                    case 0
                    SmiWeekOrMonth0SEL = "SELECTED"
                    case 1
                    SmiWeekOrMonth1SEL = "SELECTED"
                    case 2
                    SmiWeekOrMonth2SEL = "SELECTED"
                 end select
                    
                     


                    SmiantaldageCount1SEL = ""
                    SmiantaldageCount2SEL = ""
                    SmiantaldageCount3SEL = ""
                    SmiantaldageCount4SEL = ""
                    SmiantaldageCount5SEL = ""
                    SmiantaldageCount6SEL = ""
                    SmiantaldageCount7SEL = ""
                    SmiantaldageCount8SEL = ""
                    SmiantaldageCount9SEL = ""
                    SmiantaldageCount10SEL = ""
                    SmiantaldageCount11SEL = ""

                     select case SmiantaldageCount
                     case 1
                            SmiantaldageCount1SEL = "SELECTED"
                     case 2
                            SmiantaldageCount2SEL = "SELECTED"
                    case 3
                            SmiantaldageCount3SEL = "SELECTED"
                    case 4
                            SmiantaldageCount4SEL = "SELECTED"
                    case 5
                            SmiantaldageCount5SEL = "SELECTED"
                    case 6
                            SmiantaldageCount6SEL = "SELECTED"
                    case 7
                            SmiantaldageCount7SEL = "SELECTED"
                    case 8 
                            SmiantaldageCount8SEL = "SELECTED"
                    case 9 
                            SmiantaldageCount9SEL = "SELECTED"
                    case 10 
                            SmiantaldageCount10SEL = "SELECTED"
                    case 11 
                            SmiantaldageCount11SEL = "SELECTED"
                    end select

            
                    SmiantaldageCountClock7SEL = ""
                    SmiantaldageCountClock9SEL = ""
                    SmiantaldageCountClock10SEL = ""
                    SmiantaldageCountClock12SEL = ""
                    SmiantaldageCountClock17SEL = ""
                    SmiantaldageCountClock24SEL = ""      
                       
            select case SmiantaldageCountClock
                    case 7
                    SmiantaldageCountClock7SEL = "SELECTED"
                    case 9
                    SmiantaldageCountClock9SEL = "SELECTED"
                    case 10
                    SmiantaldageCountClock10SEL = "SELECTED"
                    case 12
                    SmiantaldageCountClock12SEL = "SELECTED"
                    case 17
                    SmiantaldageCountClock17SEL = "SELECTED"
                    case 24
                    SmiantaldageCountClock24SEL = "SELECTED"
                    case else
                    SmiantaldageCountClock12SEL = "SELECTED"    
            end select 

            
                    SmiTeamlederCountminus1SEL = ""
                    SmiTeamlederCount1SEL = ""
                    SmiTeamlederCount2SEL = ""
                    SmiTeamlederCount3SEL = ""
                    SmiTeamlederCount5SEL = ""
                    SmiTeamlederCount7SEL = ""
                    SmiTeamlederCount10SEL = ""
                    SmiTeamlederCount14SEL = ""

            select case SmiTeamlederCount 
                    case -1
                    SmiTeamlederCountminus1SEL = "SELECTED"  
            case 1
                    SmiTeamlederCount1SEL = "SELECTED"
                case 2
                    SmiTeamlederCount2SEL = "SELECTED"
                case 3
                    SmiTeamlederCount3SEL = "SELECTED"
                case 5
                    SmiTeamlederCount5SEL = "SELECTED"
                case 7
                    SmiTeamlederCount7SEL = "SELECTED"
                case 10
                    SmiTeamlederCount10SEL = "SELECTED"
                case 14
                    SmiTeamlederCount14SEL = "SELECTED"
            case else
                SmiTeamlederCount1SEL = "SELECTED"
        end select

                    %>
                </td></tr><tr><td style="padding:15px;">
                <b>Afslutning af perioder</b><br />
                    Perioder afsluttes p� 
                    <select name="FM_SmiWeekOrMonth">
                        <option value="0" <%=SmiWeekOrMonth0SEL %>>Ugebasis</option>
                        <option value="1" <%=SmiWeekOrMonth1SEL %>>M�nedsbasis</option>
                        <option value="2" <%=SmiWeekOrMonth2SEL %>>Dagligt</option>
                    </select>
                  
                    <br />
                    <br />
                    Perioden skal v�re afsluttet af medarbejderen den f�rst kommende 
                    <br /><br />
                     <select name="FM_SmiantaldageCount">
                         <option disabled>Ved m�nedsafslutning v�lg</option>
                         <option disabled>f�rst kommende:</option>
                        <option value="8" <%=SmiantaldageCount8SEL %>>Hverdag</option>
                         <option value="10" <%=SmiantaldageCount10SEL %>>Hverdag + 3 dage</option>
                         <option value="9" <%=SmiantaldageCount9SEL %>>Hverdag + 7 dage</option>
                         <option disabled>..i den efterf�lgende m�ned.</option>

                          <option disabled></option>
                          <option disabled>Ved ugeafslutning v�lg:</option>
                        <option value="1" <%=SmiantaldageCount1SEL %>>Mandag</option>
                        <option value="2" <%=SmiantaldageCount2SEL %>>Tirsdag</option>
                        <option value="3" <%=SmiantaldageCount3SEL %>>Onsdag</option>
                        <option value="4" <%=SmiantaldageCount4SEL %>>Torsdag</option>
                        <option value="5" <%=SmiantaldageCount5SEL %>>Fredag</option>
                         <option value="6" <%=SmiantaldageCount6SEL %>>l�rdag</option>
                         <option value="7" <%=SmiantaldageCount7SEL %>>S�ndag</option>
                          <option disabled>..i den efterf�lgende uge.</option>

                          <option disabled></option>
                          <option disabled>Ved daglig afslutning v�lg:</option>
                        <option value="11" <%=SmiantaldageCount11SEL %>>Efterf�lgende hverdag</option>

                    </select>
                   kl. <select name="FM_SmiantaldageCountClock">
                       <option value="7" <%=SmiantaldageCountClock7SEL%>>07:00</option>
                       <option value="9" <%=SmiantaldageCountClock9SEL%>>09:00</option>
                       <option value="10" <%=SmiantaldageCountClock10SEL%>>10:00</option>
                       <option value="12" <%=SmiantaldageCountClock12SEL%>>12:00</option>
                       <option value="17" <%=SmiantaldageCountClock17SEL%>>17:00</option>
                       <option value="24" <%=SmiantaldageCountClock24SEL%>>23:59</option>
                       
                       </select>
                    
                
                    <br /><br />
                    Perioden skal v�re godkendt / lukket af teamleder kl. 23:59 senest
                    <select name="FM_SmiTeamlederCount">
                        
                        <option value="1" <%=SmiTeamlederCount1SEL%>>1 dag</option>
                        <option value="2" <%=SmiTeamlederCount2SEL%>>2 dage</option>
                        <option value="3" <%=SmiTeamlederCount3SEL%>>3 dage</option>
                        <option value="5" <%=SmiTeamlederCount5SEL%>>5 dage</option>
                        
                        <option value="7" <%=SmiTeamlederCount7SEL%>>7 dage</option>
                     
                        <option value="10" <%=SmiTeamlederCount10SEL%>>10 dage</option>
                        <option value="14" <%=SmiTeamlederCount14SEL%>>14 dage</option>
                        <option value="-1" <%=SmiTeamlederCountminus1SEL%>>Aldrig</option>
                    </select> efter at medarbejderens skal have afsluttet perioden.




                    

                </td></tr>
                <tr><td style="padding:15px;">

            
                  <%if cint(smileyaggressiv) = 1 then
                    smileyAggressivChk = "CHECKED"
                    else
                    smileyAggressivChk = ""
                    end if %>

                     <%if cint(smiley_agg_lukhard) = 1 then
                    smiley_agg_lukhardChk = "CHECKED"
                    else
                    smiley_agg_lukhardChk = ""
                    end if %>
            <input type="checkbox" name="FM_smileyAggressiv" id="Checkbox2" value="1" <%=smileyAggressivChk%>> <b>Vis smiley aggressiv</b> (altid �ben) / diskret (lukket) p� timreg. siden
            <br />
            <input type="checkbox" name="FM_smiley_agg_lukhard" id="Checkbox2" value="1" <%=smiley_agg_lukhardChk%>> <b>H�rd</b> luk for registrering for medarbejder hvis peridoer ikke afsluttet.
            <br /><br />
                
                     - Timereg. siden bliver lukket hvis der er mere end 3 uafsluttede uger. (ved afslut p� ugebasis)<br />
                     - Timereg. siden bliver lukket hvis der er mere end 6 uafsluttede m�neder. (ved afslut p� m�nedsbasis)<br />
                     - Timereg. & komme/ g� bliver lukket hvis der er mere end 1 uafsluttet dag. (ved afslut daglit)<br />
             
               </td>
               </tr>


                   <tr><td style="padding:15px;">

            
                  <%if cint(hidesmileyicon) = 1 then
            hidesmileyiconChk = "CHECKED"
            else
            hidesmileyiconChk = ""
            end if %>
            <input type="checkbox" name="FM_hidesmileyicon" id="Checkbox2" value="1" <%=hidesmileyiconChk%>> <b>Skjul smiley ikoner</b> 
             
               </td>
               </tr>


               <tr>
               <td style="padding:15px;">
               
                    <%if autogk = 1 then
                    autogkChk = "CHECKED"
                    else
                    autogkChk = ""
                    end if %>
                
                    <input type="checkbox" name="FM_autogk" id="FM_autogk" value="1" <%=autogkchk%>>
                    <b>Luk periode for indtastning n�r ugen/m�neden bliver godkendt/lukket (af leder).</b>
                    <%
                    if func = "exchopd" then
                    Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
                    end if
                    %>
                    <br>
                    Denne funktion g�r at medarbejdere
                    ikke l�ngere kan indtaste/redigere timer i de uger de har afsluttet.<br />
                    <b>Administratorer kan stadigv�k �ndre i timer indtil der er oprettet en faktura p� jobbet.</b>
                    <br />Der kan godt indtastes materialeforbrug / udgiftsbilag selvom en periode er godkendt og lukket.
                   
                   </TD>
                   </tr>
                   
                <tr>
               <td style="padding:15px;">
               
                    <%if autogktimer = 1 then
                    autogktimerChk = "CHECKED"
                    else
                    autogktimerChk = ""
                    end if %>
                
                    <input type="checkbox" name="FM_autogktimer" id="Checkbox1" value="1" <%=autogktimerChk%>>
                    <b>Godkend automatisk timer n�r medarbejder afslutter uge/m�ned/dag via smiley.</b>
                    <%
                    if func = "exchopd" then
                    Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
                    end if
                    %>
                    <br>
                    Denne funktion g�r at medarbejdere
                    automatisk godkender deres egne timer n�r de afslutter deres uge/m�ned/dag via Smiley ordningen.
                   
                   </TD>
                   </tr>
                   <tr>
                   
                   
                   <%if autolukvdato = 1 then
                    autolukvdatoChk = "CHECKED"
                    else
                    autolukvdatoChk = ""
                    end if %>
                  
                
                   
                   <td style="padding:15px;">
                   <input type="checkbox" name="FM_autolukvdato" id="FM_autolukvdato" value="1" <%=autolukvdatoChk%>>
                   <b>M�nedsafslutning. </b> (Luk automatisk m�neder for indtastning.) <br />
                   Luk forrige m�ned for indtastninger d. 
                   <!--#include file="inc/weekselector_dag.asp"--> <!-- b -->
                   i den efterf�lgende m�ned.
                   <%
                    if func = "exchopd" then
                    Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
                    end if
                    %>
                    <br />Denne funktion g�r at medarbejdere
                    ikke l�ngere kan indtaste/redigere timer, efter datoen har passeret den valgte afslutningsdato i den efterf�lgende m�ned.
                    Indtastning bliver lukket uanset om medarbejder har godkendt uge.
                   
                    
                    <br />
                    <br /><br />
                    <b>
                    Regler for hvorn�r man kan redigere i indtastede timer og hvorn�r indtastning bliver lukket for redigering.
                    </b>
                    <table cellspacing=2 cellpadding=2 width=100% border=0>
                    <tr>
                        <td style="width:300px;"><b>Hvem m� godkende</b></td>
                         <td style="width:300px;"><b>Hvorn�r bliver timer indtastning lukket</b></td>
                       </tr>
                       <tr>
                         <td valign=top>- Teamledere for de grupper man er teamleder for<br />
                         - Medlemmer af Administrator brugergruppen</td>
                         <td valign=top>
                         - N�r uge godkendes via <b>Smiley</b>*<br />
                         - N�r m�ned lukkes via fast <b>M�nedsafslutning</b>*<br />
                         - N�r timer godkendes af Jobansvarlige eller administrator.<br />
                         - N�r der oprettes en faktura.
                        
                        
                        <br /><br /> 
                         <b>*)</b> Hvis luk uger ved godkendelse er sl�et til. 
                         Dog kan timer, indtil fakturering ell. godkendelse, stadigv�k �ndres af Administrator gruppen.
                         
                         
                         </td>
                    </tr>
                    </table>
                   
                   
                   </td></tr>
                   

               

               

                    <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h4>Luk for timereg. p� akt. hvis forecast/timebudget p� akt./medarb. overskreddet</h4> 
          
          <%if cint(forcebudget_onakttreg) = 1 then
          forcebudget_onakttregCHK = "CHECKED"
          else
          forcebudget_onakttregCHK = ""
          end if
          %>

           <%
          forcebudget_onakttreg_afgrSEL0 = "" 'Nej
          forcebudget_onakttreg_afgrSEL1 = "" 'Budget�r
          forcebudget_onakttreg_afgrSEL2 = "" 'MD    
          select case cint(forcebudget_onakttreg_afgr)
          case 1
          forcebudget_onakttreg_afgrSEL1 = "SELECTED"
          case 2
          forcebudget_onakttreg_afgrSEL2 = "SELECTED"
          case else
          forcebudget_onakttreg_afgrSEL0 = "SELECTED"
          end select

         if cint(forcebudget_onakttreg_filt_viskunmbgt) = 1 then
          forcebudget_onakttreg_filt_viskunmbgtCHK = "CHECKED"
          else
          forcebudget_onakttreg_filt_viskunmbgtCHK = ""
          end if
          %>

           <input type="checkbox" name="forcebudget_onakttreg" value="1" <%=forcebudget_onakttregCHK %> /> <b>Adviser</b> (vejledende) 
                 Markerer p� timeregistrering, ugeseddel og evt. TimeTag (PA:2), hvis <b>forecast pr. medarb.</b> er overskreddet. 
                 (Husk sl� vis ressourcetimer til p� akt. linje. - g�lder kun type: Fakturerbare - og ikke interne HR -1, -2 job)
                <br /><br />
                Kode 1: Forecast overskreddet, lyser�d markering. Der kan stadigv�k tastes. (tjekker afgr�nsning indenf. regnskab�rs / md.)<br />
                Kode 2: Ingen forecast angivet, gr� markering. Aktivitet kan ses men lukket for indtastning<br />
                 <span style="color:#999999;">
                Timereg. variable: resforecastMedOverskreddet<br /></span>
               <br />
                  <input type="checkbox" name="resforecastMedOverskreddet_tastok" value="1" <%=resforecastMedOverskreddettastokCHK %> DISABLED />Skal stadigv�k kunne taste i gr�-felter.<br /><br />
                

                 
          <%if cint(akt_maksforecast_treg) = 1 then
          akt_maksforecast_tregCHK = "CHECKED"
          else
          akt_maksforecast_tregCHK = ""
          end if
          %>

           <input type="checkbox" name="akt_maksforecast_treg" value="1" <%=akt_maksforecast_tregCHK %> /> <b>Adviser - H�rd</b> (kr�ver adviser vejledende er sl�et til)<br />
           Forecast pr. medarb. pr. aktivitet overskreddet. (kun fakturerbare)<br />
           Forecast kan ikke overskriddes, hvis der ikke er lagt forecast ind p� aktivitet kan der ikke tastes.<br />
           Hvis h�rd: gr� markering. Aktivitet kan ses men lukket for indtastning. (+ alert ved overskridelse)<br />
           
               <br /><br /><b>Afgr�nsning:</b> (g�lder b�de H�rd og Vejledende)
                 <select name="forcebudget_onakttreg_afgr"> 
                     <option value="0" <%=forcebudget_onakttreg_afgrSEL0 %>>Ingen afgr�nsning</option>
                     <option value="1" <%=forcebudget_onakttreg_afgrSEL1 %>>Afgr�ns indenfor regnskabs�r (se regnskabs�r)</option>
                     <option value="2" <%=forcebudget_onakttreg_afgrSEL2 %>>Afgr�ns indenfor m�ned</option>
                     </select>
                     <br /><br />

                 <input type="checkbox" name="forcebudget_onakttreg_filt_viskunmbgt" value="1" <%=forcebudget_onakttreg_filt_viskunmbgtCHK %> /> <b>Vis</b> kun aktiviteter med forecast pr. medarb. p�. - g�lder alle aktivitetstyper (viser altid interne job -1, -2)
                 <br /> Tjekker ikke periode - Ovenst�ende advisering vejl. / h�rd bestemmer om der kan tastes p� aktiviteten. <br />
                 
                 <!--<br />(admin kan sl� fra p� timereg. + kr�ver "Adviser.." ovenfor er sl�et til)-->
           
                
          <br /><br />
          
          
          <%if cint(akt_maksbudget_treg) = 1 then
          akt_maksbudget_tregCHK = "CHECKED"
          else
          akt_maksbudget_tregCHK = ""
          end if
          %>

           <input type="checkbox" name="akt_maksbudget_treg" value="1" <%=akt_maksbudget_tregCHK %> /> <b>Aktiver alert</b> p� timereg. siden, ved <b>forkalkuleret timer overskreddet (budgetlinje p� aktivitet) </b><br />G�lder alle aktivitetstyper med forkalk. p�. <br />- IKKE MOBIL og Ugeseddel<br />
                  
          </td>
            </tr>

             


                     <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h4>Regnskabs�r starter</h4> 
          
            Dato: <input id="Text6" name="FM_regnskabsaar_start" value="<%=day(regnskabsaar_start) %>-<%=month(regnskabsaar_start) %>" type="text" style="width:50px;" /> (dd-mm)</td>
            </tr>


              <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h4>Aktivitetslinjer  (vis avanceret ell. simpel)</h4> 

                Indstilling for aktivitetslinjer p� timereg. siden<br />
                Interne job, bliver altid vist "simpel".<br />
          
          <%if cint(visAktlinjerSimpel) = 1 then
          visAktlinjerSimpelCHK = "CHECKED"
          else
          visAktlinjerSimpelCHK = ""
          end if


          if cint(visAktlinjerSimpel_datoer) = 1 then
          visAktlinjerSimpel_datoerCHK = "CHECKED"
          else
          visAktlinjerSimpel_datoerCHK = ""
          end if

         
          if cint(visAktlinjerSimpel_timebudget) = 1 then
          visAktlinjerSimpel_timebudgetCHK = "CHECKED"
          else
          visAktlinjerSimpel_timebudgetCHK = ""
          end if


           if cint(visAktlinjerSimpel_realtimer) = 1 then
          visAktlinjerSimpel_realtimerCHK = "CHECKED"
          else
          visAktlinjerSimpel_realtimerCHK = ""
          end if


          if cint(visAktlinjerSimpel_restimer) = 1 then
          visAktlinjerSimpel_restimerCHK = "CHECKED"
          else
          visAktlinjerSimpel_restimerCHK = ""
          end if

         if cint(visAktlinjerSimpel_medarbtimepriser) = 1 then
          visAktlinjerSimpel_medarbtimepriserCHK = "CHECKED"
          else
          visAktlinjerSimpel_medarbtimepriserCHK = ""
          end if

         if cint(visAktlinjerSimpel_medarbrealtimer) = 1 then
          visAktlinjerSimpel_medarbrealtimerCHK = "CHECKED"
          else
          visAktlinjerSimpel_medarbrealtimerCHK = ""
          end if


         if cint(visAktlinjerSimpel_akttype) = 1 then
          visAktlinjerSimpel_akttypeCHK = "CHECKED"
          else
          visAktlinjerSimpel_akttypeCHK = ""
          end if
          %>

                

         

           <input type="checkbox" name="FM_visAktlinjerSimpel" value="1" <%=visAktlinjerSimpelCHK %> /> Vis simpel (vis ikke detaljer om timeforbrug, forecast, start- og slut -dato)<br />

                <br /><br />
                Eller v�lg hvilke data der skal vises  p� aktivitetslinjerne p� timereg. siden:<br />

                 <img src="../ill/blank.gif" width="20" height="1" /><input type="checkbox" name="FM_visAktlinjerSimpel_datoer" value="1" <%=visAktlinjerSimpel_datoerCHK %> /> Vis start og slutdato<br />
                <img src="../ill/blank.gif" width="20" height="1" /><input type="checkbox" name="FM_visAktlinjerSimpel_timebudget" value="1" <%=visAktlinjerSimpel_timebudgetCHK %> /> Vis akt. forkalkulation (hvis findes) <br />
                <img src="../ill/blank.gif" width="20" height="1" /><input type="checkbox" name="FM_visAktlinjerSimpel_realtimer" value="1" <%=visAktlinjerSimpel_realtimerCHK %> /> Vis real. timer<br />
                <img src="../ill/blank.gif" width="20" height="1" /><input type="checkbox" name="FM_visAktlinjerSimpel_restimer" value="1" <%=visAktlinjerSimpel_restimerCHK %> /> Vis ressourcetimer (kr�ver mark�r akt. hvis forecast overskreddet)<br />


                <img src="../ill/blank.gif" width="20" height="1" /><input type="checkbox" name="FM_visAktlinjerSimpel_medarbtimepriser" value="1" <%=visAktlinjerSimpel_medarbtimepriserCHK %> /> Medarb. timepriser<br />
                <img src="../ill/blank.gif" width="20" height="1" /><input type="checkbox" name="FM_visAktlinjerSimpel_medarbrealtimer" value="1" <%=visAktlinjerSimpel_medarbrealtimerCHK %> /> Medarb. real. timer<br />
                <img src="../ill/blank.gif" width="20" height="1" /><input type="checkbox" name="FM_visAktlinjerSimpel_akttype" value="1" <%=visAktlinjerSimpel_akttypeCHK %> /> Akt. type<br />

                </td>
            </tr>


                <tr>
            <td style="padding:15px; border-top:1px #5582d2 dashed;">
            <h4>Datosp�rring - vis ikke job & aktiviteter f�r startdato og efter slutdato</h4> 
          
          <%
        lukaktvdatoSel0 = ""
        lukaktvdatoSel1 = ""
        lukaktvdatoSel2 = ""
        lukaktvdatoSel3 = ""
        lukaktvdatoSel4 = ""  
                  
          select case lukaktvdato
          case 0 
          lukaktvdatoSel0 = "SELECTED"
          case 1 
          lukaktvdatoSel1 = "SELECTED"
          case 2 
          lukaktvdatoSel2 = "SELECTED"
          case 3 
          lukaktvdatoSel3 = "SELECTED"
          case 4 
          lukaktvdatoSel4 = "SELECTED"
          case else
          lukaktvdatoSel0 = "SELECTED"
          end select
          %>

        <select name="FM_lukaktvdato" style="width:600px;" size="5">
            <option value="0" <%=lukaktvdatoSel0 %>>0. L�s job, �ben aktiv: Datosp�rring p� startdato job (default)</option>
            <option value="1" <%=lukaktvdatoSel1 %>>1. L�s job, H�rd aktiv: Datosp�rring p� startdato job og aktiviteter start- og slut -dato</option>
            <option value="2" <%=lukaktvdatoSel2 %>>2. �ben job, H�rd aktiv: Datosp�rring p� aktiviteter</option>
            <option value="3" <%=lukaktvdatoSel3 %>>3. H�rd job, H�rd aktiv: Datosp�rring p� job og aktiviteter</option>
            <option value="4" <%=lukaktvdatoSel4 %>>4. �ben: Ingen datosp�rring p� job og aktiviteter</option>
        </select>
          <!--<input type="checkbox"  value="1" <%=lukaktvdatoCHK %> /> Vis ikke job & aktiviteter p� timereg. siden f�r startdato er oprindet, og vis ikke aktiviteter n�r slutdato er passeret.-->
           <br />
                Ovenst�ende indstilllinger g�lder ikke salgs-aktiviteter og tilbud (job). Disse der altid kan ses n�r aktive. <br /><br />
            <!--    0. �ben: Ingen datosp�rring p� job og aktiviteter. (n�r job og aktiviteter kan de ses n�r de er aktive)<br />
                1. L�s job, H�rd aktiv: Datosp�rring p� job og aktiviteter. (job kan ses n�r startdato er oprindet, aktiviteter kan ses n�r start og slutdato er �bne p� registreringsdato)<br />
                2. �ben job, H�rd aktiv: Datosp�rring p� aktiviteter. (aktiviteter kan ses n�r start og slutdato er �bne p� registreringsdato)<br />
                3. H�rd job, H�rd aktiv: Datosp�rring p� job og aktiviteter. (job kan ses n�r startdato er oprindet og fjernes n�r slutdato er passeret, aktiviteter kan ses n�r start og slutdato er �bne p� registreringsdato)<br />
                4. H�rd job, �ben aktiv:  Datosp�rring p� job. (job kan ses n�r startdato er oprindet og fjernes n�r slutdato er passeret, aktiviteter kan ses n�r aktive)
                -->
               
          </td>
            </tr>     
                   
                    <tr>
                    <td style="padding:15px; border-top:1px #5582d2 dashed;">
                     <!--
                    <h3>F�rdigmelding af job.</h3> 
                    
                    <%if lukafm = 1 then
                    lukafmChk = "CHECKED"
                    else
                    lukafmChk = ""
                    end if %>
                    
                    <input type="checkbox" name="FM_lukafm" id="FM_lukafm" value="1" <%=lukafmchk%>>
                    <b>Aktiver f�rdigmelding af job.</b>
                    
                    <%
                    if func = "exchopd" then
                    Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
                    end if
                    %>
                    <br>
                    Aktiver mulighed for at medarbejdere, ved timeindtastning, selv kan lukke og f�rdigmelde et job. Ved lukning af job vil der blive sendt en email til jobansvarlige.

                    -->
                    
                    
                    <%if kmDialog = 1 then
                    kmDialogChk = "CHECKED"
                    else
                    kmDialogChk = ""
                    end if %>
                    
                    <h3>Aktiver "KM til og fra" boks p� timereg. siden</h3>

                    <input type="checkbox" name="FM_kmDialog" id="FM_kmDialog" value="1" <%=kmDialogchk%>>
                    <b>Aktiver Km dialogboks med adresser, ved registrering af k�rsel.</b><br />
                    Tilf�jer automatisk adresser ved km registrering, samt opretter k�rebogs-log og regnskab over 60 dages regel til dokumentation for SKAT.
                    
                    </td>
                    </tr>

                    <tr>
                    <td style="padding:15px; border-top:1px #5582d2 dashed;">
                    <h3>Km pris (global)</h3> 
                     <input type="checkbox" name="FM_kmpris_ja" value="1" /><b>Km pris</b> (opdater alle aktive KM aktiviteter)<br />
                    <input id="Text2" name="FM_kmpris" value="<%=kmpris %>" type="text" style="width:50px;" />
                      <input id="Text5" name="FM_kmpris_old" value="<%=kmpris %>" type="hidden"/>
                    Opdater Km pris fra dato: <input id="Hidden1" name="FM_kmpris_fra" value="<%=formatdatetime(now,2) %>" type="text" style="width:100px;" /> p� eksisterende timeregistreringer.

                    </td></tr>

                    <tr>
                    <td style="padding:15px; border-top:1px #5582d2 dashed;">
                    <h3>Afrund timer</h3> 
                    <input type="checkbox" name="FM_timerround" value="1" <%=timerroundCHK %> />Afrund altid til minimun � timer. Afrunder altid opdad s�ledes at 1,15 bliver til 1,5 og 1,73 bliver til 2. 
                    <br />G�lder kun p� fakturerbare og ikke fakturerbare aktiviteter.
                     
                   
                    </td></tr>

                      <tr>
                    <td style="padding:15px; border-top:1px #5582d2 dashed;">
                    <h3>Faktor (global)</h3> 
                    <input type="checkbox" name="FM_globalfaktor_ja" value="1" /><b>Faktor </b> - opdater faktor p� alle <b>aktive</b> fakturerbare aktiviteter, 
                    s� hver time bliver omregnet til enheder (timer p� faktura) via denne faktor:<br /> <input id="Text3" name="FM_globalfaktor" value="<%=globalfaktor %>" type="text" style="width:50px;" /> (1,25) kommatal 2 decimaler
                     <input id="Text4" name="FM_globalfaktor_old" value="<%=globalfaktor %>" type="hidden" />
                     
                     
                   
                    </td></tr>
                
                    
                    <tr>
                    <td style="padding:15px; border-top:1px #5582d2 dashed;">
                    <h3>Stempelur funktion (Komme/G� tid) <span style="font-size:11px; font-weight:lighter;"><br />Hvis stempelur funktionen aktiveres bliver det muligt at f�lge medarbejdernes
                        komme/g� tid og fleks i forhold til normtid mm.</span></h3> 

                         
                    
                    <%if stempelur = 1 then
                    stempelurChk = "CHECKED"
                    session("stempelur") = 1
                    else
                    stempelurChk = ""
                    session("stempelur") = 0
                    end if %>
                   
                    
                        <input type="checkbox" name="FM_stempelur" id="FM_stempelur" value="1" <%=stempelurchk%>>

                        <b>Aktiver Stempelur.</b>&nbsp;(<a href="javascript:NewWin_help('stempelur.asp?menu=tok&ketype=e');"
                            target="_self" class='vmenu'>Stempelur indstillinger</a>)
                        <%
                        if func = "exchopd" then
                        Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
                        end if
                        %>


                        <%  
                            if cint(stempelur_hidelogin) <> 0 then
                            stempelur_hideloginCHK = "CHECKED"
                            else
                            stempelur_hideloginCHK = ""
                            end if %>

                        <br /><br />
                           <input type="checkbox" name="FM_stempelur_hidelogin" id="FM_stempelur_hidelogin" value="1" <%=stempelur_hideloginCHK%>> Pre-udfyld <u>IKKE</u> Komme/G� siden med den tid man faktisk logger ind. (intast komme/g� tid manuelt) 


                        <%  
                            if cint(stempelur_igno_komkrav) <> 0 then
                            stempelur_igno_komkravCHK = "CHECKED"
                            else
                            stempelur_igno_komkravCHK = ""
                            end if %>

                         <br /><br />
                           <input type="checkbox" name="FM_stempelur_igno_komkrav" id="FM_stempelur_igno_komkrav" value="1" <%=stempelur_igno_komkravCHK%>> Fravig krav om kommentar ved �ndret Komme/G� tid. 


                        
                        
                      <br /><br />
                        
                       <b>Der skal ikke kunne logges ind mellem:</b><br /> 
                       Hvis der logges ind i dette interval, hvil logind tidspunktet blive sat til sluttiden i nedenst�ende interval. (de r�de kasser)<br /> 
                       Der medregnes ikke minutter, for logind hvor logind tidspunkt ligger i det angivne interval. (g�lder kun fremtidige logind)
                       <br /> 
                        <input id="FM_stempel_ignorerper_st_t" name="FM_stempel_ignorerper_st_t" type="text" style="width: 23px;" style="border:1px silver solid;" value="<%= ignorertid_st_t%>" />:
                       <input id="FM_stempel_ignorerper_st_m" name="FM_stempel_ignorerper_st_m" type="text" style="width: 23px;" style="border:1px silver solid;" value="<%= ignorertid_st_m%>" />
                       -
                        <input id="FM_stempel_ignorerper_sl_t" name="FM_stempel_ignorerper_sl_t" type="text" style="width: 23px;" style="border:1px red solid;" value="<%= ignorertid_sl_t%>" />:
                       <input id="FM_stempel_ignorerper_sl_m" name="FM_stempel_ignorerper_sl_m" type="text" style="width: 23px;" style="border:1px red solid;" value="<%= ignorertid_sl_m%>" />
                      <br />00:00 - 00:00 Ingen periode<br />
                      06:45 - 07:00 Hvis der f.eks logges ind kl 06:50, bliver logind tidspunktet sat til kl 07:00.<br />
                      Logges der derimod ind kl. 02:00, bi-beholdes logind tidspunket til kl. 02:00.
                      
                      <br /><br />
                      <b>Standard pause A pr. dag:</b><br />
                      <input id="FM_stempel_standardpause_A" name="FM_stempel_standardpause_A" type="text" style="width: 23px;" value="<%= stpauseA%>" /> min. pr. dag (0-55 min, 5 min. interval)<br />
                      <br />Tilf�j pause A p� f�lgende dage: <span style="color:crimson;">Hvis der er dage uden pause A, m� pause B ikke v�re sl�et til p� disse dage.</span><br />

                         Man 
                        <input id="p1_man" name="p1_man" value="1" type="checkbox" <%=p1_manChk %> />&nbsp;&nbsp;
                         Tir
                        <input id="p1_tir" name="p1_tir" value="1"type="checkbox" <%=p1_tirChk %> />&nbsp;&nbsp;
                         Ons 
                        <input id="p1_ons" name="p1_ons" value="1" type="checkbox" <%=p1_onsChk %> />&nbsp;&nbsp;
                         Tor 
                        <input id="p1_tor" name="p1_tor" value="1" type="checkbox" <%=p1_torChk %> />&nbsp;&nbsp;
                         Fre 
                        <input id="p1_fre" name="p1_fre" value="1" type="checkbox" <%=p1_freChk %> />&nbsp;&nbsp;
                         L�r 
                        <input id="p1_lor" name="p1_lor" value="1" type="checkbox" <%=p1_lorChk %> />&nbsp;&nbsp;
                         S�n 
                        <input id="p1_son" name="p1_son" value="1" type="checkbox" <%=p1_sonChk %> />&nbsp;&nbsp;
                       
                       <br /><br />
                       Standard pause A skal kun g�lde / ikke g�lde (minus) for f�lgende projektgrupper: (blank = alle, ellers angiv projektgruppe id)<br />
                       Det er ikke tilladt at kombinere minus-grupper med plus-grupper. <br /> <input type="text" name="FM_p1_grp" value="<%=p1_grp %>" size=10 /> Angiv gerne flere grupper 1,5,19 
                       <br />

                       <br /><br />
                      <b>Standard pause B pr. dag:</b><br />
                      <input id="FM_stempel_standardpause_B" name="FM_stempel_standardpause_B" type="text" style="width: 23px;" value="<%= stpauseB%>" /> min. pr. dag (0-55 min, 5 min. interval)
                    
                       <br /><br />Tilf�j pause B p� f�lgende dage:<br />
                         Man 
                        <input id="p2_man" name="p2_man" value="1" type="checkbox" <%=p2_manChk %> />&nbsp;&nbsp;
                         Tir
                        <input id="p2_tir" name="p2_tir" value="1" type="checkbox" <%=p2_tirChk %> />&nbsp;&nbsp;
                         Ons 
                        <input id="p2_ons" name="p2_ons" value="1" type="checkbox" <%=p2_onsChk %> />&nbsp;&nbsp;
                         Tor  
                        <input id="p2_tor" name="p2_tor" value="1" type="checkbox" <%=p2_torChk %> />&nbsp;&nbsp;
                         Fre 
                        <input id="p2_fre" name="p2_fre" value="1" type="checkbox" <%=p2_freChk %> />&nbsp;&nbsp;
                         L�r 
                        <input id="p2_lor" name="p2_lor" value="1" type="checkbox" <%=p2_lorChk %> />&nbsp;&nbsp;
                         S�n  
                        <input id="p2_son" name="p2_son" value="1" type="checkbox" <%=p2_sonChk %> />&nbsp;&nbsp;
                                        
                         
                        <br />  <br />
                       Standard pause B skal kun g�lde / ikke g�lde (minus) for f�lgende projektgrupper: (blank = alle, ellers angiv projektgruppe id)<br />
                       Det er ikke tilladt at kombinere minus-grupper med plus-grupper. <br /><input type="text" name="FM_p2_grp" value="<%=p2_grp %>" size=10 /> Angiv gerne flere grupper 1,5,19
                       <br />               
                       
                   
                    </td></tr>
                    </table>


                    <br /><br />
                         <h3>Virksomhedens �bningstider:</h3>
                    
                       �bningstiderne bruges af <b>stempeluret</b> ved glemt logud fra medarbejder og i ServiceDesk (hvis tilvalgt) til beregning af responstider.<br /><br />
                            <table cellspacing=1 cellpadding=2 border=0 bgcolor="#5582d2">
                                <tr>
                                    <td width=40 bgcolor="#ffffff" align="center">
                                        <b>Man:</b></td>
                                    <td bgcolor="#ffffff" style="width: 259px">
                                        <input id="FM_abn_man_t" name="FM_abn_man_t" type="text" style="width: 23px;" value="<%= man_t%>" />:
                                        <input id="FM_abn_man_m" name="FM_abn_man_m" type="text" style="width: 23px;" value="<%= man_m%>" />
                                        til
                                        <input id="FM_abn_man_t2" name="FM_abn_man_t2" type="text" style="width: 23px;" value="<%= man_t2%>" />:
                                        <input id="FM_abn_man_m2" name="FM_abn_man_m2" type="text" style="width: 23px;" value="<%= man_m2%>" />
                                   &nbsp;&nbsp;&nbsp;Eks: 08:15 - 17:00</td>
                                </tr>
                                <tr>
                                    <td bgcolor="#ffffff" align="center">
                                        <b>Tir:</b>
                                    </td>
                                    <td bgcolor="#ffffff" style="width: 259px">
                                        <input id="FM_abn_tir_t" name="FM_abn_tir_t" type="text" style="width: 23px;" value="<%=tir_t%>"/>:
                                        <input id="FM_abn_tir_m" name="FM_abn_tir_m" type="text" style="width: 23px;" value="<%=tir_m%>"/>
                                        til
                                        <input id="FM_abn_tir_t2" name="FM_abn_tir_t2" type="text" style="width: 23px;" value="<%=tir_t2%>" />:
                                        <input id="FM_abn_tir_m2" name="FM_abn_tir_m2" type="text" style="width: 23px;" value="<%=tir_m2%>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td bgcolor="#ffffff" align="center">
                                        <b>Ons:</b>
                                    </td>
                                    <td bgcolor="#ffffff" style="width: 259px">
                                        <input id="FM_abn_ons_t" name="FM_abn_ons_t" type="text" style="width: 23px;" value="<%=ons_t%>" />:
                                        <input id="FM_abn_ons_m" name="FM_abn_ons_m" type="text" style="width: 23px;" value="<%=ons_m%>" />
                                        til
                                        <input id="FM_abn_ons_t2" name="FM_abn_ons_t2" type="text" style="width: 23px;" value="<%=ons_t2%>" />:
                                        <input id="FM_abn_ons_m2" name="FM_abn_ons_m2" type="text" style="width: 23px;" value="<%=ons_m2%>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td bgcolor="#ffffff" align="center">
                                        <b>Tor:</b>
                                    </td>
                                    <td bgcolor="#ffffff" style="width: 259px">
                                        <input id="FM_abn_tor_t" name="FM_abn_tor_t" type="text" style="width: 23px;" value="<%=tor_t%>" />:
                                        <input id="FM_abn_tor_m" name="FM_abn_tor_m" type="text" style="width: 23px;" value="<%=tor_m%>" />
                                        til
                                        <input id="FM_abn_tor_t2" name="FM_abn_tor_t2" type="text" style="width: 23px;" value="<%=tor_t2%>" />:
                                        <input id="FM_abn_tor_m2" name="FM_abn_tor_m2" type="text" style="width: 23px;" value="<%=tor_m2%>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td bgcolor="#ffffff" align="center">
                                        <b>Fre:</b>
                                    </td>
                                    <td bgcolor="#ffffff" style="width: 259px">
                                        <input id="FM_abn_fre_t" name="FM_abn_fre_t" type="text" style="width: 23px;" value="<%=fre_t%>" />:
                                        <input id="FM_abn_fre_m" name="FM_abn_fre_m" type="text" style="width: 23px;" value="<%=fre_m%>" />
                                        til
                                        <input id="FM_abn_fre_t2" name="FM_abn_fre_t2" type="text" style="width: 23px;" value="<%=fre_t2%>" />:
                                        <input id="FM_abn_fre_m2" name="FM_abn_fre_m2" type="text" style="width: 23px;" value="<%=fre_m2%>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td bgcolor="#cccccc" align="center">
                                        <b>L�r:</b>
                                    </td>
                                    <td bgcolor="#cccccc" style="width: 259px">
                                        <input id="FM_abn_lor_t" name="FM_abn_lor_t" type="text" style="width: 23px;" value="<%=lor_t%>" />:
                                        <input id="FM_abn_lor_m" name="FM_abn_lor_m" type="text" style="width: 23px;" value="<%=lor_m%>" />
                                        til
                                        <input id="FM_abn_lor_t2" name="FM_abn_lor_t2" type="text" style="width: 23px;" value="<%=lor_t2%>" />:
                                        <input id="FM_abn_lor_m2" name="FM_abn_lor_m2" type="text" style="width: 23px;" value="<%=lor_m2%>" />
                                   &nbsp;&nbsp;&nbsp;Hvis lukket: 00:00 - 00:00 </td>
                                </tr>
                                <tr>
                                    <td bgcolor="#cccccc" align="center">
                                        <b>S�n:</b>
                                    </td>
                                    <td bgcolor="#cccccc" style="width: 259px">
                                        <input id="FM_abn_son_t" name="FM_abn_son_t" type="text" style="width: 23px;" value="<%=son_t%>" />:
                                        <input id="FM_abn_son_m" name="FM_abn_son_m" type="text" style="width: 23px;" value="<%=son_m%>" />
                                        til
                                        <input id="FM_abn_son_t2" name="FM_abn_son_t2" type="text" style="width: 23px;" value="<%=son_t2%>" />:
                                        <input id="FM_abn_son_m2" name="FM_abn_son_m2" type="text" style="width: 23px;" value="<%=son_m2%>" />
                                    </td>
                                </tr>
                            </table>

                           


                    <br /><br />
            <input type="image" src="../ill/opdaterpil.gif">
            <br /><br /></table>
            <img src="../ill/blank.gif" width="1" height="1" alt="" border="0">
                    </div>
                </td>
            </tr>
            </table>
            <table cellpadding=0 cellspacing=0 border=0 width=720>
            <tr>
                <td width=520 bgcolor="#ffffff" style="border:1px #5582d2 solid; padding:10px; border-right:0px;">
                <a href="javascript:ShowHide('kerp');"><b>ERP-Modul: (Fakturering og bogf�ring)</b></a></td>
                  <td width=200 align="right" bgcolor="#ffffff" style="border:1px #5582d2 solid; border-left:0px; padding:10px;">
              Aktiver modul <b>*</b>:<input type="checkbox" name="CHKerp" align="right" <%=Strerpstat%>/></td>
                      </tr>
           <tr>
            <td style="padding:10px;" colspan="2"><div id="kerp" name="kerp" style="visibility: hidden; display: none">
            <h3>Masker til faktura skrivelser:</h3>
            
            <%
                    if func = "exchopd" then
                    Response.write "<font color=red><b>Opdateret!</b><br><br></font>"
                    end if
                    %>    
                
                      
                        
            Seneste <b>Faktura</b> havde nr: 
            <input id="FM_erp_fakturanr" name="FM_erp_fakturanr" style="width:100px;" type="text" value="<%=fakturanr%>" /><br />hel tal, maks 10 karakterer. 
            Ved brug af FI nummer m� faktura nummer maks v�re 6 karakterer.
            
            <!--
            <br /><br>
            Seneste <b>Rykker</b> havde nr: 
            <input id="FM_erp_rykkernr" name="FM_erp_rykkernr" style="width:100px;" type="text" value="<%=rykkernr%>" /> (hel tal 0-1000000)
            -->
            <br /><br>
            Seneste <b>Kreditnota</b> havde nr: 
            <input id="FM_erp_kreditnr" name="FM_erp_kreditnr" style="width:100px;" type="text" value="<%=kreditnr%>" /> (hel tal, maks 10 karakterer)<br />
            Hvis faktura- og kreditnota -nr r�kkef�lge skal v�re en og samme nummerserie, angiv da -1 i Kreditnota feltet. 
            
            <br /><br />
            Seneste <b>Fakturakladde (interne/handels fakturaer)</b> havde nr: 
            <input id="Text1" name="FM_erp_fakturanr_kladde" style="width:100px;" type="text" value="<%=fakturanr_kladde%>" /><br />hel tal, maks 10 karakterer. 
            
             <br /><br />
            <span style="background-color:lightpink;">
            <b>V�r opm�rksom p� at nummerr�kkef�lger IKKE overlapper hinanden, da I s� vil f� en fejl ved faktura oprettelse.</b></span>
            
            <%if cint(multible_licensindehavere) = 1 then
                multible_licensindehavereCHK = "CHECKED"
             else
                multible_licensindehavereCHK = ""
             end if %>

            <br /><br />
            <input type="checkbox" name="FM_multible_licensindehavere" value="1" <%=multible_licensindehavereCHK%>/><b>Tillad multible juridiske enheder</b> afsender fakturaer fra TimeOut.<br />

            <%for f = 2 to 5%>
            <br />                    
            Licens indehaver <%=f %>
            <br />

                <%select case f
                  case 2
                    fakturanr = fakturanr_2
                    kreditnr = kreditnr_2
                    fakturanr_kladde = fakturanr_kladde_2
                   case 3
                     fakturanr = fakturanr_3
                    kreditnr = kreditnr_3
                    fakturanr_kladde = fakturanr_kladde_3
                    case 4
                     fakturanr = fakturanr_4
                    kreditnr = kreditnr_4
                    fakturanr_kladde = fakturanr_kladde_4
                    case 5
                     fakturanr = fakturanr_5
                    kreditnr = kreditnr_5
                    fakturanr_kladde = fakturanr_kladde_5
                   end select %>

              Seneste <b>Faktura</b> havde nr: 
            <input name="FM_erp_fakturanr_<%=f%>" style="width:100px;" type="text" value="<%=fakturanr%>" />
         
                <br />
            Seneste <b>Kreditnota</b> havde nr: 
            <input name="FM_erp_kreditnr_<%=f%>" style="width:100px;" type="text" value="<%=kreditnr%>" /> 
            
            <br />
            Seneste <b>Fakturakladde (interne/handels fakturaer)</b> havde nr: 
            <input  name="FM_erp_fakturanr_kladde_<%=f%>" style="width:100px;" type="text" value="<%=fakturanr_kladde%>" />
            
                    <br />
                    
                <%next %>

             
            <br>
            <br /><br />
             <ul>
             <li><a href="erp_valutaer.asp" class=vmenu target=_blank>Valutaer og valutakurser</a>
             </ul>
             <!--<br /><br />
                <input id="FM_md_fak" name="FM_md_fak" value="1" type="checkbox" /> Vi benytter m�nedsfakturering, s� dato interval ved fakturaoprettelse skal altid st� til forrige fulde m�ned.
             -->
             <br /><br />
            <input type="image" src="../ill/opdaterpil.gif">
            </div>
           </td>
            </table>
            </tr>
            </table>
           
            
             <table cellspacing=0 cellpadding=0 border=0 width=720>
                <tr>
                 <td bgcolor="#ffffff" width=520 style="border:1px #5582d2 solid; padding:10px; border-right:0px;">
                   <a href="javascript:ShowHide('kcrm');"><b>CRM-Modul:</b></a></td>
                <td width=200 align="right" bgcolor="#ffffff" style="border:1px #5582d2 solid; border-left:0px; padding:10px;">
                
              Aktiver modul <b>*</b>:<input type="checkbox" name="CHKcrm" align="right" <%=Strcrmstat%>/></td>
                </tr>
                
                <tr>
                <td style="padding:10px;" colspan="2"><div id="kcrm" name="kcrm" style="visibility: hidden; display: none">
                   




            <input type="image" src="../ill/opdaterpil.gif">
            <br /><br /></div>
                </td>
               </tr>
               
               
             </table>
            <table cellspacing=0 cellpadding=0 border=0 width=720>
            <tr>
                <td width=520 bgcolor="#ffffff" style="border:1px #5582d2 solid; border-right:0px; padding:10px;">
                <a href="javascript:ShowHide('ksdsk');"><b>SDSK-Modul (Servicedesk):</b></a></td>
              <td align="right" bgcolor="#ffffff" style="border:1px #5582d2 solid; border-left:0px; padding:10px;">Aktiver modul <b>*</b>:
              <input type="checkbox" name="CHKsdsk" align="right" <%=Strsdskstat%>/></td>
              </tr>
              <tr><td style="padding:10px;" colspan="2"><div id="ksdsk" name="ksdsk" style="visibility: hidden; display: none">
                   
                   
                            
                            
                           <%if brugabningstid = 1 then
                           baCHK = "CHECKED"
                           else
                           baCHK = ""
                           end if %>
                           
                            <input name="FM_brug_abningstid" id="FM_brug_abningstid" type="checkbox" value="1" <%=baCHK%>/>
                            <b>Brug �bningstider</b> (s�ttes under TSA) 
                            <%
                            if func = "exchopd" then
                            Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
                            end if
                            %><br />
                            Bruges til at beregne ServiceDesk aftalegruppe responstider.
                            Et incident indrapporteret onsdag morgen 04:45, med en aftalegruppe responstid p� 48 timer, bliver
                            s�ledes beregnet til at skulle v�re p�begyndt senest fredag kl. 9.00 n�r firmaet �bner.<br /><br />
                            
                            
                           

                     
                 <br /><br />
            <input type="image" src="../ill/opdaterpil.gif">
               <br /><br /></div></td>
                
            </tr>
            </table>
           
            <table cellspacing=0 cellpadding=0 border=0 width=80%>
            <tr>
                <td bgcolor="#ffffff" style="border:1px #5582d2 solid; padding:10px;">
                <a href="javascript:ShowHide('kexchange');"><b>Exchange server oplysninger:</b></a></td>
              </tr>
              <tr><td style="padding:10px;"><div id="kexchange" name="kexchange" style="visibility: hidden; display: none">
            
               
                  Bruges til integration med firmaets Exchange server.<br /><br />
                   <strong>
                 
                    Web-adresse til jeres Outlook Web Access:</strong>
                     <%
		                if func = "exchopd" then
		                Response.write "&nbsp;&nbsp;<font color=red><b>Opdateret!</b></font>"
		                end if
                    %>
                    <br>
                    <b>https://</b><input type="text" name="FM_exchange" id="FM_exchange" size="40" value="<%=ExchangeServerURL%>"><b>/exchange</b>
                    &nbsp;&nbsp;(Eks: mail.outzource.dk)<br>
                    Exchangeserver dom�ne:<br>
                    <input type="text" name="FM_exdom" id="FM_exdom" size="40" value="<%=ExchangeServerDOM%>">
                    &nbsp;&nbsp;(Eks: outz)<br>
                    Brugerkonto navn:<br>
                    <input type="text" name="FM_exbruger" id="FM_exbruger" size="40" value="<%=ExchangeServerBruger%>">
                    &nbsp;&nbsp;(Eks: timeout, uden dom�nenavn som er angivet ovenfor.)<br>
                    Brugerkonto password:<br>
                    <input type="password" name="FM_expw" id="FM_expw" size="40" value="<%=ExchangeServerPW%>">
                    &nbsp;&nbsp;(Eks: Q12Wert3)<br>
                       <br>
    <br>
    <input type="image" src="../ill/opdaterpil.gif">
              </div></td>  
            </tr>
          </table>
          
      
       <br /><br />
            <input type="image" src="../ill/opdaterpil.gif">          

    </form>
    
    *) Se aktuelle priser p� TimeOut moduler her:<br />
     <a href="http://www.outzource.dk/timeout_ver.asp" target="_blank" class=vmenu>http://www.outzource.dk/timeout_ver.asp..</a><br />
    <br>
    <br>
    <br>
    <a href="Javascript:history.back()">
        <img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
    <br>
    <br>
</div>
<%
	end select
	end if%><!--#include file="../inc/regular/footer_inc.asp"-->