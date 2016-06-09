	
	
	//function showsubm_on() {
	//document.getElementById("subm_on").style.visibility = "visible"
	//document.getElementById("subm_on").style.display = ""
	//document.getElementById("subm_on").focus();
	
	//document.getElementById("oktxt").style.visibility = "hidden"
	//document.getElementById("oktxt").style.display = "none"
	//}
	
	
	
	function mOvr(divId,src,clrOver) {
	src.bgColor = clrOver;	
	}
	
	function mOut(src,clrIn) { if (!src.contains(event.toElement)) { src.style.cursor = 'default'; src.bgColor = clrIn;}}
	
	
	
	
	function decimaler(pm, n, x){
	document.getElementById("FM_mbelob_"+n+"_"+x+"").value = document.getElementById("FM_mbelob_"+n+"_"+x+"").value.replace(",",".")
	
		if (pm == 'plus'){
		document.getElementById("FM_mbelob_"+n+"_"+x+"").value = parseFloat(document.getElementById("FM_mbelob_"+n+"_"+x+"").value) + parseFloat("0.01")
		}else{
		document.getElementById("FM_mbelob_"+n+"_"+x+"").value = parseFloat(document.getElementById("FM_mbelob_"+n+"_"+x+"").value) - parseFloat("0.01")
		}
	
		document.getElementById("FM_mbelob_"+n+"_"+x+"").value = document.getElementById("FM_mbelob_"+n+"_"+x+"").value.replace(".",",")
	
			
							//afrunder decimaler
							offSet_this = String(document.getElementById("FM_mbelob_"+n+"_"+x+"").value);
							offSetL_this = String(document.getElementById("FM_mbelob_"+n+"_"+x+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_mbelob_"+n+"_"+x+"").value = offSet_this.substring(0, pkt_this + 3)
							}
			
	
		document.getElementById("medarbbelobdiv_"+n+"_"+x+"").innerHTML = "<b>"+ document.getElementById("FM_mbelob_"+n+"_"+x+"").value +"</b>"
	
	
				antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value)
				highest_aval = parseInt((document.getElementById("highest_aval_"+x+"").value/1) + 1)
				antal_n_x2 = parseInt(document.getElementById("antal_n_"+x+"").value/1 - 1)
				
				//Finder medarbejder timepriser
				this_mtp = document.getElementById("FM_mtimepris_" + n + "_" + x +"").value.replace(",",".")
				
					var tmpbel = 0;
					var nyBeloebtot = 0;
					//Samler det totale antal timer på den enhedspris der er valgt
					for (i=1;i<antal_n_x;i++){
						if (i != 0) {
							if (parseFloat(this_mtp) == parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",","."))) {
							tmpbel = document.getElementById("FM_mbelob_" + i + "_" + x +"").value.replace(",",".")
							nyBeloebtot = nyBeloebtot + parseFloat(tmpbel)
							}
						}
					}
					
					//Opdaterer den valgte sum-aktivitet (Den der passer til timeprisen)
					for (i=-(antal_n_x2);i<highest_aval;i++){
							
							if (i != 0) {
								if (parseFloat(this_mtp) == parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+i+"").value.replace(",","."))) {
								//Totaler på den valgte sum-aktivitet
								document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = nyBeloebtot 
								document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = document.getElementById("FM_beloebthis_"+x+"_"+i+"").value.replace(".",",")
								
								document.getElementById("belobdiv_"+x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_"+i+"").value+"</b>"	
								}
							}	
					}
	
	setBeloebTot2(x)
	}
	
	
	
	
	
	
	function beregntimepris() {
	belob = document.getElementById("beregn_belob").value.replace(",",".")
	timer = document.getElementById("beregn_timer").value.replace(",",".")
	res = parseFloat(belob/timer)
	
	document.getElementById("beregn_tp").value = res
	document.getElementById("beregn_tp").value = document.getElementById("beregn_tp").value.replace(".",",")
	}
	
	
	
	//nul sum aktiviteten
	function setBeloebThis2(x,a){
	//alert("her")
	var belobThis = 0;
	timerThis = document.getElementById("FM_timerthis_"+x+"_"+a+"").value.replace(",",".")
	timeprisThis = document.getElementById("FM_enhedspris_"+x+"_"+a+"").value.replace(",",".")
	belobThis = parseFloat(timerThis * timeprisThis)
	document.getElementById("FM_beloebthis_"+x+"_"+a+"").value = belobThis
	
						
						
						//afrunder decimaler
							document.getElementById("FM_hidden_timepristhis_"+x+"_"+a+"").value = document.getElementById("FM_hidden_timepristhis_"+x+"_"+a+"").value.replace(".",",")
							offSet_this = String(document.getElementById("FM_hidden_timepristhis_"+x+"_"+a+"").value);
							offSetL_this = String(document.getElementById("FM_hidden_timepristhis_"+x+"_"+a+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_hidden_timepristhis_"+x+"_"+a+"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
							//afrunder decimaler
							document.getElementById("FM_timerthis_"+x+"_"+a+"").value = document.getElementById("FM_timerthis_"+x+"_"+a+"").value.replace(".",",")
							offSet_this = String(document.getElementById("FM_timerthis_"+x+"_"+a+"").value);
							offSetL_this = String(document.getElementById("FM_timerthis_"+x+"_"+a+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_timerthis_"+x+"_"+a+"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
							//afrunder decimaler
							document.getElementById("FM_beloebthis_"+x+"_"+a+"").value = document.getElementById("FM_beloebthis_"+x+"_"+a+"").value.replace(".",",")
							offSet_this = String(document.getElementById("FM_beloebthis_"+x+"_"+a+"").value);
							offSetL_this = String(document.getElementById("FM_beloebthis_"+x+"_"+a+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_beloebthis_"+x+"_"+a+"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
							//afrunder decimaler
							offSet_this = String(document.getElementById("FM_enhedspris_"+x+"_"+a+"").value);
							offSetL_this = String(document.getElementById("FM_enhedspris_"+x+"_"+a+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_enhedspris_"+x+"_"+a+"").value = offSet_this.substring(0, pkt_this + 3)
							}
					
					
	setTimerTot2(x)
	setBeloebTot2(x)
	}
	
	
	function setTimerTot2(x) {
	var antalsubakttimer = 0;
	var antalakttimertot = 0;
	if (window.event.keyCode == 37){ 
	} else {
		if (window.event.keyCode == 39){
		} else {
				as = document.getElementById("antal_subtotal_akt_"+x+"").value
				antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value/1 - 1)
				
				for (i=-(antal_n_x);i<=as;i++) 	{
					//alert(parseInt(i))
					if (document.getElementById("FM_show_akt_"+x+"_"+i+"").checked == true) {
					//alert(document.getElementById("FM_timerthis_"+x+"_"+parseInt(i)+"").value)	
						tmptim = document.getElementById("FM_timerthis_"+x+"_"+i+"").value.replace(",",".") 
						antalsubakttimer = antalsubakttimer + parseFloat(tmptim)
							
						}
				}
					
				
				
				//Total timer på faktura
				//Sub-total på akt.
				document.getElementById("timer_subtotal_akt_"+x+"").value = antalsubakttimer
				document.getElementById("timer_subtotal_akt_"+x+"").value = document.getElementById("timer_subtotal_akt_"+x+"").value.replace(".",",")
				offSet_this = String(document.getElementById("timer_subtotal_akt_"+x+"").value);
					offSetL_this = String(document.getElementById("timer_subtotal_akt_"+x+"").length);
					pkt_this = offSet_this.indexOf(",");
					if (pkt_this > 0){
						document.getElementById("timer_subtotal_akt_"+x+"").value = offSet_this.substring(0, pkt_this + 3)
					}
				
					
				antalx = document.getElementById("antal_x").value
				for (i=0;i<antalx;i++) 	{
					tmptim2 = document.getElementById("timer_subtotal_akt_"+i+"").value.replace(",",".") 
					antalakttimertot = antalakttimertot + parseFloat(tmptim2)
				}
				
				document.getElementById("FM_Timer").value = antalakttimertot
				document.getElementById("FM_Timer").value = document.getElementById("FM_Timer").value.replace(".",",") 
				
					offSet_this = String(document.getElementById("FM_Timer").value);
					offSetL_this = String(document.getElementById("FM_Timer").length);
					pkt_this = offSet_this.indexOf(",");
					if (pkt_this > 0){
						document.getElementById("FM_Timer").value = offSet_this.substring(0, pkt_this + 3)
					}
					document.getElementById("divtimertot").innerHTML = "<b>"+document.getElementById("FM_Timer").value+"</b>"
				}
	  		}
	}
	
	function setBeloebTot2(x) {
	var belobsubakt = 0;
	var belobakttot = 0;
	if (window.event.keyCode == 37){ 
	} else {
		if (window.event.keyCode == 39){
		} else {
				as = document.getElementById("antal_subtotal_akt_"+x+"").value
				antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value/1 - 1)
				
				
				for (i=-(antal_n_x);i<=as;i++) 	{
					if (document.getElementById("FM_show_akt_"+x+"_"+i+"").checked == true) {
					
						tmptim = document.getElementById("FM_beloebthis_"+x+"_"+i+"").value.replace(",",".") 
						belobsubakt = belobsubakt + parseFloat(tmptim)
						
					}
				}
				
				//Total beløb på faktura
				//Sub-total på akt.
				document.getElementById("belob_subtotal_akt_"+x+"").value = belobsubakt
				document.getElementById("belob_subtotal_akt_"+x+"").value = document.getElementById("belob_subtotal_akt_"+x+"").value.replace(".",",")
				offSet_this = String(document.getElementById("belob_subtotal_akt_"+x+"").value);
					offSetL_this = String(document.getElementById("belob_subtotal_akt_"+x+"").length);
					pkt_this = offSet_this.indexOf(",");
					if (pkt_this > 0){
						document.getElementById("belob_subtotal_akt_"+x+"").value = offSet_this.substring(0, pkt_this + 3)
					}
				
				antalx = document.getElementById("antal_x").value
				for (i=0;i<antalx;i++) 	{
					tmptim2 = document.getElementById("belob_subtotal_akt_"+i+"").value.replace(",",".") 
					belobakttot = belobakttot + parseFloat(tmptim2)
				}
				document.getElementById("FM_beloeb").value = belobakttot
				document.getElementById("FM_beloeb").value = document.getElementById("FM_beloeb").value.replace(".",",") 
				
					offSet_this = String(document.getElementById("FM_beloeb").value);
					offSetL_this = String(document.getElementById("FM_beloeb").length);
					pkt_this = offSet_this.indexOf(",");
					if (pkt_this > 0){
						document.getElementById("FM_beloeb").value = offSet_this.substring(0, pkt_this + 3)
					}
				document.getElementById("divbelobtot").innerHTML = "<b>"+document.getElementById("FM_beloeb").value+"</b>"
				}
	  		}
	}
		
		
	//Bruges af nul-sumaktiviteten	
	function enhedspris(x,a){ 
	
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				
				//Timepris på 0-sum ak.
				document.getElementById("FM_hidden_timepristhis_"+x+"_"+a+"").value = document.getElementById("FM_enhedspris_"+x+"_"+a+"").value.replace(",",".");
				document.getElementById("FM_enhedspris_"+x+"_"+a+"").value = document.getElementById("FM_enhedspris_"+x+"_"+a+"").value.replace(".",",")
				
				//Tjekker om timepris allerede eksisterer
				antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value)
				highest_aval = parseInt((document.getElementById("highest_aval_"+x+"").value/1) + 1)
				antal_n_x2 = parseInt(document.getElementById("antal_n_"+x+"").value/1 - 1)
				
				for (i=-(antal_n_x2);i<highest_aval;i++){
							if (i != 0) {
							//alert(parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+a+"").value)+" == "+ parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+i+"").value.replace(",",".")))
							if (parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+a+"").value) == parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+i+"").value.replace(",","."))) {
							
							alert ("Denne timepris er allerede ibrug på en sum-aktivitet. 'Andet?' sum-aktviteten nulstilles.")
							document.getElementById("FM_enhedspris_" + x + "_"+a+"").value = -2
							document.getElementById("FM_timerthis_" + x + "_"+a+"").value = 0
							document.getElementById("FM_beloebthis_" + x + "_"+a+"").value = 0
							document.getElementById("FM_show_akt_"+x+"_"+a+"").checked = false
							
							
							}
							}	
					}
				
							
							
							
				
							//afrunder decimaler
							document.getElementById("FM_hidden_timepristhis_"+x+"_"+a+"").value = document.getElementById("FM_hidden_timepristhis_"+x+"_"+a+"").value.replace(".",",")
							offSet_this = String(document.getElementById("FM_hidden_timepristhis_"+x+"_"+a+"").value);
							offSetL_this = String(document.getElementById("FM_hidden_timepristhis_"+x+"_"+a+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_hidden_timepristhis_"+x+"_"+a+"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
							//afrunder decimaler
							offSet_this = String(document.getElementById("FM_enhedspris_"+x+"_"+a+"").value);
							offSetL_this = String(document.getElementById("FM_enhedspris_"+x+"_"+a+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_enhedspris_"+x+"_"+a+"").value = offSet_this.substring(0, pkt_this + 3)
							}
			}
		}		
	}
	
	
	
	
	
	
	function nulstilfaktimer(x,a){
	if (a != 0){
		
			if (document.getElementById("FM_show_akt_"+x+"_"+a+"").checked == false) {
				
				if (document.getElementById("FM_showalert").value == 0){
				var t = confirm("Når en sum-aktivitet ikke faktureres, bliver alle fakturerbare medarbejder-timer der hører til denne sum-aktivitet automatisk overført til 'vente' timer.\nDenne meddelelse bliver kun vist en gang pr. faktura oprettelse.")
				document.getElementById("FM_showalert").value = 1
				
				//Hvis alert vises.	
					if (t) {
					document.getElementById("FM_timerthis_" +x+"_"+a+"").value = 0,00
					document.getElementById("FM_beloebthis_" +x+"_"+a+"").value = 0,00
					
					document.getElementById("timeprisdiv_" +x+"_"+a+"").innerHTML = "<b>0,00</b>"
					document.getElementById("belobdiv_"+x+"_"+a+"").innerHTML = "<b>0,00</b>"	
										
					
					this_akttp = document.getElementById("FM_enhedspris_"+x+"_"+a+"").value.replace(",",".")
					antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value)
								
								for (i=1;i<antal_n_x;i++){
									//alert(parseFloat(this_akttp) +" == "+ parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",",".")))
									if (i != 0){
										if (parseFloat(this_akttp) == parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",","."))) {
										document.getElementById("FM_m_vent_" +i+ "_"+x+"").value = document.getElementById("FM_m_fak_" +i+"_"+x+"").value
										document.getElementById("FM_mbelob_" +i+ "_"+x+"").value = 0,00
										document.getElementById("medarbbelobdiv_"+i+"_"+x+"").innerHTML = "<b>0,00</b>"	
										document.getElementById("FM_m_fak_" +i+"_"+x+"").value = 0,00
										document.getElementById("FM_show_medspec_"+i+"_"+x+"").checked = false
										}
									}
								}
					
					} else {
					document.getElementById("FM_show_akt_"+x+"_"+a+"").checked = true
					}
				
				
				//Hvis alert er vist 1 gang!
				} else {
					
				 	document.getElementById("FM_timerthis_" +x+"_"+a+"").value = 0,00
					document.getElementById("FM_beloebthis_" +x+"_"+a+"").value = 0,00
					
					document.getElementById("timeprisdiv_" +x+"_"+a+"").innerHTML = "<b>0,00</b>"
					document.getElementById("belobdiv_"+x+"_"+a+"").innerHTML = "<b>0,00</b>"	
										
					this_akttp = document.getElementById("FM_enhedspris_"+x+"_"+a+"").value.replace(",",".")
					antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value)
								
								for (i=1;i<antal_n_x;i++){
									//alert(parseFloat(this_akttp) +" == "+ parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",",".")))
										if (i != 0){
										if (parseFloat(this_akttp) == parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",","."))) {
										document.getElementById("FM_m_vent_" +i+ "_"+x+"").value = document.getElementById("FM_m_fak_" +i+"_"+x+"").value
										document.getElementById("FM_mbelob_" +i+ "_"+x+"").value = 0,00
										document.getElementById("medarbbelobdiv_"+i+"_"+x+"").innerHTML = "<b>0,00</b>"	
										document.getElementById("FM_m_fak_" +i+"_"+x+"").value = 0,00
										document.getElementById("FM_show_medspec_"+i+"_"+x+"").checked = false
										}
										}
								}
					
					
				}
			}
	}
	}
	
	
	
	function enhedsprismedarb(x, n){
	
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				
				//Finder a og n værdier
				antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value)
				highest_aval = parseInt((document.getElementById("highest_aval_"+x+"").value/1) + 1)
				antal_n_x2 = parseInt(document.getElementById("antal_n_"+x+"").value/1 - 1)
				
				//Finder medarbejder timepriser
				this_mtp = document.getElementById("FM_mtimepris_" + n + "_" + x +"").value.replace(",",".")
				this_mtp_opr = document.getElementById("FM_mtimepris_opr_" + n + "_" + x +"").value
				this_mtimer = document.getElementById("FM_m_fak_" + n + "_" + x +"").value.replace(",",".")
				
				//Sætter den nye timepris på medarbejderen efter replace
				document.getElementById("FM_mtimepris_" + n + "_" + x +"").value = this_mtp.replace(".",",")
				
				//Sætter medarbejder beløb på den valgte medarbejer
				document.getElementById("FM_mbelob_" + n + "_" + x +"").value = (this_mtp * this_mtimer)
				//parseFloat(this_mtp) * parseFloat(this_mtimer)
				document.getElementById("FM_mbelob_" + n + "_" + x +"").value = document.getElementById("FM_mbelob_" + n + "_" + x +"").value.replace(".",",")
				
				offSet_this = String(document.getElementById("FM_mbelob_" + n + "_" + x +"").value);
				offSetL_this = String(document.getElementById("FM_mbelob_" + n + "_" + x +"").length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.getElementById("FM_mbelob_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
				}
				document.getElementById("medarbbelobdiv_" + n + "_"+ x +"").innerHTML = "<b>"+document.getElementById("FM_mbelob_" + n + "_" + x +"").value+"</b>"	
					
				
				var nyBeloebtot = 0;
				var tmpbel = 0;
				var antalmedarbtimertot = 0;
				var tmptim = 0;
				
				//Afrunder og Tjekker for kommatal i timepris
				this_mtp_compare = this_mtp
				this_mtp_opr_compare = String(document.getElementById("FM_mtimepris_opr_" + n + "_" + x +"").value.replace(",","."));
				
				
					//Samler det totale antal timer på den enhedspris der er valgt
					for (i=1;i<antal_n_x;i++){
						
						//alert(parseFloat(this_mtp_compare) +" == "+ parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",",".")))
						if (i != 0) {
						if (parseFloat(this_mtp_compare) == parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",","."))) {
						tmpbel = document.getElementById("FM_mbelob_" + i + "_" + x +"").value.replace(",",".")
						nyBeloebtot = nyBeloebtot + parseFloat(tmpbel)
						tmptim = document.getElementById("FM_m_fak_" + i + "_" + x +"").value.replace(",",".") 
						antalmedarbtimertot = antalmedarbtimertot + parseFloat(tmptim)
						
						}
						}
					}
					
					
					
					ifHighestA = -1000
					foundone = 0
					//Opdaterer den valgte sum-aktivitet (Den der passer til timeprisen)
					//highest_aval
					
					for (i=-(antal_n_x2);i<highest_aval;i++){
							//alert(parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+i+"").value.replace(",",".")))
							//alert(parseFloat(this_mtp_compare) +" == "+ parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+i+"").value.replace(",",".")))
							//alert(i +" : "+ parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+i+"").value.replace(",",".")))
							//alert(i)
							if (i != 0) {
							
							if (parseFloat(this_mtp_compare) == parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+i+"").value.replace(",","."))) {
							
								if (foundone == 0){
								foundone = 1
								// ifHighestA bruges hvis der er flere medarbejdere med samme 
								// timepris og sidste medarbejder's a værdi er højere end den
								// sidste sum-aktivitets.
								ifHighestA = i
								}
							
							
							
							//Totaler på den valgte sum-aktivitet
							document.getElementById("FM_timerthis_"+x+"_"+i+"").value = antalmedarbtimertot
							document.getElementById("FM_timerthis_"+x+"_"+i+"").value = document.getElementById("FM_timerthis_"+x+"_"+i+"").value.replace(".",",") //document.getElementById("FM_m_fak_"+n+"_"+x+"").value
							document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = nyBeloebtot 
							document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = document.getElementById("FM_beloebthis_"+x+"_"+i+"").value.replace(".",",")
							
								
								//if (i != 0) {
								document.getElementById("timeprisdiv_" +x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_timerthis_"+x+"_"+i+"").value+"</b>"
								document.getElementById("belobdiv_"+x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_"+i+"").value+"</b>"	
								//}
							document.getElementById("FM_show_akt_"+x+"_"+i+"").checked = true
							}
							}	
					}
							
							
					
					var nyBeloebtot2 = 0;
					var tmpbel2 = 0;
					var antalmedarbtimertot2 = 0;
					var tmptim2 = 0;
					var showbesk = "";
					var showbesk1 = "";
					
					
					//Genberegner summen af timer på den opr. medarbejdertimepris
					//alert(parseFloat(this_mtp_compare) +"!="+ parseFloat(this_mtp_opr_compare))
					if (parseFloat(this_mtp_compare) != parseFloat(this_mtp_opr_compare)) {
						
						
						
						for (i=1;i<antal_n_x;i++){
							//alert(parseFloat(this_mtp_opr_compare) +"=="+ parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",",".")))
							if (i != 0) {
							if (parseFloat(this_mtp_opr_compare) == parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",","."))) {
							tmpbel2 = document.getElementById("FM_mbelob_" + i + "_" + x +"").value.replace(",",".")
							nyBeloebtot2 = nyBeloebtot2 + parseFloat(tmpbel2)
							
							tmptim2 = document.getElementById("FM_m_fak_" + i + "_" + x +"").value.replace(",",".") 
							antalmedarbtimertot2 = antalmedarbtimertot2 + parseFloat(tmptim2)
							}
							}
						}
						
						
						//Opdaterer sum-aktiviteten på den opr værdi
						for (i=-(antal_n_x2);i<highest_aval;i++){
							if (i != 0) {
							if (parseFloat(this_mtp_opr_compare) == parseFloat(document.getElementById("FM_enhedspris_"+x+"_"+i+"").value.replace(",","."))) {
							//Totaler på forladte sum-aktivitet
							document.getElementById("FM_timerthis_"+x+"_"+i+"").value = antalmedarbtimertot2 //document.getElementById("FM_m_fak_"+n+"_"+x+"").value
							document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = nyBeloebtot2
							
							document.getElementById("FM_timerthis_"+x+"_"+i+"").value = document.getElementById("FM_timerthis_"+x+"_"+i+"").value.replace(".",",")
							offSet_this = String(document.getElementById("FM_timerthis_"+x+"_"+i+"").value);
							offSetL_this = String(document.getElementById("FM_timerthis_"+x+"_"+i+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_timerthis_"+x+"_"+i+"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
							document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = document.getElementById("FM_beloebthis_"+x+"_"+i+"").value.replace(".",",")
							offSet_this = String(document.getElementById("FM_beloebthis_"+x+"_"+i+"").value);
							offSetL_this = String(document.getElementById("FM_beloebthis_"+x+"_"+i+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_beloebthis_"+x+"_"+i+"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
							document.getElementById("timeprisdiv_" +x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_timerthis_"+x+"_"+i+"").value+"</b>"
							document.getElementById("belobdiv_"+x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_"+i+"").value+"</b>"
							
								if (antalmedarbtimertot2 == 0) {
								document.getElementById("FM_show_akt_" + x + "_" + i +"").checked = false
								}
							}
							}	
						}
					
					}
							
							//Viser sumaktivitet div. (bruges ved redigering af fak.)
							if (parseInt(foundone) == 1) {
								//alert("her")
								//er n større end det højest brugte A på denne sum-aktivitet?
								if (n > ifHighestA) {
								useThisN = ifHighestA
								} else {
								useThisN = n
								}
							
								document.getElementById("sumaktdiv_"+x+"_"+useThisN+"").style.display = "";
								document.getElementById("sumaktdiv_"+x+"_"+useThisN+"").style.visibility = "visible";
								//alert(document.getElementById("sumaktbesk_"+x+"").value)
								
								//Beskrivelse / tekst opdateres i felt.  
								//showbesk1 = document.getElementById("sumaktbesk_"+x+"").value
								//document.getElementById("FM_aktbesk_" +x+"_"+useThisN+"").value = showbesk1
								
							}
							
							
							//Hvis der ikke findes en åben sumaktivitet med den valgte timepris
							if (parseInt(foundone) == 0) {
							//Åbner ny div
							document.getElementById("sumaktdiv_"+x+"_"+-(n)+"").style.display = "";
							document.getElementById("sumaktdiv_"+x+"_"+-(n)+"").style.visibility = "visible";
							
							 
							//Sætter værdier i de nye div
							document.getElementById("FM_enhedspris_" + x + "_"+ -(n) +"").value = document.getElementById("FM_mtimepris_" + n + "_" + x +"").value.replace(".",",")
							document.getElementById("enhprisdiv_" + x + "_"+ -(n) +"").innerHTML = "<b>"+ document.getElementById("FM_enhedspris_" + x + "_"+ -(n) +"").value +"</b>"
							
							//alert for samme pris som 0 akt.
							if (parseFloat(document.getElementById("FM_enhedspris_" + x + "_"+ -(n) +"").value) == parseFloat(document.getElementById("FM_enhedspris_" + x + "_0").value)) {
							alert ("Du har angivet samme timepris som 'Andet?' sum-aktiviteten, som derfor nulstilles.\nDet kan være nødvendigt at klikke på 'Calc' knappen endnu engang for at opdatere timer og beløb.")
							document.getElementById("FM_enhedspris_" + x + "_0").value = -2
							document.getElementById("FM_timerthis_" + x + "_0").value = 0
							document.getElementById("FM_beloebthis_" + x + "_0").value = 0
							document.getElementById("FM_show_akt_"+x+"_0").checked = false
							}
							
							
							document.getElementById("FM_timerthis_" + x + "_"+ -(n) +"").value = this_mtimer
							document.getElementById("FM_timerthis_" + x + "_"+ -(n) +"").value = document.getElementById("FM_timerthis_" + x + "_"+ -(n) +"").value.replace(".",",")
							
							offSet_this = String(document.getElementById("FM_timerthis_"+x+"_"+-(n)+"").value);
							offSetL_this = String(document.getElementById("FM_timerthis_"+x+"_"+-(n)+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_timerthis_"+x+"_"+-(n)+"").value = offSet_this.substring(0, pkt_this + 3)
							}
							document.getElementById("timeprisdiv_" + x + "_"+ -(n) +"").innerHTML = "<b>"+ document.getElementById("FM_timerthis_"+x+"_"+-(n)+"").value +"</b>"
							
							
							
							document.getElementById("FM_beloebthis_" + x + "_"+ -(n) +"").value = (this_mtimer * this_mtp) 
							document.getElementById("FM_beloebthis_" + x + "_"+ -(n) +"").value = document.getElementById("FM_beloebthis_" + x + "_"+ -(n) +"").value.replace(".",",")
							
							offSet_this = String(document.getElementById("FM_beloebthis_"+x+"_"+-(n)+"").value);
							offSetL_this = String(document.getElementById("FM_beloebthis_"+x+"_"+-(n)+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_beloebthis_"+x+"_"+-(n)+"").value = offSet_this.substring(0, pkt_this + 3)
							}
							document.getElementById("belobdiv_" + x + "_"+ -(n) +"").innerHTML = "<b>"+ document.getElementById("FM_beloebthis_" + x + "_"+ -(n) +"").value +"</b>"
							
							document.getElementById("FM_show_akt_" + x + "_" + -(n) +"").checked = true
							
							//Tekst (Fjernet, da det er irriterende at den tekst man har skrevet bilver slettet med tryk på "calc")
							//showbesk = document.getElementById("sumaktbesk_"+x+"").value
							//document.getElementById("FM_aktbesk_" + x + "_"+ -(n) +"").value = showbesk
							} 
						
							//Sætter timepris opr
							document.getElementById("FM_mtimepris_opr_" + n + "_" + x +"").value = this_mtp.replace(".",",")
							
							//Afrunder
							offSet_this = String(document.getElementById("FM_mtimepris_opr_" + n + "_" + x +"").value);
							offSetL_this = String(document.getElementById("FM_mtimepris_opr_" + n + "_" + x +"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_mtimepris_opr_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
							
				setTimerTot2(x)
				setBeloebTot2(x)
				}
		}		
	}
	
	
	function offsetventer(x,n){
		if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				offSet_this = String(document.getElementById("FM_m_vent_" + n + "_" + x +"").value.replace(".",","));
				offSetL_this = String(document.getElementById("FM_m_vent_" + n + "_" + x +"").length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.getElementById("FM_m_vent_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
				}
			}
		}
	}
	
	
	function offsetfak(x,n){
		if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
					offSet_this = String(document.getElementById("FM_m_fak_" + n + "_" + x +"").value.replace(".",","));
					offSetL_this = String(document.getElementById("FM_m_fak_" + n + "_" + x +"").length);
					pkt_this = offSet_this.indexOf(",");
					if (pkt_this > 0){
						document.getElementById("FM_m_fak_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
					}
				}
		}
	}			
	
	function offsetmtp(x,n){
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				offSet_this = String(document.getElementById("FM_mtimepris_" + n + "_" + x +"").value.replace(".",","));
				offSetL_this = String(document.getElementById("FM_mtimepris_" + n + "_" + x +"").length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.getElementById("FM_mtimepris_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
				}
			}
		}
	}
	
	
	//Bruges ikke mere
	function updantaltimerprakt(x, n, oprnew){
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				var fak_opr = 0;
				var fak = 0;
				var mtimepris = 0;
				var vent = 0;
				
				a = document.getElementById("FM_m_aval_"+n+ "_"+x+"").value
				//fak_opr = document.getElementById("FM_m_fak_opr_" + n + "_" + x +"").value;
				
				fak = document.getElementById("FM_m_fak_" + n + "_" + x +"").value.replace(",",".") / 1; 
				
				this_mtp = document.getElementById("FM_mtimepris_" + n + "_" + x +"").value.replace(",",".")
				mtimepris = document.getElementById("FM_mtimepris_" + n + "_" + x +"").value.replace(",",".")
				document.getElementById("FM_mbelob_" + n + "_" + x +"").value = parseFloat(mtimepris) * parseFloat(fak)
				
				var nyBeloebtot = 0;
				var tmpbel = 0;
				var antalmedarbtimertot = 0;
				var tmptim = 0;
				
				//Afrunder og tjekker for kommatal
				this_mtp_compare = this_mtp
				antal_n_x = document.getElementById("antal_n_"+x+"").value
				
					for (i=1;i<antal_n_x;i++){
						if (parseFloat(this_mtp_compare) == parseFloat(document.getElementById("FM_mtimepris_" + i + "_" + x +"").value.replace(",","."))) {
						tmpbel = document.getElementById("FM_mbelob_" + i + "_" + x +"").value.replace(",",".")
						nyBeloebtot = nyBeloebtot + parseFloat(tmpbel)
						
						tmptim = document.getElementById("FM_m_fak_" + i + "_" + x +"").value.replace(",",".") 
						antalmedarbtimertot = antalmedarbtimertot + parseFloat(tmptim)
						}
					}
					
					//Totaler på aktiviteten
					document.getElementById("FM_timerthis_"+x+"_"+a+"").value = antalmedarbtimertot //document.getElementById("FM_m_fak_"+n+"_"+x+"").value
					document.getElementById("FM_beloebthis_"+x+"_"+a+"").value = nyBeloebtot 
					
					//replace . med ,
					document.getElementById("FM_mbelob_" + n + "_" + x +"").value = document.getElementById("FM_mbelob_" + n + "_" + x +"").value.replace(".",",")
					document.getElementById("FM_m_fak_" + n + "_" + x +"").value = document.getElementById("FM_m_fak_" + n + "_" + x +"").value.replace(".",",")
					
					
					//afrunder decimaler
							offSet_this = String(document.getElementById("FM_mbelob_" + n + "_" + x +"").value);
							offSetL_this = String(document.getElementById("FM_mbelob_" + n + "_" + x +"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_mbelob_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
							offSet_this = String(document.getElementById("FM_m_fak_" + n + "_" + x +"").value);
							offSetL_this = String(document.getElementById("FM_m_fak_" + n + "_" + x +"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_m_fak_" + n + "_" + x +"").value = offSet_this.substring(0, pkt_this + 3)
							}
							
					
					//document.getElementById("timeprisdiv_" +x+"_"+a+"").innerHTML = "<b>"+ document.getElementById("FM_timerthis_"+x+"_"+a+"").value.replace(".",",") +"</b>"
					document.getElementById("medarbbelobdiv_" + n + "_"+ x +"").innerHTML = "<b>"+document.getElementById("FM_mbelob_" + n + "_" + x +"").value+"</b>"	
					//document.getElementById("belobdiv_" + x + "_"+ a +"").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_"+a+"").value+"</b>"
							
					enhedsprismedarb(x, n)
							
			}
		}
	}
	
	
	function isNum(passedVal){
	invalidChars = " /:;<>abcdefghijklmnopqrstuvwxyzæøå"
	if (passedVal == ""){
	return false
	}
	
	for (i = 0; i < invalidChars.length; i++) {
	badChar = invalidChars.charAt(i)
		if (passedVal.indexOf(badChar, 0) != -1){
		return false
		}
	}
	
		for (i = 0; i < passedVal.length; i++) {
			if (passedVal.charAt(i) == ".") {
			return true
			}
			
			if (passedVal.charAt(i) == ",") {
			return true
			}
			
			else {
				if (passedVal.charAt(i) < "0") {
				return false
				}
					if (passedVal.charAt(i) > "9") {
					return false
					}
				}
			return true
			} 
			
	}
	
	function validZip(inZip){
	if (inZip == "") {
	return true
	}
	if (isNum(inZip)){
	return true
	}
	return false
	}
	
	function tjektimer(x,a){
	var oprTimerThisTj = 0;
	oprTimerThisTj = (document.getElementById("FM_hidden_timerthis_"+x+"_"+a+"").value.replace(",",".") / 1);
		if (!validZip(document.getElementById("FM_timerthis_"+x+"_"+a+"").value)){
		alert("Der er angivet et ugyldigt tegn.")
		document.getElementById("FM_timerthis_"+x+"_"+a+"").value = oprTimerThisTj;
		document.getElementById("FM_timerthis_"+x+"_"+a+"").focus()
		document.getElementById("FM_timerthis_"+x+"_"+a+"").select()
		return false
		}
	return true
	}
	
	function tjektimer2(x,n){
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				var oprTimerThisTj = 0;
				oprTimerThisTj = (document.getElementById("FM_m_fak_opr_"+n+"_"+x+"").value.replace(",",".") / 1);
					if (!validZip(document.getElementById("FM_m_fak_"+n+"_"+x+"").value)){
					alert("Der er angivet et ugyldigt tegn.")
					document.getElementById("FM_m_fak_"+n+"_"+x+"").value = oprTimerThisTj;
					document.getElementById("FM_m_fak_"+n+"_"+x+"").focus()
					document.getElementById("FM_m_fak_"+n+"_"+x+"").select()
					return false
					}
				return true
				}
			}
	}
	
	function tjektimer4(x,n){
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				var oprTimerThisTj = 0;
				oprTimerThisTj = (document.getElementById("FM_m_vent_opr_"+n+"_"+x+"").value.replace(",",".") / 1);
					if (!validZip(document.getElementById("FM_m_vent_"+n+"_"+x+"").value)){
					alert("Der er angivet et ugyldigt tegn.")
					document.getElementById("FM_m_vent_"+n+"_"+x+"").value = oprTimerThisTj;
					document.getElementById("FM_m_vent_"+n+"_"+x+"").focus()
					document.getElementById("FM_m_vent_"+n+"_"+x+"").select()
					return false
					}
				return true
				}
			}
	}
	
	function tjektimepris(x){
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
					var oprTimerThisTj = 0;
					oprTimerThisTj = (document.getElementById("FM_hidden_timepristhis_"+x+"_0").value.replace(",",".") / 1);
					if (!validZip(document.getElementById("FM_enhedspris_"+x+"_0").value)){
					alert("Der er angivet et ugyldigt tegn.")
					document.getElementById("FM_enhedspris_"+x+"_0").value = oprTimerThisTj;
					document.getElementById("FM_enhedspris_"+x+"_0").focus()
					document.getElementById("FM_enhedspris_"+x+"_0").select()
					return false
					}
				return true
				}
			}
	}
	
	function tjektimer3(x,n){
	var oprTimerThisTj = 0;
	oprTimerThisTj = (document.getElementById("FM_mtimepris_opr_"+n+"_"+x+"").value.replace(",",".") / 1);
		if (!validZip(document.getElementById("FM_mtimepris_"+n+"_"+x+"").value)){
		alert("Der er angivet et ugyldigt tegn.")
		document.getElementById("FM_mtimepris_"+n+"_"+x+"").value = oprTimerThisTj;
		document.getElementById("FM_mtimepris_"+n+"_"+x+"").focus()
		document.getElementById("FM_mtimepris_"+n+"_"+x+"").select()
		return false
		}
	return true
	}
	
	
	
	function tjekBeloeb(x,a){
		if (!validZip(document.getElementById("FM_beloebthis_"+x+"_"+a+"").value)){
		alert("Der er angivet et ugyldigt tegn.")
		setBeloebTot2(x)
		return false
		}
	return true
	}
	
	
	function showborder(x,n) {
	document.getElementById("beregn_"+n+"_"+x+"_a").style.borderColor = "red"
	document.getElementById("beregn_"+n+"_"+x+"_b").style.borderColor = "red"
	document.getElementById("subm_on").style.visibility = "hidden"
	document.getElementById("subm_on").style.display = "none"
	document.getElementById("subm_off").style.visibility = "visible"
	document.getElementById("subm_off").style.display = ""
	}
	
	function hideborder(x,n) {
	document.getElementById("beregn_"+n+"_"+x+"_a").style.borderColor = ""
	document.getElementById("beregn_"+n+"_"+x+"_b").style.borderColor = ""
	document.getElementById("subm_on").style.visibility = "visible"
	document.getElementById("subm_on").style.display = ""
	document.getElementById("subm_off").style.visibility = "hidden"
	document.getElementById("subm_off").style.display = "none"
	}
	
	function showborder2(x,a) {
	document.getElementById("beregn2_"+x+"_"+a+"_a").style.borderColor = "red"
	document.getElementById("beregn2_"+x+"_"+a+"_b").style.borderColor = "red"
	document.getElementById("subm_on").style.visibility = "hidden"
	document.getElementById("subm_on").style.display = "none"
	document.getElementById("subm_off").style.visibility = "visible"
	document.getElementById("subm_off").style.display = ""
	}
	
	function hideborder2(x,a) {
	document.getElementById("beregn2_"+x+"_"+a+"_a").style.borderColor = ""
	document.getElementById("beregn2_"+x+"_"+a+"_b").style.borderColor = ""
	document.getElementById("subm_on").style.visibility = "visible"
	document.getElementById("subm_on").style.display = ""
	document.getElementById("subm_off").style.visibility = "hidden"
	document.getElementById("subm_off").style.display = "none"
	}
	
	
	/// Opret fakturaer på en aftale
	function aft_stk_pris() {
	var bel = 0;
	bel = document.getElementById("FM_beloeb").value.replace(",",".");
	var tim = 0;
	tim = document.getElementById("FM_timer").value.replace(",",".");
	var tmpstkprisval = 0;
	tmpstkprisval = (bel/tim);
	document.getElementById("FM_stkpris").value = tmpstkprisval
	
	document.getElementById("FM_stkpris").value = document.getElementById("FM_stkpris").value.replace(".",",");
	
	offSet_this = String(document.getElementById("FM_stkpris").value);
	offSetL_this = String(document.getElementById("FM_stkpris").length);
	pkt_this = offSet_this.indexOf(",");
	if (pkt_this > 0){
		document.getElementById("FM_stkpris").value = offSet_this.substring(0, pkt_this + 3)
	}
	
	}