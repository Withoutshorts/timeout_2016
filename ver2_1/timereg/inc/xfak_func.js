	
	
	function overtilsum(x){
	////alert("her")
	var timer = 0;
	var beloeb = 0;
	var enhedspris = 0;
	timer = document.getElementById("timer_subtotal_akt_"+x).value.replace(",",".")
	beloeb = document.getElementById("belob_subtotal_akt_"+x).value.replace(",",".")
	
	if (timer != 0){
	enhedspris = parseFloat(beloeb/timer)
	} else {
	enhedspris = 0
	}
	//enhedspris = 300
	
	document.getElementById("FM_timerthis_"+x+"_0").value = timer
	document.getElementById("FM_timerthis_"+x+"_0").value = document.getElementById("FM_timerthis_"+x+"_0").value.replace(".",",")
	
    ////document.getElementById("FM_hidden_timerthis"+x+"_0").value = timer //document.getElementById("FM_timerthis_"+x+"_0").value
	
	//alert(""+document.getElementById("FM_hidden_timerthis"+x+"_0").value+"")
	
	document.getElementById("FM_beloebthis_"+x+"_0").value = beloeb
	document.getElementById("FM_beloebthis_"+x+"_0").value = document.getElementById("FM_beloebthis_"+x+"_0").value.replace(".",",")
	
	document.getElementById("FM_enhedspris_"+x+"_0").value = enhedspris
	document.getElementById("FM_enhedspris_"+x+"_0").value = document.getElementById("FM_enhedspris_"+x+"_0").value.replace(".",",")
	
				    offSet_this = String(document.getElementById("FM_enhedspris_"+x+"_0").value);
					offSetL_this = String(document.getElementById("FM_enhedspris_"+x+"_0").length);
					pkt_this = offSet_this.indexOf(",");
					if (pkt_this > 0){
						document.getElementById("FM_enhedspris_"+x+"_0").value = offSet_this.substring(0, pkt_this + 3)
					}
					
	/// Sætter Andre sum-aktivitet aktiv 				
	document.getElementById("FM_show_akt_"+x+"_0").checked = true
	
	    /// Sætter sum-aktiviteter inaktive 
	    antal_x = (document.getElementById("antal_x").value)
	    for (i=0;i<antal_x;i++){
	        
	        antal_highest_aval = parseInt(document.getElementById("highest_aval_"+i+"").value/1 + 1) 
            for (e=1;e<antal_highest_aval;e++){
            document.getElementById("FM_show_akt_"+i+"_"+e).checked = false
            }
            
	    }
	    
	    
	    // Sætter medarbejdere inaktive 
	    antal_n_x = parseInt(document.getElementById("antal_n_"+x+"").value)
		for (i=1;i<antal_n_x;i++){
				if (i != 0){
				document.getElementById("FM_show_medspec_"+i+"_"+x+"").checked = false
				}
	     }
	     
	     
	     document.getElementById("FM_rabat_"+x+"_0").value = "0"
	     document.getElementById("FM_rabat_"+x+"_0").style.backgroundColor = "#ffff99"
	     
	     document.getElementById("angivtxt_"+x+"_0").style.visibility = "visible"
	     document.getElementById("angivtxt_"+x+"_0").style.display = ""
	     
	     
	}
	
	
	function setbgcol(x){
	document.getElementById("FM_rabat_"+x+"_0").style.backgroundColor = "#ffffff"
	}
	
	
	
	
	function offsetSumTtimer(x){
    if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				offSet_this = String(document.getElementById("timer_subtotal_akt_"+x).value.replace(".",","));
				offSetL_this = String(document.getElementById("timer_subtotal_akt_"+x).length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.getElementById("timer_subtotal_akt_"+x).value = offSet_this.substring(0, pkt_this + 3)
				}
			}
		}
	}
	
	function offsetSumTbeloeb(x){
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
				offSet_this = String(document.getElementById("belob_subtotal_akt_"+x).value.replace(".",","));
				offSetL_this = String(document.getElementById("belob_subtotal_akt_"+x).length);
				pkt_this = offSet_this.indexOf(",");
				if (pkt_this > 0){
					document.getElementById("belob_subtotal_akt_"+x).value = offSet_this.substring(0, pkt_this + 3)
				}
			}
		}
	}
	
	
	function showmodtagerafsender() {
	document.getElementById("modtagerafsender").style.visibility = "visible"
	document.getElementById("modtagerafsender").style.display = ""
	}
	
	function hidemodtagerafsender() {
	document.getElementById("modtagerafsender").style.visibility = "hidden"
	document.getElementById("modtagerafsender").style.display = "none"
	}
	
	
	
	function opd_akt_endhed(nyenhed,val) {
	    var enh_val = 0;
	    enh_val = val
	    antal_x = (document.getElementById("antal_x").value)
	    for (i=0;i<antal_x;i++){
	    
	         // Medarbejder felter
	        antal_nx = (document.getElementById("antal_n_"+i+"").value) 
            for (e=1;e<antal_nx;e++){
                document.getElementById("FM_med_enh_"+e+"_"+i).value = enh_val
            } 
	        
	        // Aktiviteter
	        antal_highest_aval = parseInt(document.getElementById("highest_aval_"+i+"").value/1 + 1) 
            for (e=0;e<antal_highest_aval;e++){
            document.getElementById("FM_akt_enh_"+i+"_"+e).value = enh_val
            //nyenhed
            }
            
	    }
	}
	
	function opd_timer_faktor(xval) {
	    var fak = 0;
	    var newfak = 0;
	    var faktorval = document.getElementById("aktfaktor_"+xval).value 
	         
	         
	         
	         
	        // Medarbejder felter
	        antal_nx = (document.getElementById("antal_n_"+xval+"").value) 
            for (e=1;e<antal_nx;e++){
                fak = document.getElementById("FM_m_fak_"+e+"_"+xval).value 
                newfak = parseFloat(fak * faktorval)
                document.getElementById("FM_m_fak_"+e+"_"+xval).value = newfak 
                document.getElementById("FM_m_fak_"+e+"_"+xval).value = document.getElementById("FM_m_fak_"+e+"_"+xval).value.replace(",",".")
            
                enhedsprismedarb(xval,e)
                
            } 
            
            
	        
	        // Aktiviteter
	        // antal_highest_aval = parseInt(document.getElementById("highest_aval_"+i+"").value/1 + 1) 
            // for (e=0;e<antal_highest_aval;e++){
             
             // timer = document.getElementById("FM_timerthis_"+xval+"_"+e).value 
             // newtimer = parseFloat(timer * faktorval)
             // document.getElementById("FM_timerthis_"+xval+"_"+e).value = newtimer 
             // document.getElementById("FM_timerthis_"+xval+"_"+e).value = document.getElementById("FM_timerthis_"+xval+"_"+e).value.replace(",",".")
           
            // }
          
	}
	
	
	
	function setenhedpaakt(n,x,aktmed) {
	//alert("n:"+n+" x:"+x)
	    if (aktmed == 1){
	        tpris = document.getElementById("FM_mtimepris_"+n+"_"+x).value 
	        ny_enheds_ang = document.getElementById("FM_med_enh_"+n+"_"+x).value
	        } else {
	        tpris = document.getElementById("FM_enhedspris_"+x+"_"+n).value 
	        ny_enheds_ang = document.getElementById("FM_akt_enh_"+x+"_"+n).value
	        //alert(tpris)
	    }
	   
	    
	        
	        // Medarbejder felter
	        antal_nx = (document.getElementById("antal_n_"+x+"").value) 
            for (e=1;e<antal_nx;e++){
                if (document.getElementById("FM_mtimepris_"+e+"_"+x).value == tpris){ 
                document.getElementById("FM_med_enh_"+e+"_"+x).value = ny_enheds_ang
                }
            } 
	        
	        // Aktiviteter
	        antal_highest_aval = parseInt(document.getElementById("highest_aval_"+x+"").value/1 + 1) 
            for (e=0;e<antal_highest_aval;e++){
                if (document.getElementById("FM_enhedspris_"+x+"_"+e).value == tpris){
                document.getElementById("FM_akt_enh_"+x+"_"+e).value = ny_enheds_ang
                }
            }
            
	    
	}
	
	
	function opd_rabatrall() {
	    val = document.getElementById("FM_rabat_all").value
	    antal_x = (document.getElementById("antal_x").value)
	    
	    for (i=0;i<antal_x;i++){
	        
	        // Medarbejder felter
	        antal_nx = (document.getElementById("antal_n_"+i+"").value) 
            for (e=1;e<antal_nx;e++){
                document.getElementById("FM_mrabat_"+e+"_"+i).value = val
                //document.getElementById("FM_mrabat_"+m).options["1"].selected = true
            } 
	        
	        
	        // Sum aktivitets felter        
            as_this = parseInt(document.getElementById("antal_subtotal_akt_"+i+"").value/1 + 1)
            antal_n_x_this = parseInt(document.getElementById("antal_n_"+i+"").value/1 - 1)
             
            for (e=-(antal_n_x_this);e<as_this;e++){
                
                //alert("ok")
                document.getElementById("FM_rabat_"+i+"_"+e).value = val
                if (e != 0) {
                document.getElementById("rabatdiv_"+i+"_"+e).innerHTML = "<b>"+(val*100)+" %</b>"
                }
            
            }
            
	    }
	    
	    // Materialer 
	    antal_mat = parseInt(document.getElementById("FM_antal_materialer_ialt").value/1)
	    for (m=0;m<antal_mat;m++){
	    document.getElementById("FM_matrabat_"+m).value = val
	    }
	    
	}
	
	
	function setBeloebTotAll(){
        //val = document.getElementById("FM_rabat_all").value
	    antal_x = (document.getElementById("antal_x").value)
	    
	    for (x=0;x<antal_x;x++){
	        
	        setBeloebThis2(x,0)
	        
	        // Medarbejder felter
	        antal_nx = (document.getElementById("antal_n_"+x+"").value) 
            for (n=1;n<antal_nx;n++){
               enhedsprismedarb(x, n)
            } 
	        
	        
	    }
	    
	    
	    
	    // Materialer 
	    antal_mat = parseInt(document.getElementById("FM_antal_materialer_ialt").value/1)
	    //alert(antal_mat)
	        for (m=0;m<antal_mat;m++){
	        beregnmatpris(m,0)
	        }
	    
	        materialerTot()
	    
	    alert("Rabat opdateret!")
	    
	}
	

	
	
	function genberegntimeprissumakt(x){
	var timer = 0;
	var beloeb = 0;
	var enhedspris = 0;
	var kvotient = 0;
	
	beloeb = document.getElementById("FM_beloebthis_"+x+"_0").value.replace(",",".")
	timer = document.getElementById("FM_timerthis_"+x+"_0").value.replace(",",".")
	
	var rbt = 0;
	rbt = document.getElementById("FM_rabat_"+x+"_0").value 
	enhedspris = (beloeb/timer) / (1-rbt)
	
	
	//kvotient = (1/rbt)
	//nytbelob = (beloeb*parseFloat(kvotient))
	//alert(nytbelob)
	//enhedspris = parseFloat(nytbelob/timer)
	
	//alert(enhedspris)
	
	
				    offSet_this = String(enhedspris);
					offSetL_this = String(enhedspris.length);
					pkt_this = offSet_this.indexOf(".");
					if (pkt_this > 0){
					document.getElementById("FM_enhedspris_"+x+"_0").value = offSet_this.substring(0, pkt_this + 3)
					} else {
					document.getElementById("FM_enhedspris_"+x+"_0").value = enhedspris
					}
	
	document.getElementById("FM_enhedspris_"+x+"_0").value = document.getElementById("FM_enhedspris_"+x+"_0").value.replace(".",",")
	
	}
	
	
	
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
	
	
	
	//Beregning af Andre? Sum aktiviteten
	function setBeloebThis2(x,a){
	//alert("her")
	var belobThis = 0;
	timerThis = document.getElementById("FM_timerthis_"+x+"_"+a+"").value.replace(",",".")
	timeprisThis = document.getElementById("FM_enhedspris_"+x+"_"+a+"").value.replace(",",".")
	
	
	rabat = document.getElementById("FM_rabat_"+x+"_"+a+"").value
	belobThis = parseFloat(timerThis * timeprisThis - (timerThis * timeprisThis * rabat))
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
	var tmpmatbelob = 0;
	var belobmattot = 0;
	
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
				
				// MATERIALER
			    antalm = document.getElementById("FM_antal_materialer_ialt").value
				antalm = antalm 
				
				for (i=0;i<antalm;i++) 	{
					tmpmatbelob = document.getElementById("FM_matbeloeb_"+i+"").value.replace(",",".") 
					belobmattot = belobmattot + parseFloat(tmpmatbelob)
				}
				
			
				document.getElementById("FM_timer_beloeb").value = belobakttot 
				document.getElementById("FM_timer_beloeb").value = document.getElementById("FM_timer_beloeb").value.replace(".",",") 
				document.getElementById("FM_beloeb").value = belobakttot + belobmattot
				document.getElementById("FM_beloeb").value = document.getElementById("FM_beloeb").value.replace(".",",") 
				
					offSet_this = String(document.getElementById("FM_beloeb").value);
					offSetL_this = String(document.getElementById("FM_beloeb").length);
					pkt_this = offSet_this.indexOf(",");
					if (pkt_this > 0){
						document.getElementById("FM_beloeb").value = offSet_this.substring(0, pkt_this + 3)
					}
				document.getElementById("divbelobtot").innerHTML = "<b>"+document.getElementById("FM_beloeb").value+"</b>"
				document.getElementById("divtimerbelobtot").innerHTML = "<b>"+document.getElementById("FM_timer_beloeb").value+"</b>"
				
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
							document.getElementById("FM_rabat_" + x + "_"+a+"").value = 0
							
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
					
					//rbt
					document.getElementById("rabatdiv_"+x+"_"+a+"").innerHTML = "<b>0 %</b>"					
					
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
					//rbt
					document.getElementById("rabatdiv_"+x+"_"+a+"").innerHTML = "<b>0 %</b>"
							
					
										
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
				//alert("hej")
				rabat = document.getElementById("FM_mrabat_" + n + "_" + x +"").value
				document.getElementById("FM_mbelob_" + n + "_" + x +"").value = (this_mtp * this_mtimer) - (this_mtp * this_mtimer * rabat)
			    
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
							//rbt
							document.getElementById("FM_rabat_"+x+"_"+i+"").value = rabat
							
								
								//if (i != 0) {
								document.getElementById("timeprisdiv_" +x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_timerthis_"+x+"_"+i+"").value+"</b>"
								document.getElementById("belobdiv_"+x+"_"+i+"").innerHTML = "<b>"+document.getElementById("FM_beloebthis_"+x+"_"+i+"").value+"</b>"	
								//rbt
							    document.getElementById("rabatdiv_"+x+"_"+i+"").innerHTML = "<b>"+(100 * document.getElementById("FM_rabat_"+x+"_"+i+"").value)+" %</b>"
							
								//}
							document.getElementById("FM_show_akt_"+x+"_"+i+"").checked = true
							}
							}	
					}
							
							
					
					
					//Beregning af totoler
					
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
							//rbt
							document.getElementById("FM_rabat_"+x+"_"+i+"").value = rabat
							
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
							//rbt
							document.getElementById("rabatdiv_"+x+"_"+i+"").innerHTML = "<b>"+(100 * document.getElementById("FM_rabat_"+x+"_"+i+"").value)+" %</b>"
							
							
							
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
							
							
							rabat = document.getElementById("FM_mrabat_" + n + "_" + x +"").value
				            document.getElementById("FM_beloebthis_" + x + "_"+ -(n) +"").value = (this_mtp * this_mtimer) - (this_mtp * this_mtimer * rabat)
			                //rbt 
							document.getElementById("FM_rabat_"+x+"_"+-(n)+"").value = rabat
						
			                
							//document.getElementById("FM_beloebthis_" + x + "_"+ -(n) +"").value = (this_mtimer * this_mtp) 
							document.getElementById("FM_beloebthis_" + x + "_"+ -(n) +"").value = document.getElementById("FM_beloebthis_" + x + "_"+ -(n) +"").value.replace(".",",")
							
							offSet_this = String(document.getElementById("FM_beloebthis_"+x+"_"+-(n)+"").value);
							offSetL_this = String(document.getElementById("FM_beloebthis_"+x+"_"+-(n)+"").length);
							pkt_this = offSet_this.indexOf(",");
							if (pkt_this > 0){
								document.getElementById("FM_beloebthis_"+x+"_"+-(n)+"").value = offSet_this.substring(0, pkt_this + 3)
							    }
							document.getElementById("belobdiv_" + x + "_"+ -(n) +"").innerHTML = "<b>"+ document.getElementById("FM_beloebthis_" + x + "_"+ -(n) +"").value +"</b>"
							
							//rbt
							document.getElementById("rabatdiv_" + x + "_"+ -(n) +"").innerHTML = "<b>"+(100 * document.getElementById("FM_rabat_" + x + "_"+ -(n) +"").value) +" %</b>"
							
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
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
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
			}
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
	//document.getElementById("beregn_"+n+"_"+x+"_b").style.borderColor = "red"
	
	//document.getElementById("subm_on").style.visibility = "hidden"
	//document.getElementById("subm_on").style.display = "none"
	//document.getElementById("subm_off").style.visibility = "visible"
	//document.getElementById("subm_off").style.display = ""
	}
	
	function hideborder(x,n) {
	
	document.getElementById("beregn_"+n+"_"+x+"_a").style.borderColor = ""
	//document.getElementById("beregn_"+n+"_"+x+"_b").style.borderColor = ""
	
	//document.getElementById("subm_on").style.visibility = "visible"
	//document.getElementById("subm_on").style.display = ""
	//document.getElementById("subm_off").style.visibility = "hidden"
	//document.getElementById("subm_off").style.display = "none"
	}
	
	function showborder2(x,a) {
	document.getElementById("beregn2_"+x+"_"+a+"_a").style.borderColor = "red"
	
	//document.getElementById("beregn2_"+x+"_"+a+"_b").style.borderColor = "red"
	
	//document.getElementById("subm_on").style.visibility = "hidden"
	//document.getElementById("subm_on").style.display = "none"
	//document.getElementById("subm_off").style.visibility = "visible"
	//document.getElementById("subm_off").style.display = ""
	}
	
	function hideborder2(x,a) {
	document.getElementById("beregn2_"+x+"_"+a+"_a").style.borderColor = ""
	
	//document.getElementById("beregn2_"+x+"_"+a+"_b").style.borderColor = ""
	
	//document.getElementById("subm_on").style.visibility = "visible"
	//document.getElementById("subm_on").style.display = ""
	//document.getElementById("subm_off").style.visibility = "hidden"
	//document.getElementById("subm_off").style.display = "none"
	}
	
	
	/// Opret fakturaer på en aftale
	function aft_stk_pris() {
	//var bel = 0;
	//bel = document.getElementById("FM_beloeb").value.replace(",",".");
	var tim = 0;
	tim = document.getElementById("FM_timer").value.replace(",",".");
	var rabat = 0;
	rbt = document.getElementById("FM_rabat").value
	var stkpris = 0;
	stkpris = document.getElementById("FM_stkpris").value.replace(",",".")
	
	var belob = 0;
	belob = parseFloat(tim * stkpris - (tim * stkpris * rbt))
	
	document.getElementById("FM_beloeb").value = belob
	document.getElementById("FM_beloeb").value = document.getElementById("FM_beloeb").value.replace(".",",")
	
	//document.getElementById("FM_stkpris").value = tmpstkprisval
    //document.getElementById("FM_stkpris").value = document.getElementById("FM_stkpris").value.replace(".",",");
	
	offSet_this = String(document.getElementById("FM_beloeb").value);
	offSetL_this = String(document.getElementById("FM_beloeb").length);
	pkt_this = offSet_this.indexOf(",");
	if (pkt_this > 0){
		document.getElementById("FM_beloeb").value = offSet_this.substring(0, pkt_this + 3)
	}
	
	}
	
	
	
	function opdaterFakdato(){
	document.getElementById("FM_start_dag").value = document.getElementById("FM_interval_slutdag").value
	document.getElementById("FM_start_mrd").value = document.getElementById("FM_interval_slutmrd").value
	document.getElementById("FM_start_aar").value = document.getElementById("FM_interval_slutaar").value
	}
	
	
	
	function addbr(x,a) {
	
	//var selText=document.getElementById("FM_aktbesk_"+x+"_"+a).selection.createRange();
	//alert(selText)
	document.getElementById("FM_aktbesk_"+x+"_"+a).value = document.getElementById("FM_aktbesk_"+x+"_"+a).value + "<br>"
	}
	
	
	
	
	//function matkomma(m, felt){
	//document.getElementById("FM_"+felt+"_"+m).value = document.getElementById("FM_"+felt+"_"+m).value.replace(".",",")
	//}
	
	function beregnmatpris(m,beregntot) {
	antal = document.getElementById("FM_matantal_"+m).value.replace(",",".")
	enhedspris = document.getElementById("FM_matenhedspris_"+m).value.replace(",",".")
	rabat = document.getElementById("FM_matrabat_"+m).value
	
	pris = parseFloat(enhedspris*antal) - parseFloat(enhedspris * antal * rabat)
	
	document.getElementById("FM_matbeloeb_"+m).value = pris 
	document.getElementById("FM_matbeloeb_"+m).value = document.getElementById("FM_matbeloeb_"+m).value.replace(".",",")
	document.getElementById("matbelobdiv_"+m).innerHTML = document.getElementById("FM_matbeloeb_"+m).value
	
	//alert(enhedspris +" % "+ rabat)            
	    if (beregntot == 1){        
	    materialerTot()
	    }            
	}
	
	function materialerTot() { 
	var matBelobialt = 0;
	var matBelobialtHTML = 0;
	antal_m = document.getElementById("FM_antal_materialer_ialt").value
	            
	            for (m=0;m<antal_m;m++){
	                if (document.getElementById("FM_vis_"+m+"").checked == true){
	                matBelobialt = parseFloat(matBelobialt) + parseFloat(document.getElementById("FM_matbeloeb_"+m).value.replace(",","."))
	                }
	            }
	            
	         
	            document.getElementById("FM_materialer_beloeb").value = matBelobialt
	            document.getElementById("FM_materialer_beloeb").value = document.getElementById("FM_materialer_beloeb").value.replace(".",",")
	            matBelobialtHTML = document.getElementById("FM_materialer_beloeb").value.replace(".",",")
	            
	            document.getElementById("divmatbelobtot").innerHTML = "<b>"+matBelobialtHTML+"</b>"
	            
	            document.getElementById("FM_beloeb").value = parseFloat(matBelobialt) + parseFloat(document.getElementById("FM_timer_beloeb").value) 
	            document.getElementById("FM_beloeb").value = document.getElementById("FM_beloeb").value.replace(".",",")
	            belobialtHTML = document.getElementById("FM_beloeb").value
	            document.getElementById("divbelobtot").innerHTML = "<b>"+belobialtHTML+"</b>"
	            
	}
	
	function hidebordermat(m) {
	document.getElementById("beregn_"+m+"_m").style.borderColor = ""
	}
	
	function showbordermat(m) {
	document.getElementById("beregn_"+m+"_m").style.borderColor = "red"
	}
	
	                
	                function offsetmatant(m){
	                    if (window.event.keyCode == 37){ 
			            } else {
				            if (window.event.keyCode == 39){
				            } else {
	                                    offSet_this = String(document.getElementById("FM_matantal_"+m+"").value.replace(".",","));
					                    offSetL_this = String(document.getElementById("FM_matantal_"+m+"").length);
					                    pkt_this = offSet_this.indexOf(",");
					                    if (pkt_this > 0){
						                    document.getElementById("FM_matantal_"+m+"").value = offSet_this.substring(0, pkt_this + 3)
					                    }
					                }
	                            }
					}
					
					
					function offsetmatenhpris(m){
	                    if (window.event.keyCode == 37){ 
			            } else {
				            if (window.event.keyCode == 39){
				            } else {
	                                    offSet_this = String(document.getElementById("FM_matenhedspris_"+m+"").value.replace(".",","));
					                    offSetL_this = String(document.getElementById("FM_matenhedspris_"+m+"").length);
					                    pkt_this = offSet_this.indexOf(",");
					                    if (pkt_this > 0){
						                    document.getElementById("FM_matenhedspris_"+m+"").value = offSet_this.substring(0, pkt_this + 3)
					                    }
					                }
	                            }
					}
					
					
	
	function tjekmatantal(m){
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
	                    var oprAntalthisTj = 0;
	                    oprAntalthisTj = (document.getElementById("FM_matantal_opr_"+m+"").value.replace(",",".") / 1);
		                    if (!validZip(document.getElementById("FM_matantal_"+m+"").value)){
		                    alert("Der er angivet et ugyldigt tegn.")
		                    document.getElementById("FM_matantal_"+m+"").value = oprAntalthisTj;
		                    document.getElementById("FM_matantal_"+m+"").focus()
		                    document.getElementById("FM_matantal_"+m+"").select()
		                    return false
		                    }
		                    document.getElementById("FM_matantal_opr_"+m+"").value = document.getElementById("FM_matantal_"+m+"").value
	                    return true
	            }
	        }
	}
	
	function tjekmatehpris(m){
	if (window.event.keyCode == 37){ 
			} else {
				if (window.event.keyCode == 39){
				} else {
	                    var oprAntalthisTj = 0;
	                    oprAntalthisTj = (document.getElementById("FM_matenhedspris_opr_"+m+"").value.replace(",",".") / 1);
		                    if (!validZip(document.getElementById("FM_matenhedspris_"+m+"").value)){
		                    alert("Der er angivet et ugyldigt tegn.")
		                    document.getElementById("FM_matenhedspris_"+m+"").value = oprAntalthisTj;
		                    document.getElementById("FM_matenhedspris_"+m+"").focus()
		                    document.getElementById("FM_matenhedspris_"+m+"").select()
		                    return false
		                    }
		                    document.getElementById("FM_matenhedspris_opr_"+m+"").value = document.getElementById("FM_matenhedspris_"+m+"").value
	                    return true
	            }
	        }
	}
	
	
	
	function showdiv(dennediv){
	
	lastopendiv = document.getElementById("lastopendiv").value
	if (lastopendiv != 0) {
	document.getElementById(lastopendiv).style.display = "none";
	document.getElementById(lastopendiv).style.visibility = "hidden";
	document.getElementById(lastopendiv+"_2").style.display = "none";
	document.getElementById(lastopendiv+"_2").style.visibility = "hidden";
	if (lastopendiv != "betdiv") {
	document.getElementById("knap_"+lastopendiv).style.background = "#ffffff";
	}
	}
	
	document.getElementById(dennediv).style.display = "";
	document.getElementById(dennediv).style.visibility = "visible";
	
	document.getElementById(dennediv+"_2").style.display = "";
	document.getElementById(dennediv+"_2").style.visibility = "visible";
	
	if (dennediv != "betdiv") {
	document.getElementById("knap_"+dennediv).style.background = "#ffff99";
	}
	
	if (dennediv == "fidiv") {
	document.getElementById("kontodiv").style.display = "";
	document.getElementById("kontodiv").style.visibility = "visible";
	} else {
	document.getElementById("kontodiv").style.display = "none";
	document.getElementById("kontodiv").style.visibility = "hidden";
	
	}
	
	document.getElementById("lastopendiv").value = dennediv
	
	}
	
	



function insertTag(what_type, field) 
{ 
var totalString; //store original string so that we can manipulated 
var selectedString; //the selected string 
var totalLength; //total length of the string 
var selectedLength; //length of the selected area 
var indexStart; //index of the string selected 
var indexFinished; //last index of the string selected 
var firstTextValue; //the first part of the string before index 
var lastTextValue; //the last part of the string after index 

totalString = document.getElementById(field).innerText; 
selectedString = document.selection.createRange().text; 
totalLength = document.getElementById(field).innerText.length; 
selectedLength = document.selection.createRange().text.length;
if (selectedLength == 0){
indexStart = totalLength
} else {
indexStart = totalString.indexOf(selectedString);
}

indexFinished = indexStart + selectedLength; 
firstTextValue = totalString.substring(0, indexStart); 
lastTextValue = totalString.substring(indexFinished, totalLength); 
if(what_type == 'b') 
{ 
selectedString = "<b>"+selectedString+"</b>"; 
} 

if(what_type == 'br') 
{ 
selectedString = selectedString+"<br>"; 
} 

if(what_type == 'i') 
{ 
selectedString = "<i>"+selectedString+"</i>"; 
}

if(what_type == 'u') 
{ 
selectedString = "<u>"+selectedString+"</u>"; 
}

document.getElementById(field).innerText = firstTextValue + selectedString + lastTextValue; 

//document.form1.text1.value = firstTextValue + selectedString + lastTextValue; 

//document.form1.text1.value = document.all("BodySpan").innerText; 
//document.form1.submit(); 

} 

 

	
	
	
	