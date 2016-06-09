

//function showdivakt(rloop, thisHTML){ 
//obsampletxt = document.getElementById("Menu"+rloop);
//strHTML = "<table width='100%' border='0' cellspacing='1' cellpadding='0' bgcolor='White'>" + thisHTML + "</table>"; 

//obsampletxt.insertAdjacentHTML("BeforeEnd", strHTML);
//document.getElementById("Menu"+rloop).style.display = "";
//}


//function visf5(){
//	window.opener.document.all["f5"].innerHTML = "<b>Status:</b> <font color=red>Opdater siden [F5].</font>"; 
//} 


function checkAll(field)
	{
	for (i = 0; i < field.length; i++)
		field[i].checked = true ;
	}
	
	function uncheckAll(field)
	{
	for (i = 0; i < field.length; i++)
		field[i].checked = false ;
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


function setTimerTot(nummer, dagtype, job, aktivitet) {
	
	//piletaster
	
	if (window.event.keyCode == 37){ 
	} else {
		if (window.event.keyCode == 39){
		//alert(window.event.keyCode)
		} else {
		
	
		var varValue = 0;
		var varValue_total = 0;
		
		document.getElementById("FM_"+ dagtype +"_opr_" + nummer + "").value = document.getElementById("FM_"+ dagtype +"_opr_" + nummer + "").value.replace(",",".")
		oprVerdi = (document.getElementById("FM_"+ dagtype +"_opr_" + nummer + "").value / 1);
		
		document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value = document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value.replace(",",".")
		varValue = (document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value / 1);
		document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value = document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value.replace(".",",")
		
		document.getElementById("FM_"+ dagtype +"_total").value = document.getElementById("FM_"+ dagtype +"_total").value.replace(",",".")
		varValue_total = (document.getElementById("FM_"+ dagtype +"_total").value / 1); 
	 	
		if (varValue > 24) {
		alert("En time-indtastning må ikke overstige 24 timer. \n Der er angivet: " + varValue + " timer.")
		document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value = document.getElementById("FM_"+ dagtype +"_opr_" + nummer + "").value.replace(".",",") //oprVerdi;
		document.getElementById("FM_"+ dagtype +"_total").value = document.getElementById("FM_"+ dagtype +"_total").value.replace(".",",")
		}
		else {
			varTillaeg = (varValue - oprVerdi);
			varTotal_dag_beg = (varValue_total + varTillaeg);
				
			if (varTotal_dag_beg > 24) {
			alert("Et døgn indeholder kun 24 timer!!")
			document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value = parseFloat(oprVerdi);
			}
			
			if (varTotal_dag_beg <= 24){
			document.getElementById("FM_"+ dagtype +"_total").value = parseFloat(varTotal_dag_beg);
			document.getElementById("FM_"+ dagtype +"_opr_" + nummer + "").value = parseFloat(varValue);
			sonValue = document.getElementById("FM_son_total").value.replace(",",".")/1;
			manValue = document.getElementById("FM_man_total").value.replace(",",".")/1;
			tirValue = document.getElementById("FM_tir_total").value.replace(",",".")/1;
			onsValue = document.getElementById("FM_ons_total").value.replace(",",".")/1;
			torValue = document.getElementById("FM_tor_total").value.replace(",",".")/1;
			freValue = document.getElementById("FM_fre_total").value.replace(",",".")/1;
			lorValue = document.getElementById("FM_lor_total").value.replace(",",".")/1;
			document.getElementById("FM_week_total").value = parseFloat(sonValue + manValue + tirValue + onsValue + torValue + freValue + lorValue) 
			document.getElementById("FM_week_total").value = document.getElementById("FM_week_total").value.replace(".",",")
			document.getElementById("FM_"+ dagtype +"_total").value = document.getElementById("FM_"+ dagtype +"_total").value.replace(".",",")
			}
		}
	}} // piletaster
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

function tjektimer(dagtype, nummer){
	if (!validZip(document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value)){
	alert("Der er angivet et ugyldigt tegn.")
	document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value = oprVerdi;
	document.getElementById("Timer_"+ dagtype +"_" + nummer + "").focus()
	document.getElementById("Timer_"+ dagtype +"_" + nummer + "").select()
	return false
	}
return true
}

function tjekkm(dagtype, nummer){
	oprVerdi = (document.getElementById("FM_"+ dagtype +"_opr_" + nummer + "").value / 1);
	if (!validZip(document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value)){
	alert("Der er angivet et ugyldigt tegn.")
	document.getElementById("Timer_"+ dagtype +"_" + nummer + "").value = oprVerdi;
	document.getElementById("Timer_"+ dagtype +"_" + nummer + "").focus()
	document.getElementById("Timer_"+ dagtype +"_" + nummer + "").select()
	return false
	}
return true
}

	
	
	//Aktiviteter expand
 	if (document.images){
		plus = new Image(200, 200);
		plus.src = "ill/plus.gif";
		minus = new Image(200, 200);
		minus.src = "ill/minus2.gif";
		}

		function expand(de){
		//alert(navigator.appVersion.indexOf("MSIE"))
		    if (navigator.appVersion.indexOf("MSIE")==-1){
			//alert("hej")
				//if (document.getElementById("Menu" + de)){
				if (document.getElementById("Menu" + de).style.display == "none"){
					document.getElementById("Menu" + de).style.display = "";
					document.images["Menub" + de].src = minus.src;
				}else{
					document.getElementById("Menu" + de).style.display = "none";
					document.images["Menub" + de].src = plus.src;
				} // else
			}else{
				//alert("pc")
				if (document.all("Menu" + de)){
				if (document.all("Menu" + de).style.display == "none"){
					document.all("Menu" + de).style.display = "";
					document.images["Menub" + de].src = minus.src;
				}else{
					document.all("Menu" + de).style.display = "none";
					document.images["Menub" + de].src = plus.src;
				} //else
			} //else
		} //function
	} //images
	
	
	//Kommentar expand
	function expandkomm(dagtype, nummer){
	
				thisval = document.getElementById("FM_kom_"+ dagtype +"_"+ nummer +"").value
				thisvaloff = document.getElementById("FM_off_"+ dagtype +"_"+ nummer +"").value
				document.getElementById("rowcounter").value = nummer;
				document.getElementById("daytype").value = dagtype;
				document.getElementById("kom").style.display = "";
				document.getElementById("kom").style.visibility = "visible";
				document.getElementById("FM_kom").value = thisval;
				document.getElementById("FM_off").value = thisvaloff;
				document.getElementById("kom").style.top = window.event.offsetY + 200;
				thisstring = document.getElementById("FM_kom").value.length;
				document.getElementById("antch").value = thisstring;
				document.getElementById("FM_kom").select();
				document.getElementById("FM_kom").focus();
				
	}
	
	function closekomm(){
		//if (navigator.appVersion.indexOf("MSIE")==-1){
			thisstring = document.getElementById("FM_kom").value.length
			thisvalue = document.getElementById("FM_kom").value
			thisrow = document.getElementById("rowcounter").value
			thisdaytype = document.getElementById("daytype").value
			thisoffentlig = document.getElementById("FM_off").value
			
			if (thisstring > 255){
			alert("Der er ikke tilladt mere end 255 karakterer i en kommentar! \n Du har brugt: "+ thisstring +"")
			}else{
			document.getElementById("FM_kom_"+thisdaytype+"_"+thisrow+"").value = thisvalue;
			document.getElementById("FM_off_"+thisdaytype+"_"+thisrow+"").value = thisoffentlig;
			document.getElementById("FM_off").value = 0
			document.getElementById("kom").style.display = "none";
			//document.getElementById("Timer_"+thisdaytype+"_"+thisrow+"").style.backgroundColor = "#ffffe1";
			}
	}
	
	function antalchar(){
	//if (navigator.appVersion.indexOf("MSIE")==-1){
		thisstring = document.getElementById("FM_kom").value.length
		if (thisstring > 255){
			alert("Der er ikke tilladt mere end 255 karakterer i en kommentar! \n Du har brugt: "+ thisstring +"")
		}else{
			document.getElementById("antch").value = thisstring;
		}
	}
	

 function antalakt(jid){
 newval = document.all["FM_akt_" + jid].value;
 document.all["FM_antal_akt_" + jid].value = newval;
 }

 function closetomanyjob() {
 document.all["tomanyjob"].style.visibility = "hidden";
 }
 
  ///  ressource timer tildelt ///////
 	function showrestimer(){
	document.getElementById("ressourcetimer").style.display = "";
	document.getElementById("ressourcetimer").style.visibility = "visible";
	}
	
	function hiderestimer(){
	document.getElementById("ressourcetimer").style.display = "none";
	document.getElementById("ressourcetimer").style.visibility = "hidden";
	}
	
 ///  viser indtastede timer på den enkelte dag ///////
 	function showtimedetail(thisnameid){
	document.getElementById("ressourcetimer").style.display = "none";
	document.getElementById("ressourcetimer").style.visibility = "hidden";
	
	document.getElementById("timedetailson").style.display = "none";
	document.getElementById("timedetailson").style.visibility = "hidden";
	
	document.getElementById("timedetailman").style.display = "none";
	document.getElementById("timedetailman").style.visibility = "hidden";
	
	document.getElementById("timedetailtir").style.display = "none";
	document.getElementById("timedetailtir").style.visibility = "hidden";
	
	document.getElementById("timedetailons").style.display = "none";
	document.getElementById("timedetailons").style.visibility = "hidden";
	
	document.getElementById("timedetailtor").style.display = "none";
	document.getElementById("timedetailtor").style.visibility = "hidden";
	
	document.getElementById("timedetailfre").style.display = "none";
	document.getElementById("timedetailfre").style.visibility = "hidden";
	
	document.getElementById("timedetaillor").style.display = "none";
	document.getElementById("timedetaillor").style.visibility = "hidden";
	
	document.getElementById(thisnameid).style.display = "";
	document.getElementById(thisnameid).style.visibility = "visible";
	}
	
	function hidetimedetail(thisid){
	document.getElementById("timedetailson").style.display = "none";
	document.getElementById("timedetailson").style.visibility = "hidden";
	
	document.getElementById("timedetailman").style.display = "none";
	document.getElementById("timedetailman").style.visibility = "hidden";
	
	document.getElementById("timedetailtir").style.display = "none";
	document.getElementById("timedetailtir").style.visibility = "hidden";
	
	document.getElementById("timedetailons").style.display = "none";
	document.getElementById("timedetailons").style.visibility = "hidden";
	
	document.getElementById("timedetailtor").style.display = "none";
	document.getElementById("timedetailtor").style.visibility = "hidden";
	
	document.getElementById("timedetailfre").style.display = "none";
	document.getElementById("timedetailfre").style.visibility = "hidden";
	
	document.getElementById("timedetaillor").style.display = "none";
	document.getElementById("timedetaillor").style.visibility = "hidden";
	}
	
	
	function showmilepal(){
	document.getElementById("divmilepal").style.display = "";
	document.getElementById("divmilepal").style.visibility = "visible";
	
	document.images("knapopgliste").src = "../ill/knap_opglisten_off.gif";
	document.images("knapmilepal").src = "../ill/knap_milepal_on.gif";
	document.images("knaptimer").src = "../ill/knap_timer_off.gif";
	
	document.getElementById("divopgave").style.display = "none";
	document.getElementById("divopgave").style.visibility = "hidden";
	
	document.getElementById("vistimerformedarb").style.display = "none";
	document.getElementById("vistimerformedarb").style.visibility = "hidden";
	}
	
	function showopgave(){
	document.getElementById("divmilepal").style.display = "none";
	document.getElementById("divmilepal").style.visibility = "hidden";
	
	document.images("knapopgliste").src = "../ill/knap_opglisten_on.gif";
	document.images("knapmilepal").src = "../ill/knap_milepal_off.gif";
	document.images("knaptimer").src = "../ill/knap_timer_off.gif";
	
	document.getElementById("divopgave").style.display = "";
	document.getElementById("divopgave").style.visibility = "visible";
	
	document.getElementById("vistimerformedarb").style.display = "none";
	document.getElementById("vistimerformedarb").style.visibility = "hidden";
	}
	
	function showtimer(){
	document.getElementById("divmilepal").style.display = "none";
	document.getElementById("divmilepal").style.visibility = "hidden";
	
	document.getElementById("divopgave").style.display = "none";
	document.getElementById("divopgave").style.visibility = "hidden";
	
	document.images("knapopgliste").src = "../ill/knap_opglisten_off.gif";
	document.images("knapmilepal").src = "../ill/knap_milepal_off.gif";
	document.images("knaptimer").src = "../ill/knap_timer_on.gif";
	
	document.getElementById("vistimerformedarb").style.display = "";
	document.getElementById("vistimerformedarb").style.visibility = "visible";
	}