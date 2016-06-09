//Bruges af:
//serviceaft_osigt.asp
//fak_serviceaft_osigt.asp
	
	
	
	
	
	
	
	
	
	//Aktiviteter expand
 	
	function rensaftnr(){
	document.getElementById("FM_aftnr").value = "";
	}
	
	
	
	function serviceaft(thisid, kid, fm_soeg, forny){
		if (thisid == 0){
		window.open("serviceaft.asp?func=opret&kundeid="+kid+"&id=0&FM_soeg="+fm_soeg+"&forny="+forny+"","TimeOut","width=800,height=780,left=300,top=20,screenX=100,screenY=100,scrollbars=yes")
		} else {
        window.open("serviceaft.asp?func=red&kundeid=" + kid + "&id=" + thisid + "&FM_soeg=" + fm_soeg + "&forny=" + forny + "", "TimeOut", "width=800,height=780,left=300,top=20,screenX=100,screenY=100,scrollbars=yes")
		}
	}
	
	
	
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
	

	
	function checkAll(field)
	{
	field.checked = true;
	for (i = 0; i < field.length; i++)
		field[i].checked = true ;
	}
	
	function uncheckAll(field)
	{
	field.checked = false;
	for (i = 0; i < field.length; i++)
		field[i].checked = false ;
	}