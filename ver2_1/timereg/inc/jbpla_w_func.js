
	
	currentX = currentY = 0; 
	whichEl = null; 
	
	
	function bookdivon(){
	whichEl = event.srcElement;
	
	whichEl.style.width = 30
	whichEl.style.pixelLeft = whichEl.offsetLeft;
    whichEl.style.pixelTop = whichEl.offsetTop;
	
	
	whichEl.style.pixelLeft = whichEl.style.pixelLeft - 10
	
	//whichEl.style.zindex = "50000px";

    currentX = (event.clientX + document.body.scrollLeft);
    currentY = (event.clientY + document.body.scrollTop); 
	}
	
	
	function bookdivmove(){
	if (whichEl == null) { return };

    newX = (event.clientX + document.body.scrollLeft);
    newY = (event.clientY + document.body.scrollTop);
	
    distanceX = (newX - currentX);
    distanceY = (newY - currentY);
    currentX = newX;
    currentY = newY;
	
    whichEl.style.pixelLeft += distanceX;
    whichEl.style.pixelTop += distanceY;
			
	event.returnValue = false;
	}
	
	
	function bookdivoff(mid,u){
	var mid = mid;
	var u = u;
	whichEl.style.width = 8
	var agree = confirm("Vil du flytte opgaven hertil? \n \n (Hvis denne opgave er en del af en serie og du ønsker at flytte hele serien,\n klik på cancel og benyt rediger funktionen istedet.)");
	if (agree == true) {
		bookdivoff2(mid,u)
		} else {
		whichEl.style.pixelLeft = parseInt(document.getElementById("bookdiv_x_"+mid+"_"+u+"").value) + parseInt(document.getElementById("faktor_"+mid+"_"+u+"").value)
		whichEl.style.pixelTop = document.getElementById("bookdiv_y_"+mid+"_"+u+"").value
		//whichEl.style.pixelLeft = parseInt(whichEl.style.pixelLeft) + parseInt(document.getElementById("faktor_"+mid+"_"+u+"").value)
		whichEl = null;
		}
	}	
	
	
	function bookdivoff2(mid,u){
	if (whichEl == null) { return };
	var basefeltnr = 0;
	var WeekdayVal = 0;

	
			if (whichEl.style.pixelLeft <= 280){
			whichEl.style.pixelLeft = 202
			basefeltnr = 1
			
			} else {
					if (whichEl.style.pixelLeft <= 360){
					whichEl.style.pixelLeft = 282
					basefeltnr = 39
					
					} else {
					if (whichEl.style.pixelLeft <= 440){
					whichEl.style.pixelLeft = 362
					basefeltnr = 77
					
					} else {
					if (whichEl.style.pixelLeft <= 520){
					whichEl.style.pixelLeft = 442
					basefeltnr = 115
					
					} else {
					if (whichEl.style.pixelLeft <= 600){
					whichEl.style.pixelLeft = 522
					basefeltnr = 153
					
					} else {
					if (whichEl.style.pixelLeft <= 682){
					whichEl.style.pixelLeft = 602
					basefeltnr = 191
					
					} else {
					if (whichEl.style.pixelLeft > 682){
					whichEl.style.pixelLeft = 682
					basefeltnr = 229
					
					}
			  		}
			  		}
			  		}
			  		}
			  		}
			  }
			 
			 
			  
					//alert(whichEl.style.pixelTop)
					if (whichEl.style.pixelTop <= 68){
					whichEl.style.pixelTop = 57
					basefeltnr = basefeltnr
					} else {
					
					if (whichEl.style.pixelTop <= 83){
					whichEl.style.pixelTop = 72
					basefeltnr = basefeltnr + 1
					} else {
					
					if (whichEl.style.pixelTop <= 98){
					whichEl.style.pixelTop = 87
					basefeltnr = basefeltnr + 2
					} else {
					if (whichEl.style.pixelTop <= 113){
					whichEl.style.pixelTop = 102
					basefeltnr = basefeltnr + 3
					} else {
					if (whichEl.style.pixelTop <= 128){
					whichEl.style.pixelTop = 117
					basefeltnr = basefeltnr + 4
					} else {
					if (whichEl.style.pixelTop <= 143){
					whichEl.style.pixelTop = 132
					basefeltnr = basefeltnr + 5
					} else {
					if (whichEl.style.pixelTop <= 158){
					whichEl.style.pixelTop = 147
					basefeltnr = basefeltnr + 6
					} else {
					if (whichEl.style.pixelTop <= 173){
					whichEl.style.pixelTop = 162
					basefeltnr = basefeltnr + 7
					} else {
					if (whichEl.style.pixelTop <= 188){
					whichEl.style.pixelTop = 177
					basefeltnr = basefeltnr + 8
					} else {
					
					if (whichEl.style.pixelTop <= 203){
					whichEl.style.pixelTop = 192
					basefeltnr = basefeltnr + 9
					} else {
					if (whichEl.style.pixelTop <= 218){
					whichEl.style.pixelTop = 207
					basefeltnr = basefeltnr + 10
					} else {
					if (whichEl.style.pixelTop <= 233){
					whichEl.style.pixelTop = 222
					basefeltnr = basefeltnr + 11
					} else {
					if (whichEl.style.pixelTop <= 248){
					whichEl.style.pixelTop = 237
					basefeltnr = basefeltnr + 12
					} else {
					if (whichEl.style.pixelTop <= 263){
					whichEl.style.pixelTop = 252
					basefeltnr = basefeltnr + 13
					} else {
					if (whichEl.style.pixelTop <= 278){
					whichEl.style.pixelTop = 267
					basefeltnr = basefeltnr + 14
					} else {
					if (whichEl.style.pixelTop <= 293){
					whichEl.style.pixelTop = 282
					basefeltnr = basefeltnr + 15
					} else {
					if (whichEl.style.pixelTop <= 308){
					whichEl.style.pixelTop = 297
					basefeltnr = basefeltnr + 16
					} else {
					if (whichEl.style.pixelTop <= 323){
					whichEl.style.pixelTop = 312
					basefeltnr = basefeltnr + 17
					} else {
					if (whichEl.style.pixelTop <= 338){
					whichEl.style.pixelTop = 327
					basefeltnr = basefeltnr + 18
					} else {
					if (whichEl.style.pixelTop <= 353){
					whichEl.style.pixelTop = 342
					basefeltnr = basefeltnr + 19
					} else {
					if (whichEl.style.pixelTop <= 368){
					whichEl.style.pixelTop = 357
					basefeltnr = basefeltnr + 20
					} else {
					if (whichEl.style.pixelTop <= 383){
					whichEl.style.pixelTop = 372
					basefeltnr = basefeltnr + 21
					} else {
					if (whichEl.style.pixelTop <= 398){
					whichEl.style.pixelTop = 387
					basefeltnr = basefeltnr + 22
					} else {
					if (whichEl.style.pixelTop <= 413){
					whichEl.style.pixelTop = 402
					basefeltnr = basefeltnr + 23
					} else {
					if (whichEl.style.pixelTop <= 428){
					whichEl.style.pixelTop = 417
					basefeltnr = basefeltnr + 24
					} else {
					if (whichEl.style.pixelTop <= 443){
					whichEl.style.pixelTop = 432
					basefeltnr = basefeltnr + 25
					} else {
					if (whichEl.style.pixelTop <= 458){
					whichEl.style.pixelTop = 447
					basefeltnr = basefeltnr + 26
					} else {
					if (whichEl.style.pixelTop <= 473){
					whichEl.style.pixelTop = 462
					basefeltnr = basefeltnr + 27
					} else {
					if (whichEl.style.pixelTop <= 488){
					whichEl.style.pixelTop = 477
					basefeltnr = basefeltnr + 28
					} else {
					if (whichEl.style.pixelTop <= 503){
					whichEl.style.pixelTop = 492
					basefeltnr = basefeltnr + 29
					} else {
					if (whichEl.style.pixelTop <= 518){
					whichEl.style.pixelTop = 507
					basefeltnr = basefeltnr + 30
					} else {
					if (whichEl.style.pixelTop <= 533){
					whichEl.style.pixelTop = 522
					basefeltnr = basefeltnr + 31
					} else {
					if (whichEl.style.pixelTop <= 548){
					whichEl.style.pixelTop = 537
					basefeltnr = basefeltnr + 32
					} else {
					if (whichEl.style.pixelTop <= 563){
					whichEl.style.pixelTop = 552
					basefeltnr = basefeltnr + 33
					} else {
					if (whichEl.style.pixelTop <= 578){
					whichEl.style.pixelTop = 567
					basefeltnr = basefeltnr + 34
					} else {
					if (whichEl.style.pixelTop <= 593){
					whichEl.style.pixelTop = 582
					basefeltnr = basefeltnr + 35
					} else {
					if (whichEl.style.pixelTop <= 608){
					whichEl.style.pixelTop = 597
					basefeltnr = basefeltnr + 36
					} else {
					if (whichEl.style.pixelTop <= 623){
					whichEl.style.pixelTop = 612
					basefeltnr = basefeltnr + 37
					} else {
					if (whichEl.style.pixelTop <= 638){
					whichEl.style.pixelTop = 627
					basefeltnr = basefeltnr + 38
					} else {
					if (whichEl.style.pixelTop <= 653){
					whichEl.style.pixelTop = 642
					basefeltnr = basefeltnr + 39
					} else {
					if (whichEl.style.pixelTop <= 668){
					whichEl.style.pixelTop = 657
					basefeltnr = basefeltnr + 40
					} else {
					if (whichEl.style.pixelTop <= 683){
					whichEl.style.pixelTop = 672
					basefeltnr = basefeltnr + 41
					} else {
					if (whichEl.style.pixelTop <= 698){
					whichEl.style.pixelTop = 687
					basefeltnr = basefeltnr + 42
					} else {
					if (whichEl.style.pixelTop <= 2713){
					whichEl.style.pixelTop = 702
					basefeltnr = basefeltnr + 43
					} else {
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					}
					
					event.returnValue = false;
					}
		
	
	//alert("l:" + whichEl.style.pixelLeft + " t:" + whichEl.style.pixelTop)		
	document.getElementById("bookdiv_x_"+mid+"_"+u+"").value = whichEl.style.pixelLeft	
	document.getElementById("bookdiv_y_"+mid+"_"+u+"").value = whichEl.style.pixelTop
	
	var thisdato = 0;
	var bundval = 0;
	var	topval = 280;
	st = 1
	sl = 7
	
	for (i=st;i<=sl;i++){
	
		if (i == 1) {
		bundval = 0
		topval = 280
		}
		
		if (i == 2) {
		bundval = 280
		topval = 360
		}	
		
		if (i == 3) {
		bundval = 360
		topval = 440
		}	
		
		if (i == 4) {
		bundval = 440
		topval = 520
		}	
		
		if (i == 5) {
		bundval = 520
		topval = 600
		}	
		
		if (i == 6) {
		bundval = 600
		topval = 680
		}	
		
		if (i == 7) {
		bundval = 680
		topval = 760
		}			
					
		if (parseInt(document.getElementById("bookdiv_x_"+mid+"_"+u+"").value) > parseInt(bundval)) {
			if(parseInt(document.getElementById("bookdiv_x_"+mid+"_"+u+"").value) <= parseInt(topval)) {
			thisdato = document.getElementById("datoweekday"+i).value
			//alert(thisdato)
			}
		} 
	}
	
	
	//Placerer div efter flytning
	//Hvis felt er forskelliget fra det felt div kom fra tjekkes hvor mange div'er der allerede er i det nye felt.
	//alert(basefeltnr +"#"+ feltbaseorg)
	feltbaseorg = document.getElementById("bookdiv_feltbasenr_"+mid+"_"+u+"").value
	if (basefeltnr != feltbaseorg) {
		whichEl.style.pixelLeft = parseInt(document.getElementById("bookdiv_x_"+mid+"_"+u+"").value) + parseInt((document.getElementById("felt_"+basefeltnr).value) * 8)
		document.getElementById("felt_"+basefeltnr).value = parseInt(document.getElementById("felt_"+basefeltnr).value) + 1	
		document.getElementById("bookdiv_feltbasenr_"+mid+"_"+u+"").value = basefeltnr
		//document.getElementById("faktor_"+mid+"_"+u+"").value = parseInt(document.getElementById("felt_"+basefeltnr).value) + 1 * 8
		
		
		//Sætter værdier inden db opdateres 
		monththis = document.getElementById("mdselect").value
		yearthis = document.getElementById("aarselect").value
		aktuge = document.getElementById("aktuge").value
		antalu = document.getElementById("antalu").value
		
		divid = document.getElementById("bookdiv_id_"+mid+"_"+u+"").value 
		divxval = document.getElementById("bookdiv_x_"+mid+"_"+u+"").value
		divyval = document.getElementById("bookdiv_y_"+mid+"_"+u+"").value
		divheight = document.getElementById("bookdiv_height_"+mid+"_"+u+"").value
		
		window.location = "jbpla_w.asp?menu=res&func=reddb&monththis="+monththis+"&yearthis="+yearthis+"&activeuge="+aktuge+"&divid="+divid+"&divxval="+divxval+"&divyval="+divyval+"&divheight="+divheight+"&thisdato="+thisdato
		
	
	} else {
		whichEl.style.pixelLeft = parseInt(whichEl.style.pixelLeft) + parseInt(document.getElementById("faktor_"+mid+"_"+u+"").value)
	}	
	
	whichEl = null;
	}
	
	
	
	
	function cursormove() {
    	event.srcElement.style.cursor = "move"
   	}
	
	function cursorhand() {
    	event.srcElement.style.cursor = "crosshair" //hand
   	}
	
	function bookminus(mid,u){
	var currentH = 0;
	currentH = parseInt(document.getElementById("bookdiv_"+mid+"_"+u+"").style.height)
		if (currentH != 15) {
		document.getElementById("bookdiv_"+mid+"_"+u+"").style.height = parseInt(currentH - 15)
		document.getElementById("bookdiv_height_"+mid+"_"+u+"").value = document.getElementById("bookdiv_"+mid+"_"+u+"").style.height
		document.getElementById("bookdiv_y_"+mid+"_"+u+"").value = document.getElementById("bookdiv_"+mid+"_"+u+"").style.pixelTop 
		document.getElementById("bookdiv_x_"+mid+"_"+u+"").value = document.getElementById("bookdiv_"+mid+"_"+u+"").style.pixelLeft
		}
	}
	
	function bookplus(mid,u){
	var currentH = 0;
	currentH = parseInt(document.getElementById("bookdiv_"+mid+"_"+u+"").style.height)
		if (currentH < 400) {
		document.getElementById("bookdiv_"+mid+"_"+u+"").style.height = parseInt(currentH + 15)
		document.getElementById("bookdiv_height_"+mid+"_"+u+"").value = document.getElementById("bookdiv_"+mid+"_"+u+"").style.height
		document.getElementById("bookdiv_y_"+mid+"_"+u+"").value = document.getElementById("bookdiv_"+mid+"_"+u+"").style.pixelTop 
		document.getElementById("bookdiv_x_"+mid+"_"+u+"").value = document.getElementById("bookdiv_"+mid+"_"+u+"").style.pixelLeft
		}
	}
	
	function showinfo(u){
	//alert("her")
	document.getElementById("info_"+u).style.visibility = "visible"
	document.getElementById("info_"+u).style.display = ""
	}
	
	function hideinfo(u){
	document.getElementById("info_"+u).style.visibility = "hidden"
	document.getElementById("info_"+u).style.display = "none"
	}
	
	//function showmidbook(mid) {
	//var uval = 0;
	//var lastopen = document.getElementById("showlastmid").value
	//uval = document.getElementById("antalu").value
		
		
		//lukker alle
		//for (i=1;i<=uval;i++){
			//if (parseInt(document.getElementById("u_this_"+i+"").value) == parseInt(lastopen)) {
			//	document.getElementById("bookdiv_"+lastopen+"_"+i).style.visibility = "hidden"
			//	document.getElementById("bookdiv_"+lastopen+"_"+i).style.display = "none"
			//}
		//}
		
		//åbner den aktive
		//for (i=1;i<=uval;i++){
			//if (parseInt(document.getElementById("u_this_"+i+"").value) == parseInt(mid)) {
			//	document.getElementById("bookdiv_"+mid+"_"+i).style.visibility = "visible"
			//	document.getElementById("bookdiv_"+mid+"_"+i).style.display = ""
			//}
		//}
		
		
		//document.getElementById("selmedarbtd_"+lastopen+"").style.borderColor = "#eff3ff"
		//document.getElementById("selmedarbtd_"+mid+"").style.borderColor = "#000000"
		
	//document.getElementById("showlastmid").value = mid
	//}
	
	
	
	function NewWin_opr(url)    {
		window.open(url, 'Opretny', 'left=450, top=280, width=400,height=280,scrollbars=yes,toolbar=no');    
	}
	
	//function felter(felt){
	//document.getElementById("felt_"+felt+"").value = 5
	//}
	
	//function popUp(URL,width,height) {
	//	window.open(URL, 'navn', 'left=250,top=180,toolbar=0,scrollbars=1,location=0,statusbar=1,menubar=0,resizable=1,width=' + width + ',height=' + height + '');
	//}
