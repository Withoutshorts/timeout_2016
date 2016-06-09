
	
	
	
	currentX = currentY = 0; 
	whichEl = null; 
	
	function bookdivon(){
	whichEl = event.srcElement;
	
	whichEl.style.pixelLeft = whichEl.offsetLeft;
    whichEl.style.pixelTop = whichEl.offsetTop;

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
	if (whichEl == null) { return };
	//alert(whichEl.style.pixelLeft)
			//Grid
			if (whichEl.style.pixelLeft <= 77){
			whichEl.style.pixelLeft = 54
			} else {
					if (whichEl.style.pixelLeft <= 120){
					whichEl.style.pixelLeft = 98
					} else {
					if (whichEl.style.pixelLeft <= 164){
					whichEl.style.pixelLeft = 142
					} else {
					if (whichEl.style.pixelLeft <= 207){
					whichEl.style.pixelLeft = 186
					} else {
					if (whichEl.style.pixelLeft <= 252){
					whichEl.style.pixelLeft = 230
					} else {
					if (whichEl.style.pixelLeft <= 294){
					whichEl.style.pixelLeft = 274
					} else {
					if (whichEl.style.pixelLeft > 294){
					whichEl.style.pixelLeft = 318
					}
			  		}
			  		}
			  		}
			  		}
			  		}
			  }
			
	
	//alert(whichEl.style.pixelTop)
	if (whichEl.style.pixelTop <= 188){
	whichEl.style.pixelTop = 177
	} else {
	if (whichEl.style.pixelTop <= 203){
	whichEl.style.pixelTop = 192
	} else {
	if (whichEl.style.pixelTop <= 218){
	whichEl.style.pixelTop = 207
	} else {
	if (whichEl.style.pixelTop <= 233){
	whichEl.style.pixelTop = 222
	} else {
	if (whichEl.style.pixelTop <= 248){
	whichEl.style.pixelTop = 237
	} else {
	if (whichEl.style.pixelTop <= 263){
	whichEl.style.pixelTop = 252
	} else {
	if (whichEl.style.pixelTop <= 278){
	whichEl.style.pixelTop = 267
	} else {
	if (whichEl.style.pixelTop <= 293){
	whichEl.style.pixelTop = 282
	} else {
	if (whichEl.style.pixelTop <= 308){
	whichEl.style.pixelTop = 297
	} else {
	if (whichEl.style.pixelTop <= 323){
	whichEl.style.pixelTop = 312
	} else {
	if (whichEl.style.pixelTop <= 338){
	whichEl.style.pixelTop = 327
	} else {
	if (whichEl.style.pixelTop <= 353){
	whichEl.style.pixelTop = 342
	} else {
	if (whichEl.style.pixelTop <= 368){
	whichEl.style.pixelTop = 357
	} else {
	if (whichEl.style.pixelTop <= 383){
	whichEl.style.pixelTop = 372
	} else {
	if (whichEl.style.pixelTop <= 398){
	whichEl.style.pixelTop = 387
	} else {
	if (whichEl.style.pixelTop <= 413){
	whichEl.style.pixelTop = 402
	} else {
	if (whichEl.style.pixelTop <= 428){
	whichEl.style.pixelTop = 417
	} else {
	if (whichEl.style.pixelTop <= 443){
	whichEl.style.pixelTop = 432
	} else {
	if (whichEl.style.pixelTop <= 458){
	whichEl.style.pixelTop = 447
	} else {
	if (whichEl.style.pixelTop <= 473){
	whichEl.style.pixelTop = 462
	} else {
	if (whichEl.style.pixelTop <= 488){
	whichEl.style.pixelTop = 477
	} else {
	if (whichEl.style.pixelTop <= 503){
	whichEl.style.pixelTop = 492
	} else {
	if (whichEl.style.pixelTop <= 518){
	whichEl.style.pixelTop = 507
	} else {
	if (whichEl.style.pixelTop <= 533){
	whichEl.style.pixelTop = 522
	} else {
	if (whichEl.style.pixelTop <= 548){
	whichEl.style.pixelTop = 537
	} else {
	if (whichEl.style.pixelTop <= 563){
	whichEl.style.pixelTop = 552
	} else {
	if (whichEl.style.pixelTop <= 578){
	whichEl.style.pixelTop = 567
	} else {
	if (whichEl.style.pixelTop <= 593){
	whichEl.style.pixelTop = 582
	} else {
	if (whichEl.style.pixelTop <= 608){
	whichEl.style.pixelTop = 597
	} else {
	if (whichEl.style.pixelTop <= 623){
	whichEl.style.pixelTop = 612
	} else {
	if (whichEl.style.pixelTop <= 638){
	whichEl.style.pixelTop = 627
	} else {
	if (whichEl.style.pixelTop <= 653){
	whichEl.style.pixelTop = 642
	} else {
	if (whichEl.style.pixelTop <= 668){
	whichEl.style.pixelTop = 657
	} else {
	if (whichEl.style.pixelTop <= 683){
	whichEl.style.pixelTop = 672
	} else {
	if (whichEl.style.pixelTop <= 698){
	whichEl.style.pixelTop = 687
	} else {
	if (whichEl.style.pixelTop <= 2713){
	whichEl.style.pixelTop = 702
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
	
	event.returnValue = false;
	}
			
	
	document.getElementById("bookdiv_x_"+mid+"_"+u+"").value = whichEl.style.pixelLeft	
	document.getElementById("bookdiv_y_"+mid+"_"+u+"").value = whichEl.style.pixelTop	
	whichEl = null;				
	}
	
	function cursormove() {
    	event.srcElement.style.cursor = "move"
   	}
	
	function cursorhand() {
    	event.srcElement.style.cursor = "hand"
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
	document.getElementById("info_"+u).style.visibility = "visible"
	document.getElementById("info_"+u).style.display = ""
	}
	
	function hideinfo(u){
	document.getElementById("info_"+u).style.visibility = "hidden"
	document.getElementById("info_"+u).style.display = "none"
	}
	
	function showmidbook(mid) {
	var uval = 0;
	var lastopen = document.getElementById("showlastmid").value
	uval = document.getElementById("antalu").value
		
		
		//lukker alle
		for (i=1;i<=uval;i++){
			if (parseInt(document.getElementById("u_this_"+i+"").value) == parseInt(lastopen)) {
				document.getElementById("bookdiv_"+lastopen+"_"+i).style.visibility = "hidden"
				document.getElementById("bookdiv_"+lastopen+"_"+i).style.display = "none"
			}
		}
		
		//åbner den aktive
		for (i=1;i<=uval;i++){
			if (parseInt(document.getElementById("u_this_"+i+"").value) == parseInt(mid)) {
				document.getElementById("bookdiv_"+mid+"_"+i).style.visibility = "visible"
				document.getElementById("bookdiv_"+mid+"_"+i).style.display = ""
			}
		}
	document.getElementById("showlastmid").value = mid
	}
	
	function NewWin_opr(url)    {
		window.open(url, 'Opretny', 'left=550, top=280, width=400,height=280,scrollbars=yes,toolbar=no');    
	}
