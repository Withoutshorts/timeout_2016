var OxOacff=["0123456789ABCDEF","0000","all","getElementById","","|","fond","\x3Cimg src=\x22../Images/multiclavier.gif\x22 width=404 height=152 border=\x220\x22\x3E\x3Cbr /\x3E","sign","car","simpledia","simple","majus","\x26nbsp;","double","minus","doubledia","kb-","kb+","Delete","Clear","Back","CapsLock","Enter","Shift","\x3C|\x3C","Space","\x3E|\x3E","clavscroll(-3)","clavscroll(3)","faire(\x22del\x22)","RAZ()","faire(\x22bck\x22)","bloq()","faire(\x22\x5Cn\x22)","haut()","faire(\x22ar\x22)","faire(\x22 \x22)","faire(\x22av\x22)","act","action","clav","clavier","masque","\x3Cimg src=\x22../Images/1x1.gif\x22 width=404 height=152 border=\x220\x22 usemap=\x22#clavier\x22\x3E","\x3Cmap name=\x22clavier\x22\x3E","\x3Carea coords=\x22",",","\x22 href=\x22javascript:void(0)\x22 onClick=\x27javascript:ecrire(",")\x27\x3E","\x22 href=\x22javascript:void(0)\x22 onClick=\x27javascript:","\x27\x3E","\x22 href=\x27javascript:charger(","\x3C/map\x3E","length"," ","%0D%0A","del","bck","ar","av","\x3Cspan class=","\x3E","\x3C/span\x3E","\x3Cdiv id=\x22","\x22 \x3E","\x3C/div\x3E","left","style","px","top","innerHTML","act0","act1","langue=","cookie",";","langue","=","; ","expires="];var caps=0x0,lock=0x0,hexchars=OxOacff[0x0],accent=OxOacff[0x1],clavdeb=0x0;var clav= new Array(); j=0x0 ;for( i  in Maj){ clav[j]=i ; j++ ;} ;var ns6=((!document[OxOacff[0x2]])&&(document[OxOacff[0x3]]));var ie=document[OxOacff[0x2]];var langue=getCk();if(langue==OxOacff[0x4]){ langue=clav[clavdeb] ;} ; CarMaj=Maj[langue].split(OxOacff[0x5]) ; CarMin=Min[langue].split(OxOacff[0x5]) ;var posClavierLeft=0x0,posClavierTop=0x0;if(ns6){ posClavierLeft=0x0 ; posClavierTop=0x50 ;} else {if(ie){ posClavierLeft=0x0 ; posClavierTop=0x50 ;} ;} ; tracer(OxOacff[0x6],posClavierLeft,posClavierTop,OxOacff[0x7],OxOacff[0x8]) ;var posX= new Array(0x0,0x1c,0x38,0x54,0x70,0x8c,0xa8,0xc4,0xe0,0xfc,0x118,0x134,0x150,0x2a,0x46,0x62,0x7e,0x9a,0xb6,0xd2,0xee,0x10a,0x126,0x142,0x15e,0x32,0x4e,0x6a,0x86,0xa2,0xbe,0xda,0xf6,0x112,0x12e,0x14a,0x40,0x5c,0x78,0x94,0xb0,0xcc,0xe8,0x104,0x120,0x13c,0x1c,0x38,0x54,0x126,0x142,0x15e);var posY= new Array(0xe,0xe,0xe,0xe,0xe,0xe,0xe,0xe,0xe,0xe,0xe,0xe,0xe,0x2a,0x2a,0x2a,0x2a,0x2a,0x2a,0x2a,0x2a,0x2a,0x2a,0x2a,0x2a,0x46,0x46,0x46,0x46,0x46,0x46,0x46,0x46,0x46,0x46,0x46,0x62,0x62,0x62,0x62,0x62,0x62,0x62,0x62,0x62,0x62,0x7e,0x7e,0x7e,0x7e,0x7e,0x7e);var nbTouches=0x34;for( i=0x0 ;i<nbTouches;i++){ CarMaj[i]=((CarMaj[i]!=OxOacff[0x1])?(fromhexby4tocar(CarMaj[i])):OxOacff[0x4]) ; CarMin[i]=((CarMin[i]!=OxOacff[0x1])?(fromhexby4tocar(CarMin[i])):OxOacff[0x4]) ;if(CarMaj[i]==CarMin[i].toUpperCase()){ cecar=((lock==0x0)&&(caps==0x0)?CarMin[i]:CarMaj[i]) ; tracer(OxOacff[0x9]+i,posClavierLeft+0x6+posX[i],posClavierTop+0x3+posY[i],cecar,((dia[hexa(cecar)]!=null)?OxOacff[0xa]:OxOacff[0xb])) ; tracer(OxOacff[0xc]+i,posClavierLeft+0xf+posX[i],posClavierTop+0x1+posY[i],OxOacff[0xd],OxOacff[0xe]) ; tracer(OxOacff[0xf]+i,posClavierLeft+0x3+posX[i],posClavierTop+0x9+posY[i],OxOacff[0xd],OxOacff[0xe]) ;} else { tracer(OxOacff[0x9]+i,posClavierLeft+0x6+posX[i],posClavierTop+0x3+posY[i],OxOacff[0xd],OxOacff[0xb]) ; cecar=CarMin[i] ; tracer(OxOacff[0xf]+i,posClavierLeft+0x3+posX[i],posClavierTop+0x9+posY[i],cecar,((dia[hexa(cecar)]!=null)?OxOacff[0x10]:OxOacff[0xe])) ; cecar=CarMaj[i] ; tracer(OxOacff[0xc]+i,posClavierLeft+0xf+posX[i],posClavierTop+0x1+posY[i],cecar,((dia[hexa(cecar)]!=null)?OxOacff[0x10]:OxOacff[0xe])) ;} ;} ;var actC1= new Array(0x0,0x173,0x16c,0x0,0x17a,0x0,0x166,0x0,0x158,0x0,0x70,0x17a);var actC2= new Array(0x0,0x0,0xe,0x2a,0x2a,0x46,0x46,0x62,0x62,0x7e,0x7e,0x7e);var actC3= new Array(0x20,0x193,0x193,0x27,0x193,0x2f,0x193,0x3d,0x193,0x19,0x123,0x193);var actC4= new Array(0xb,0xb,0x27,0x43,0x43,0x5f,0x5f,0x7b,0x7b,0x97,0x97,0x97);var act= new Array(OxOacff[0x11],OxOacff[0x12],OxOacff[0x13],OxOacff[0x14],OxOacff[0x15],OxOacff[0x16],OxOacff[0x17],OxOacff[0x18],OxOacff[0x18],OxOacff[0x19],OxOacff[0x1a],OxOacff[0x1b]);var effet= new Array(OxOacff[0x1c],OxOacff[0x1d],OxOacff[0x1e],OxOacff[0x1f],OxOacff[0x20],OxOacff[0x21],OxOacff[0x22],OxOacff[0x23],OxOacff[0x23],OxOacff[0x24],OxOacff[0x25],OxOacff[0x26]);var nbActions=0xc;for( i=0x0 ;i<nbActions;i++){ tracer(OxOacff[0x27]+i,posClavierLeft+0x1+actC1[i],posClavierTop-0x1+actC2[i],act[i],OxOacff[0x28]) ;} ;var clavC1= new Array(0x23,0x77,0xcb,0x11f);var clavC2= new Array(0x0,0x0,0x0,0x0);var clavC3= new Array(0x74,0xc8,0x11c,0x170);var clavC4= new Array(0xb,0xb,0xb,0xb);for( i=0x0 ;i<0x4;i++){ tracer(OxOacff[0x29]+i,posClavierLeft+0x5+clavC1[i],posClavierTop-0x1+clavC2[i],clav[i],OxOacff[0x2a]) ;} ; tracer(OxOacff[0x2b],posClavierLeft,posClavierTop,OxOacff[0x2c]) ; document.write(OxOacff[0x2d]) ;for( i=0x0 ;i<nbTouches;i++){ document.write(OxOacff[0x2e]+posX[i]+OxOacff[0x2f]+posY[i]+OxOacff[0x2f]+(posX[i]+0x19)+OxOacff[0x2f]+(posY[i]+0x19)+OxOacff[0x30]+i+OxOacff[0x31]) ;} ;for( i=0x0 ;i<nbActions;i++){ document.write(OxOacff[0x2e]+actC1[i]+OxOacff[0x2f]+actC2[i]+OxOacff[0x2f]+actC3[i]+OxOacff[0x2f]+actC4[i]+OxOacff[0x32]+effet[i]+OxOacff[0x33]) ;} ;for( i=0x0 ;i<0x4;i++){ document.write(OxOacff[0x2e]+clavC1[i]+OxOacff[0x2f]+clavC2[i]+OxOacff[0x2f]+clavC3[i]+OxOacff[0x2f]+clavC4[i]+OxOacff[0x34]+i+OxOacff[0x31]) ;} ; document.write(OxOacff[0x35]) ; function ecrire(i){ txt=rechercher()+OxOacff[0x5] ; subtxt=txt.split(OxOacff[0x5]) ; ceci=(lock==0x1)?CarMaj[i]:((caps==0x1)?CarMaj[i]:CarMin[i]) ;if(test(ceci)){ subtxt[0x0]+=cardia(ceci) ; distinguer(false) ;} else {if(dia[accent]!=null&&dia[hexa(ceci)]!=null){ distinguer(false) ; accent=hexa(ceci) ; distinguer(true) ;} else {if(dia[accent]!=null){ subtxt[0x0]+=fromhexby4tocar(accent)+ceci ; distinguer(false) ;} else {if(dia[hexa(ceci)]!=null){ accent=hexa(ceci) ; distinguer(true) ;} else { subtxt[0x0]+=ceci ;} ;} ;} ;} ; txt=subtxt[0x0]+OxOacff[0x5]+subtxt[0x1] ; afficher(txt) ;if(caps==0x1){ caps=0x0 ; MinusMajus() ;} ;}  ; function faire(Ox9aa){ txt=rechercher()+OxOacff[0x5] ; subtxt=txt.split(OxOacff[0x5]) ; l0=subtxt[0x0][OxOacff[0x36]] ; l1=subtxt[0x1][OxOacff[0x36]] ; c1=subtxt[0x0].substring(0x0,(l0-0x2)) ; c2=subtxt[0x0].substring(0x0,(l0-0x1)) ; c3=subtxt[0x1].substring(0x0,0x1) ; c4=subtxt[0x1].substring(0x0,0x2) ; c5=subtxt[0x0].substring((l0-0x2),l0) ; c6=subtxt[0x0].substring((l0-0x1),l0) ; c7=subtxt[0x1].substring(0x1,l1) ; c8=subtxt[0x1].substring(0x2,l1) ;if(dia[accent]!=null){if(Ox9aa==OxOacff[0x37]){ Ox9aa=fromhexby4tocar(accent) ;} ; distinguer(false) ;} ;switch(Ox9aa){case (OxOacff[0x3c]):if(escape(c4)!=OxOacff[0x38]){ txt=subtxt[0x0]+c3+OxOacff[0x5]+c7 ;} else { txt=subtxt[0x0]+c4+OxOacff[0x5]+c8 ;} ;break ;case (OxOacff[0x3b]):if(escape(c5)!=OxOacff[0x38]){ txt=c2+OxOacff[0x5]+c6+subtxt[0x1] ;} else { txt=c1+OxOacff[0x5]+c5+subtxt[0x1] ;} ;break ;case (OxOacff[0x3a]):if(escape(c5)!=OxOacff[0x38]){ txt=c2+OxOacff[0x5]+subtxt[0x1] ;} else { txt=c1+OxOacff[0x5]+subtxt[0x1] ;} ;break ;case (OxOacff[0x39]):if(escape(c4)!=OxOacff[0x38]){ txt=subtxt[0x0]+OxOacff[0x5]+c7 ;} else { txt=subtxt[0x0]+OxOacff[0x5]+c8 ;} ;break ;default: txt=subtxt[0x0]+Ox9aa+OxOacff[0x5]+subtxt[0x1] ;break ;;;;;;} ; afficher(txt) ;}  ; function RAZ(){ txt=OxOacff[0x4] ;if(dia[accent]!=null){ distinguer(false) ;} ; afficher(txt) ;}  ; function haut(){ caps=0x1 ; MinusMajus() ;}  ; function bloq(){ lock=(lock==0x1)?0x0:0x1 ; MinusMajus() ;}  ; function tracer(Ox9af,Ox9b0,haut,Ox9aa,Ox9b1){ Ox9aa=OxOacff[0x3d]+Ox9b1+OxOacff[0x3e]+Ox9aa+OxOacff[0x3f] ; document.write(OxOacff[0x40]+Ox9af+OxOacff[0x41]+Ox9aa+OxOacff[0x42]) ;if(ns6){ document.getElementById(Ox9af)[OxOacff[0x44]][OxOacff[0x43]]=Ox9b0+OxOacff[0x45] ; document.getElementById(Ox9af)[OxOacff[0x44]][OxOacff[0x46]]=haut+OxOacff[0x45] ;} else {if(ie){ document.all(Ox9af)[OxOacff[0x44]][OxOacff[0x43]]=Ox9b0 ; document.all(Ox9af)[OxOacff[0x44]][OxOacff[0x46]]=haut ;} ;} ;}  ; function retracer(Ox9af,Ox9aa,Ox9b1){ Ox9aa=OxOacff[0x3d]+Ox9b1+OxOacff[0x3e]+Ox9aa+OxOacff[0x3f] ;if(ns6){ document.getElementById(Ox9af)[OxOacff[0x47]]=Ox9aa ;} else {if(ie){ doc=document.all(Ox9af) ; doc[OxOacff[0x47]]=Ox9aa ;} ;} ;}  ; function clavscroll(n){ clavdeb+=n ;if(clavdeb<0x0){ clavdeb=0x0 ;} ;if(clavdeb>clav[OxOacff[0x36]]-0x4){ clavdeb=clav[OxOacff[0x36]]-0x4 ;} ;for( i=clavdeb ;i<clavdeb+0x4;i++){ retracer(OxOacff[0x29]+(i-clavdeb),clav[i],OxOacff[0x2a]) ;} ;if(clavdeb==0x0){ retracer(OxOacff[0x48],OxOacff[0xd],OxOacff[0x28]) ;} else { retracer(OxOacff[0x48],act[0x0],OxOacff[0x28]) ;} ;if(clavdeb==clav[OxOacff[0x36]]-0x4){ retracer(OxOacff[0x49],OxOacff[0xd],OxOacff[0x28]) ;} else { retracer(OxOacff[0x49],act[0x1],OxOacff[0x28]) ;} ;}  ; function charger(i){ langue=clav[i+clavdeb] ; setCk(langue) ; accent=OxOacff[0x1] ; CarMaj=Maj[langue].split(OxOacff[0x5]) ; CarMin=Min[langue].split(OxOacff[0x5]) ;for( i=0x0 ;i<nbTouches;i++){ CarMaj[i]=((CarMaj[i]!=OxOacff[0x1])?(fromhexby4tocar(CarMaj[i])):OxOacff[0x4]) ; CarMin[i]=((CarMin[i]!=OxOacff[0x1])?(fromhexby4tocar(CarMin[i])):OxOacff[0x4]) ;if(CarMaj[i]==CarMin[i].toUpperCase()){ cecar=((lock==0x0)&&(caps==0x0)?CarMin[i]:CarMaj[i]) ; retracer(OxOacff[0x9]+i,cecar,((dia[hexa(cecar)]!=null)?OxOacff[0xa]:OxOacff[0xb])) ; retracer(OxOacff[0xf]+i,OxOacff[0xd]) ; retracer(OxOacff[0xc]+i,OxOacff[0xd]) ;} else { retracer(OxOacff[0x9]+i,OxOacff[0xd]) ; cecar=CarMin[i] ; retracer(OxOacff[0xf]+i,cecar,((dia[hexa(cecar)]!=null)?OxOacff[0x10]:OxOacff[0xe])) ; cecar=CarMaj[i] ; retracer(OxOacff[0xc]+i,cecar,((dia[hexa(cecar)]!=null)?OxOacff[0x10]:OxOacff[0xe])) ;} ;} ;}  ; function distinguer(Ox9b6){for( i=0x0 ;i<nbTouches;i++){if(CarMaj[i]==CarMin[i].toUpperCase()){ cecar=((lock==0x0)&&(caps==0x0)?CarMin[i]:CarMaj[i]) ;if(test(cecar)){ retracer(OxOacff[0x9]+i,Ox9b6?(cardia(cecar)):cecar,Ox9b6?OxOacff[0xa]:OxOacff[0xb]) ;} ;} else { cecar=CarMin[i] ;if(test(cecar)){ retracer(OxOacff[0xf]+i,Ox9b6?(cardia(cecar)):cecar,Ox9b6?OxOacff[0x10]:OxOacff[0xe]) ;} ; cecar=CarMaj[i] ;if(test(cecar)){ retracer(OxOacff[0xc]+i,Ox9b6?(cardia(cecar)):cecar,Ox9b6?OxOacff[0x10]:OxOacff[0xe]) ;} ;} ;} ;if(!Ox9b6){ accent=OxOacff[0x1] ;} ;}  ; function MinusMajus(){for( i=0x0 ;i<nbTouches;i++){if(CarMaj[i]==CarMin[i].toUpperCase()){ cecar=((lock==0x0)&&(caps==0x0)?CarMin[i]:CarMaj[i]) ; retracer(OxOacff[0x9]+i,(test(cecar)?cardia(cecar):cecar),((dia[hexa(cecar)]!=null||test(cecar))?OxOacff[0xa]:OxOacff[0xb])) ;} ;} ;}  ; function test(Ox9b9){return (dia[accent]!=null&&dia[accent][hexa(Ox9b9)]!=null);}  ; function cardia(Ox9b9){return (fromhexby4tocar(dia[accent][hexa(Ox9b9)]));}  ; function fromhex(Ox9bc){ out=0x0 ;for( a=Ox9bc[OxOacff[0x36]]-0x1 ;a>=0x0;a--){ out+=Math.pow(0x10,Ox9bc[OxOacff[0x36]]-a-0x1)*hexchars.indexOf(Ox9bc.charAt(a)) ;} ;return out;}  ; function fromhexby4tocar(Ox9aa){ out4= new String() ;for( l=0x0 ;l<Ox9aa[OxOacff[0x36]];l+=0x4){ out4+=String.fromCharCode(fromhex(Ox9aa.substring(l,l+0x4))) ;} ;return out4;}  ; function tohex(Ox9bc){return hexchars.charAt(Ox9bc/0x10)+hexchars.charAt(Ox9bc%0x10);}  ; function tohex2(Ox9bc){return tohex(Ox9bc/0x100)+tohex(Ox9bc%0x100);}  ; function hexa(Ox9aa){ out=OxOacff[0x4] ;for( k=0x0 ;k<Ox9aa[OxOacff[0x36]];k++){ out+=(tohex2(Ox9aa.charCodeAt(k))) ;} ;return out;}  ; function getCk(){ fromN=document[OxOacff[0x4b]].indexOf(OxOacff[0x4a])+0x0 ;if((fromN)!=-0x1){ fromN+=0x7 ; toN=document[OxOacff[0x4b]].indexOf(OxOacff[0x4c],fromN)+0x0 ;if(toN==-0x1){ toN=document[OxOacff[0x4b]][OxOacff[0x36]] ;} ;return unescape(document[OxOacff[0x4b]].substring(fromN,toN));} ;return OxOacff[0x4];}  ; function setCk(Ox9bc){if(Ox9bc!=null){ exp= new Date() ; time=0x16d*0x3c*0x3c*0x18*0x3e8 ; exp.setTime(exp.getTime()+time) ; document[OxOacff[0x4b]]=escape(OxOacff[0x4d])+OxOacff[0x4e]+escape(Ox9bc)+OxOacff[0x4f]+OxOacff[0x50]+exp.toGMTString() ;} ;}  ;