var OxOaeac=["SetStyle","length","","GetStyle","GetText",":",";","cssText","tblBorderStyle","sel_style","sel_border","sel_part","bordercolor","bordercolor_Preview","inp_color","inp_color_Preview","inp_shade","inp_shade_Preview","inp_MarginLeft","inp_MarginRight","inp_MarginTop","inp_MarginBottom","inp_PaddingLeft","inp_PaddingRight","inp_PaddingTop","inp_PaddingBottom","inp_width","sel_width_unit","inp_height","sel_height_unit","inp_id","inp_class","sel_align","sel_textalign","sel_float","inp_tooltip","value","borderColor","style"," ","backgroundColor","color","id","className","width","height","inp_","sel_","_unit","selectedIndex","options","align","styleFloat","cssFloat","textAlign","title","borderWidth","borderBottomWidth","borderLeftWidth","borderRightWidth","borderTopWidth","borderBottomStyle","borderLeftStyle","borderRightStyle","borderTopStyle","none","border","-","red","paddingLeft","paddingRight","paddingTop","paddingBottom","marginLeft","marginRight","marginTop","marginBottom","ValidID","class","onclick"]; function pause(Ox3b4){var Ox330= new Date();var Ox3b5=Ox330.getTime()+Ox3b4;while(true){ Ox330= new Date() ;if(Ox330.getTime()>Ox3b5){return ;} ;} ;}  ; function StyleClass(Ox154){var Ox3b7=[];var Ox3b8={};if(Ox154){ Ox3bd() ;} ; this[OxOaeac[0x0]]=function SetStyle(name,Ox128,Ox3ba){ name=name.toLowerCase() ;for(var i=0x0;i<Ox3b7[OxOaeac[0x1]];i++){if(Ox3b7[i]==name){break ;} ;} ; Ox3b7[i]=name ; Ox3b8[name]=Ox128?(Ox128+(Ox3ba||OxOaeac[0x2])):OxOaeac[0x2] ;}  ; this[OxOaeac[0x3]]=function GetStyle(name){ name=name.toLowerCase() ;return Ox3b8[name]||OxOaeac[0x2];}  ; this[OxOaeac[0x4]]=function Ox3bc(){var Ox154=OxOaeac[0x2];for(var i=0x0;i<Ox3b7[OxOaeac[0x1]];i++){var n=Ox3b7[i];var p=Ox3b8[n];if(p){ Ox154+=n+OxOaeac[0x5]+p+OxOaeac[0x6] ;} ;} ;return Ox154;}  ; function Ox3bd(){var arr=Ox154.split(OxOaeac[0x6]);for(var i=0x0;i<arr[OxOaeac[0x1]];i++){var p=arr[i].split(OxOaeac[0x5]);var n=p[0x0].replace(/^\s+/g,OxOaeac[0x2]).replace(/\s+$/g,OxOaeac[0x2]).toLowerCase(); Ox3b7[Ox3b7[OxOaeac[0x1]]]=n ; Ox3b8[n]=p[0x1] ;} ;}  ;}  ; function GetStyle(Oxf6,name){return  new StyleClass(Oxf6.cssText).GetStyle(name);}  ; function SetStyle(Oxf6,name,Ox128,Ox3be){var Ox3bf= new StyleClass(Oxf6.cssText); Ox3bf.SetStyle(name,Ox128,Ox3be) ; Oxf6[OxOaeac[0x7]]=Ox3bf.GetText() ;}  ; function ParseFloatToString(Ox7d){if(!Ox7d){return OxOaeac[0x2];} ;var Ox5c=parseFloat(Ox7d);if(isNaN(Ox5c)){return OxOaeac[0x2];} ;return Ox5c+OxOaeac[0x2];}  ;var tblBorderStyle=Window_GetElement(window,OxOaeac[0x8],true);var sel_style=Window_GetElement(window,OxOaeac[0x9],true);var sel_border=Window_GetElement(window,OxOaeac[0xa],true);var sel_part=Window_GetElement(window,OxOaeac[0xb],true);var bordercolor=Window_GetElement(window,OxOaeac[0xc],true);var bordercolor_Preview=Window_GetElement(window,OxOaeac[0xd],true);var inp_color=Window_GetElement(window,OxOaeac[0xe],true);var inp_color_Preview=Window_GetElement(window,OxOaeac[0xf],true);var inp_shade=Window_GetElement(window,OxOaeac[0x10],true);var inp_shade_Preview=Window_GetElement(window,OxOaeac[0x11],true);var inp_MarginLeft=Window_GetElement(window,OxOaeac[0x12],true);var inp_MarginRight=Window_GetElement(window,OxOaeac[0x13],true);var inp_MarginTop=Window_GetElement(window,OxOaeac[0x14],true);var inp_MarginBottom=Window_GetElement(window,OxOaeac[0x15],true);var inp_PaddingLeft=Window_GetElement(window,OxOaeac[0x16],true);var inp_PaddingRight=Window_GetElement(window,OxOaeac[0x17],true);var inp_PaddingTop=Window_GetElement(window,OxOaeac[0x18],true);var inp_PaddingBottom=Window_GetElement(window,OxOaeac[0x19],true);var inp_width=Window_GetElement(window,OxOaeac[0x1a],true);var sel_width_unit=Window_GetElement(window,OxOaeac[0x1b],true);var inp_height=Window_GetElement(window,OxOaeac[0x1c],true);var sel_height_unit=Window_GetElement(window,OxOaeac[0x1d],true);var inp_id=Window_GetElement(window,OxOaeac[0x1e],true);var inp_class=Window_GetElement(window,OxOaeac[0x1f],true);var sel_align=Window_GetElement(window,OxOaeac[0x20],true);var sel_textalign=Window_GetElement(window,OxOaeac[0x21],true);var sel_float=Window_GetElement(window,OxOaeac[0x22],true);var inp_tooltip=Window_GetElement(window,OxOaeac[0x23],true); UpdateState=function UpdateState_Div(){}  ; function doBorderStyle(Ox106){ sel_style[OxOaeac[0x24]]=Ox106 ;}  ; function doPart(Ox106){ sel_part[OxOaeac[0x24]]=Ox106 ;}  ; function doWidth(Ox106){ sel_border[OxOaeac[0x24]]=Ox106 ;}  ; SyncToView=function SyncToView_Div(){if(Browser_IsWinIE()){ bordercolor[OxOaeac[0x24]]=element[OxOaeac[0x26]][OxOaeac[0x25]] ;} else {var arr=revertColor(element[OxOaeac[0x26]].borderColor).split(OxOaeac[0x27]); bordercolor[OxOaeac[0x24]]=arr[0x0] ;} ; bordercolor[OxOaeac[0x26]][OxOaeac[0x28]]=bordercolor[OxOaeac[0x24]] ; bordercolor_Preview[OxOaeac[0x26]][OxOaeac[0x28]]=bordercolor[OxOaeac[0x24]] ; inp_color[OxOaeac[0x24]]=revertColor(element[OxOaeac[0x26]].color) ; inp_color[OxOaeac[0x26]][OxOaeac[0x28]]=element[OxOaeac[0x26]][OxOaeac[0x29]] ; inp_color_Preview[OxOaeac[0x26]][OxOaeac[0x28]]=element[OxOaeac[0x26]][OxOaeac[0x29]] ; inp_shade[OxOaeac[0x24]]=revertColor(element[OxOaeac[0x26]].backgroundColor) ; inp_shade[OxOaeac[0x26]][OxOaeac[0x28]]=element[OxOaeac[0x26]][OxOaeac[0x28]] ; inp_shade_Preview[OxOaeac[0x26]][OxOaeac[0x28]]=element[OxOaeac[0x26]][OxOaeac[0x28]] ; inp_id[OxOaeac[0x24]]=element[OxOaeac[0x2a]] ;if(ParseFloatToString(element[OxOaeac[0x26]].marginLeft)){ inp_MarginLeft[OxOaeac[0x24]]=ParseFloatToString(element[OxOaeac[0x26]].marginLeft) ;} ;if(ParseFloatToString(element[OxOaeac[0x26]].marginRight)){ inp_MarginRight[OxOaeac[0x24]]=ParseFloatToString(element[OxOaeac[0x26]].marginRight) ;} ;if(ParseFloatToString(element[OxOaeac[0x26]].marginTop)){ inp_MarginTop[OxOaeac[0x24]]=ParseFloatToString(element[OxOaeac[0x26]].marginTop) ;} ;if(ParseFloatToString(element[OxOaeac[0x26]].marginBottom)){ inp_MarginBottom[OxOaeac[0x24]]=ParseFloatToString(element[OxOaeac[0x26]].marginBottom) ;} ;if(ParseFloatToString(element[OxOaeac[0x26]].paddingLeft)){ inp_PaddingLeft[OxOaeac[0x24]]=ParseFloatToString(element[OxOaeac[0x26]].paddingLeft) ;} ;if(ParseFloatToString(element[OxOaeac[0x26]].paddingRight)){ inp_PaddingRight[OxOaeac[0x24]]=ParseFloatToString(element[OxOaeac[0x26]].paddingRight) ;} ;if(ParseFloatToString(element[OxOaeac[0x26]].paddingTop)){ inp_PaddingTop[OxOaeac[0x24]]=ParseFloatToString(element[OxOaeac[0x26]].paddingTop) ;} ;if(ParseFloatToString(element[OxOaeac[0x26]].paddingBottom)){ inp_PaddingBottom[OxOaeac[0x24]]=ParseFloatToString(element[OxOaeac[0x26]].paddingBottom) ;} ; inp_class[OxOaeac[0x24]]=element[OxOaeac[0x2b]] ;var arr=[OxOaeac[0x2c],OxOaeac[0x2d]];for(var Ox14f=0x0;Ox14f<arr[OxOaeac[0x1]];Ox14f++){var n=arr[Ox14f];var Ox41=Window_GetElement(window,OxOaeac[0x2e]+n,true);var Ox106=Window_GetElement(window,OxOaeac[0x2f]+n+OxOaeac[0x30],true); Ox106[OxOaeac[0x31]]=0x0 ;if(ParseFloatToString(element[OxOaeac[0x26]][n])){ Ox41[OxOaeac[0x24]]=ParseFloatToString(element[OxOaeac[0x26]][n]) ;for(var i=0x0;i<Ox106[OxOaeac[0x32]][OxOaeac[0x1]];i++){var Oxe9=Ox106[OxOaeac[0x32]][i][OxOaeac[0x24]];if(Oxe9&&element[OxOaeac[0x26]][n].indexOf(Oxe9)!=-0x1){ Ox106[OxOaeac[0x31]]=i ;break ;} ;} ;} ;} ; sel_align[OxOaeac[0x24]]=element[OxOaeac[0x33]] ;if(Browser_IsWinIE()){ sel_float[OxOaeac[0x24]]=element[OxOaeac[0x26]][OxOaeac[0x34]] ;} else { sel_float[OxOaeac[0x24]]=element[OxOaeac[0x26]][OxOaeac[0x35]] ;} ; sel_textalign[OxOaeac[0x24]]=element[OxOaeac[0x26]][OxOaeac[0x36]] ; inp_tooltip[OxOaeac[0x24]]=element[OxOaeac[0x37]] ;try{ sel_border[OxOaeac[0x24]]=element[OxOaeac[0x26]][OxOaeac[0x38]] ;if(element[OxOaeac[0x26]][OxOaeac[0x3a]]==element[OxOaeac[0x26]][OxOaeac[0x3c]]&&element[OxOaeac[0x26]][OxOaeac[0x3a]]==element[OxOaeac[0x26]][OxOaeac[0x3b]]&&element[OxOaeac[0x26]][OxOaeac[0x3a]]==element[OxOaeac[0x26]][OxOaeac[0x39]]){ sel_border[OxOaeac[0x24]]=element[OxOaeac[0x26]][OxOaeac[0x3a]] ;} ;if(element[OxOaeac[0x26]][OxOaeac[0x3e]]==element[OxOaeac[0x26]][OxOaeac[0x40]]&&element[OxOaeac[0x26]][OxOaeac[0x3e]]==element[OxOaeac[0x26]][OxOaeac[0x3f]]&&element[OxOaeac[0x26]][OxOaeac[0x3e]]==element[OxOaeac[0x26]][OxOaeac[0x3d]]){ sel_style[OxOaeac[0x24]]=element[OxOaeac[0x26]][OxOaeac[0x3e]] ;} ;} catch(x){} ;}  ; SyncTo=function SyncTo_Div(element){var Ox3d8=sel_part[OxOaeac[0x24]];if(Ox3d8==OxOaeac[0x41]){ element[OxOaeac[0x26]][OxOaeac[0x42]]=OxOaeac[0x41] ;} else {var Ox3d9=Ox3d8?OxOaeac[0x43]+Ox3d8:Ox3d8;var Oxe4=OxOaeac[0x44];var Oxf6=sel_style[OxOaeac[0x24]]||OxOaeac[0x2];var Ox3da=sel_border[OxOaeac[0x24]]; element[OxOaeac[0x26]][OxOaeac[0x42]]=OxOaeac[0x41] ;if(Ox3da||Oxf6){ SetStyle(element[OxOaeac[0x26]],OxOaeac[0x42]+Ox3d9,Ox3da+OxOaeac[0x27]+Oxf6+OxOaeac[0x27]+Oxe4) ;} else { SetStyle(element[OxOaeac[0x26]],OxOaeac[0x42]+Ox3d9,OxOaeac[0x2]) ;} ; SetStyle(element[OxOaeac[0x26]],OxOaeac[0x42]+Ox3d9,Ox3da+OxOaeac[0x27]+Oxf6+OxOaeac[0x27]+Oxe4) ;} ;try{ element[OxOaeac[0x26]][OxOaeac[0x29]]=inp_color[OxOaeac[0x24]]||OxOaeac[0x2] ;} catch(x){ element[OxOaeac[0x26]][OxOaeac[0x29]]=OxOaeac[0x2] ;} ;try{ element[OxOaeac[0x26]][OxOaeac[0x28]]=inp_shade[OxOaeac[0x24]]||OxOaeac[0x2] ;} catch(x){ element[OxOaeac[0x26]][OxOaeac[0x28]]=OxOaeac[0x2] ;} ;try{ element[OxOaeac[0x26]][OxOaeac[0x25]]=bordercolor[OxOaeac[0x24]]||OxOaeac[0x2] ;} catch(x){ element[OxOaeac[0x26]][OxOaeac[0x25]]=OxOaeac[0x2] ;} ; element[OxOaeac[0x26]][OxOaeac[0x45]]=inp_PaddingLeft[OxOaeac[0x24]] ; element[OxOaeac[0x26]][OxOaeac[0x46]]=inp_PaddingRight[OxOaeac[0x24]] ; element[OxOaeac[0x26]][OxOaeac[0x47]]=inp_PaddingTop[OxOaeac[0x24]] ; element[OxOaeac[0x26]][OxOaeac[0x48]]=inp_PaddingBottom[OxOaeac[0x24]] ; element[OxOaeac[0x26]][OxOaeac[0x49]]=inp_MarginLeft[OxOaeac[0x24]] ; element[OxOaeac[0x26]][OxOaeac[0x4a]]=inp_MarginRight[OxOaeac[0x24]] ; element[OxOaeac[0x26]][OxOaeac[0x4b]]=inp_MarginTop[OxOaeac[0x24]] ; element[OxOaeac[0x26]][OxOaeac[0x4c]]=inp_MarginBottom[OxOaeac[0x24]] ; element[OxOaeac[0x2b]]=inp_class[OxOaeac[0x24]] ;var arr=[OxOaeac[0x2c],OxOaeac[0x2d]];for(var Ox14f=0x0;Ox14f<arr[OxOaeac[0x1]];Ox14f++){var n=arr[Ox14f];var Ox41=Window_GetElement(window,OxOaeac[0x2e]+n,true);var Ox3db=Window_GetElement(window,OxOaeac[0x2f]+n+OxOaeac[0x30],true);if(ParseFloatToString(Ox41.value)){ element[OxOaeac[0x26]][n]=ParseFloatToString(Ox41.value)+Ox3db[OxOaeac[0x24]] ;} else { element[OxOaeac[0x26]][n]=OxOaeac[0x2] ;} ;} ;var Ox2ce=/[^a-z\d]/i;if(Ox2ce.test(inp_id.value)){ alert(CE_GetStr(OxOaeac[0x4d])) ;return ;} ; element[OxOaeac[0x33]]=sel_align[OxOaeac[0x24]] ; element[OxOaeac[0x2a]]=inp_id[OxOaeac[0x24]] ;if(Browser_IsWinIE()){ element[OxOaeac[0x26]][OxOaeac[0x34]]=sel_float[OxOaeac[0x24]] ;} else { element[OxOaeac[0x26]][OxOaeac[0x35]]=sel_float[OxOaeac[0x24]] ;} ; element[OxOaeac[0x26]][OxOaeac[0x36]]=sel_textalign[OxOaeac[0x24]] ; element[OxOaeac[0x37]]=inp_tooltip[OxOaeac[0x24]] ;if(element[OxOaeac[0x37]]==OxOaeac[0x2]){ element.removeAttribute(OxOaeac[0x37]) ;} ;if(element[OxOaeac[0x2b]]==OxOaeac[0x2]){ element.removeAttribute(OxOaeac[0x2b]) ;} ;if(element[OxOaeac[0x2b]]==OxOaeac[0x2]){ element.removeAttribute(OxOaeac[0x4e]) ;} ;if(element[OxOaeac[0x33]]==OxOaeac[0x2]){ element.removeAttribute(OxOaeac[0x33]) ;} ;if(element[OxOaeac[0x2a]]==OxOaeac[0x2]){ element.removeAttribute(OxOaeac[0x2a]) ;} ;}  ; bordercolor[OxOaeac[0x4f]]=bordercolor_Preview[OxOaeac[0x4f]]=function bordercolor_onclick(){ SelectColor(bordercolor,bordercolor_Preview) ;}  ; inp_color[OxOaeac[0x4f]]=inp_color_Preview[OxOaeac[0x4f]]=function inp_color_onclick(){ SelectColor(inp_color,inp_color_Preview) ;}  ; inp_shade[OxOaeac[0x4f]]=inp_shade_Preview[OxOaeac[0x4f]]=function inp_shade_onclick(){ SelectColor(inp_shade,inp_shade_Preview) ;}  ;