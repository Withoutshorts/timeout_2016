var OxO8e13=["Verdana","innerHTML","Unicode","innerText","\x3Cspan style=\x27font-family:","\x27\x3E","\x3C/span\x3E","selfont","length","checked","value","charstable1","charstable2","fontFamily","style","display","block","none"];var editor=Window_GetDialogArguments(window); function getchar(obj){var Ox83;var Ox84=getFontValue()||OxO8e13[0x0];if(!obj[OxO8e13[0x1]]){return ;} ; Ox83=obj[OxO8e13[0x1]] ;if(Ox84==OxO8e13[0x2]){ Ox83=obj[OxO8e13[0x3]] ;} else {if(Ox84!=OxO8e13[0x0]){ Ox83=OxO8e13[0x4]+Ox84+OxO8e13[0x5]+obj[OxO8e13[0x1]]+OxO8e13[0x6] ;} ;} ; editor.PasteHTML(Ox83) ; Window_CloseDialog(window) ;}  ; function cancel(){ Window_CloseDialog(window) ;}  ; function getFontValue(){var Ox87=document.getElementsByName(OxO8e13[0x7]);for(var i=0x0;i<Ox87[OxO8e13[0x8]];i++){if(Ox87.item(i)[OxO8e13[0x9]]){return Ox87.item(i)[OxO8e13[0xa]];} ;} ;}  ; function sel_font_change(){var Ox89=getFontValue()||OxO8e13[0x0];var Ox2d5=Window_GetElement(window,OxO8e13[0xb],true);var Ox2d6=Window_GetElement(window,OxO8e13[0xc],true); Ox2d5[OxO8e13[0xe]][OxO8e13[0xd]]=Ox89 ; Ox2d5[OxO8e13[0xe]][OxO8e13[0xf]]=(Ox89!=OxO8e13[0x2]?OxO8e13[0x10]:OxO8e13[0x11]) ; Ox2d6[OxO8e13[0xe]][OxO8e13[0xf]]=(Ox89==OxO8e13[0x2]?OxO8e13[0x10]:OxO8e13[0x11]) ;}  ;