var OxO2d33=["inp_width","eenheid","alignment","hrcolor","hrcolorpreview","shade","sel_size","width","style","value","px","%","size","align","color","backgroundColor","noShade","noshade","","onclick"];var inp_width=Window_GetElement(window,OxO2d33[0],true);var eenheid=Window_GetElement(window,OxO2d33[1],true);var alignment=Window_GetElement(window,OxO2d33[2],true);var hrcolor=Window_GetElement(window,OxO2d33[3],true);var hrcolorpreview=Window_GetElement(window,OxO2d33[4],true);var shade=Window_GetElement(window,OxO2d33[5],true);var sel_size=Window_GetElement(window,OxO2d33[6],true);UpdateState=function UpdateState_Hr(){} ;SyncToView=function SyncToView_Hr(){if(element[OxO2d33[8]][OxO2d33[7]]){if(element[OxO2d33[8]][OxO2d33[7]].search(/%/)<0){eenheid[OxO2d33[9]]=OxO2d33[10];inp_width[OxO2d33[9]]=element[OxO2d33[8]][OxO2d33[7]].split(OxO2d33[10])[0];} else {eenheid[OxO2d33[9]]=OxO2d33[11];inp_width[OxO2d33[9]]=element[OxO2d33[8]][OxO2d33[7]].split(OxO2d33[11])[0];} ;} ;sel_size[OxO2d33[9]]=element[OxO2d33[12]];alignment[OxO2d33[9]]=element[OxO2d33[13]];hrcolor[OxO2d33[9]]=element[OxO2d33[14]];if(element[OxO2d33[14]]){hrcolor[OxO2d33[8]][OxO2d33[15]]=element[OxO2d33[14]];} ;if(element[OxO2d33[16]]){shade[OxO2d33[9]]=OxO2d33[17];} else {shade[OxO2d33[9]]=OxO2d33[18];} ;} ;SyncTo=function SyncTo_Hr(element){if(sel_size[OxO2d33[9]]){element[OxO2d33[12]]=sel_size[OxO2d33[9]];} ;if(hrcolor[OxO2d33[9]]){element[OxO2d33[14]]=hrcolor[OxO2d33[9]];} ;if(alignment[OxO2d33[9]]){element[OxO2d33[13]]=alignment[OxO2d33[9]];} ;if(shade[OxO2d33[9]]==OxO2d33[17]){element[OxO2d33[16]]=true;} else {element[OxO2d33[16]]=false;} ;if(inp_width[OxO2d33[9]]){element[OxO2d33[8]][OxO2d33[7]]=inp_width[OxO2d33[9]]+eenheid[OxO2d33[9]];} ;} ;hrcolor[OxO2d33[19]]=hrcolorpreview[OxO2d33[19]]=function hrcolor_onclick(){SelectColor(hrcolor,hrcolorpreview);} ;