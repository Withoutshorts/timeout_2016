var OxO58a4=["inp_width","inp_height","sel_align","sel_valign","inp_bgColor","inp_borderColor","inp_borderColorLight","inp_borderColorDark","inp_class","inp_id","inp_tooltip","sel_noWrap","sel_CellScope","value","bgColor","backgroundColor","style","","id","borderColor","borderColorLight","borderColorDark","className","width","height","align","vAlign","title","noWrap","scope","ValidNumber","ValidID","class","valign","onclick"];var inp_width=Window_GetElement(window,OxO58a4[0],true);var inp_height=Window_GetElement(window,OxO58a4[1],true);var sel_align=Window_GetElement(window,OxO58a4[2],true);var sel_valign=Window_GetElement(window,OxO58a4[3],true);var inp_bgColor=Window_GetElement(window,OxO58a4[4],true);var inp_borderColor=Window_GetElement(window,OxO58a4[5],true);var inp_borderColorLight=Window_GetElement(window,OxO58a4[6],true);var inp_borderColorDark=Window_GetElement(window,OxO58a4[7],true);var inp_class=Window_GetElement(window,OxO58a4[8],true);var inp_id=Window_GetElement(window,OxO58a4[9],true);var inp_tooltip=Window_GetElement(window,OxO58a4[10],true);var sel_noWrap=Window_GetElement(window,OxO58a4[11],true);var sel_CellScope=Window_GetElement(window,OxO58a4[12],true);SyncToView=function SyncToView_Td(){inp_bgColor[OxO58a4[13]]=element.getAttribute(OxO58a4[14])||element[OxO58a4[16]][OxO58a4[15]]||OxO58a4[17];inp_id[OxO58a4[13]]=element.getAttribute(OxO58a4[18])||OxO58a4[17];inp_bgColor[OxO58a4[16]][OxO58a4[15]]=inp_bgColor[OxO58a4[13]];inp_borderColor[OxO58a4[13]]=element.getAttribute(OxO58a4[19])||OxO58a4[17];inp_borderColor[OxO58a4[16]][OxO58a4[15]]=inp_borderColor[OxO58a4[13]];inp_borderColorLight[OxO58a4[13]]=element.getAttribute(OxO58a4[20])||OxO58a4[17];inp_borderColorLight[OxO58a4[16]][OxO58a4[15]]=inp_borderColorLight[OxO58a4[13]];inp_borderColorDark[OxO58a4[13]]=element.getAttribute(OxO58a4[21])||OxO58a4[17];inp_borderColorDark[OxO58a4[16]][OxO58a4[15]]=inp_borderColorDark[OxO58a4[13]];inp_class[OxO58a4[13]]=element[OxO58a4[22]];inp_width[OxO58a4[13]]=element.getAttribute(OxO58a4[23])||element[OxO58a4[16]][OxO58a4[23]]||OxO58a4[17];inp_height[OxO58a4[13]]=element.getAttribute(OxO58a4[24])||element[OxO58a4[16]][OxO58a4[24]]||OxO58a4[17];sel_align[OxO58a4[13]]=element.getAttribute(OxO58a4[25])||OxO58a4[17];sel_valign[OxO58a4[13]]=element.getAttribute(OxO58a4[26])||OxO58a4[17];inp_tooltip[OxO58a4[13]]=element.getAttribute(OxO58a4[27])||OxO58a4[17];sel_noWrap[OxO58a4[13]]=element.getAttribute(OxO58a4[28])||OxO58a4[17];sel_CellScope[OxO58a4[13]]=element.getAttribute(OxO58a4[29])||OxO58a4[17];} ;SyncTo=function SyncTo_Td(element){if(inp_bgColor[OxO58a4[13]]){if(element[OxO58a4[16]][OxO58a4[15]]){element[OxO58a4[16]][OxO58a4[15]]=inp_bgColor[OxO58a4[13]];} else {element[OxO58a4[14]]=inp_bgColor[OxO58a4[13]];} ;} else {element.removeAttribute(OxO58a4[14]);} ;element[OxO58a4[19]]=inp_borderColor[OxO58a4[13]];element[OxO58a4[20]]=inp_borderColorLight[OxO58a4[13]];element[OxO58a4[21]]=inp_borderColorDark[OxO58a4[13]];element[OxO58a4[22]]=inp_class[OxO58a4[13]];if(element[OxO58a4[16]][OxO58a4[23]]||element[OxO58a4[16]][OxO58a4[24]]){try{element[OxO58a4[16]][OxO58a4[23]]=inp_width[OxO58a4[13]];element[OxO58a4[16]][OxO58a4[24]]=inp_height[OxO58a4[13]];} catch(er){alert(CE_GetStr(OxO58a4[30]));} ;} else {try{element[OxO58a4[23]]=inp_width[OxO58a4[13]];element[OxO58a4[24]]=inp_height[OxO58a4[13]];} catch(er){alert(CE_GetStr(OxO58a4[30]));} ;} ;var Ox27a=/[^a-z\d]/i;if(Ox27a.test(inp_id.value)){alert(CE_GetStr(OxO58a4[31]));return ;} ;element[OxO58a4[25]]=sel_align[OxO58a4[13]];element[OxO58a4[18]]=inp_id[OxO58a4[13]];element[OxO58a4[26]]=sel_valign[OxO58a4[13]];element[OxO58a4[28]]=sel_noWrap[OxO58a4[13]];element[OxO58a4[27]]=inp_tooltip[OxO58a4[13]];element[OxO58a4[29]]=sel_CellScope[OxO58a4[13]];if(element[OxO58a4[18]]==OxO58a4[17]){element.removeAttribute(OxO58a4[18]);} ;if(element[OxO58a4[29]]==OxO58a4[17]){element.removeAttribute(OxO58a4[29]);} ;if(element[OxO58a4[28]]==OxO58a4[17]){element.removeAttribute(OxO58a4[28]);} ;if(element[OxO58a4[14]]==OxO58a4[17]){element.removeAttribute(OxO58a4[14]);} ;if(element[OxO58a4[19]]==OxO58a4[17]){element.removeAttribute(OxO58a4[19]);} ;if(element[OxO58a4[20]]==OxO58a4[17]){element.removeAttribute(OxO58a4[20]);} ;if(element[OxO58a4[7]]==OxO58a4[17]){element.removeAttribute(OxO58a4[7]);} ;if(element[OxO58a4[22]]==OxO58a4[17]){element.removeAttribute(OxO58a4[22]);} ;if(element[OxO58a4[22]]==OxO58a4[17]){element.removeAttribute(OxO58a4[32]);} ;if(element[OxO58a4[25]]==OxO58a4[17]){element.removeAttribute(OxO58a4[25]);} ;if(element[OxO58a4[26]]==OxO58a4[17]){element.removeAttribute(OxO58a4[33]);} ;if(element[OxO58a4[27]]==OxO58a4[17]){element.removeAttribute(OxO58a4[27]);} ;if(element[OxO58a4[23]]==OxO58a4[17]){element.removeAttribute(OxO58a4[23]);} ;if(element[OxO58a4[24]]==OxO58a4[17]){element.removeAttribute(OxO58a4[24]);} ;} ;inp_borderColor[OxO58a4[34]]=function inp_borderColor_onclick(){SelectColor(inp_borderColor);} ;inp_bgColor[OxO58a4[34]]=function inp_bgColor_onclick(){SelectColor(inp_bgColor);} ;inp_borderColorLight[OxO58a4[34]]=function inp_borderColorLight_onclick(){SelectColor(inp_borderColorLight);} ;inp_borderColorDark[OxO58a4[34]]=function inp_borderColorDark_onclick(){SelectColor(inp_borderColorDark);} ;