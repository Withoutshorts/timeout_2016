var OxOa5a2=["inp_name","inp_cols","inp_rows","inp_value","sel_Wrap","inp_id","inp_access","inp_index","inp_Disabled","inp_Readonly","Name","value","name","id","cols","","rows","checked","disabled","readOnly","wrap","tabIndex","accessKey","textContent"];var inp_name=Window_GetElement(window,OxOa5a2[0],true);var inp_cols=Window_GetElement(window,OxOa5a2[1],true);var inp_rows=Window_GetElement(window,OxOa5a2[2],true);var inp_value=Window_GetElement(window,OxOa5a2[3],true);var sel_Wrap=Window_GetElement(window,OxOa5a2[4],true);var inp_id=Window_GetElement(window,OxOa5a2[5],true);var inp_access=Window_GetElement(window,OxOa5a2[6],true);var inp_index=Window_GetElement(window,OxOa5a2[7],true);var inp_Disabled=Window_GetElement(window,OxOa5a2[8],true);var inp_Readonly=Window_GetElement(window,OxOa5a2[9],true);UpdateState=function UpdateState_Textarea(){} ;SyncToView=function SyncToView_Textarea(){if(element[OxOa5a2[10]]){inp_name[OxOa5a2[11]]=element[OxOa5a2[10]];} ;if(element[OxOa5a2[12]]){inp_name[OxOa5a2[11]]=element[OxOa5a2[12]];} ;inp_id[OxOa5a2[11]]=element[OxOa5a2[13]];inp_value[OxOa5a2[11]]=element[OxOa5a2[11]];if(element[OxOa5a2[14]]){if(element[OxOa5a2[14]]==20){inp_cols[OxOa5a2[11]]=OxOa5a2[15];} else {inp_cols[OxOa5a2[11]]=element[OxOa5a2[14]];} ;} ;if(element[OxOa5a2[16]]){if(element[OxOa5a2[16]]==2){inp_rows[OxOa5a2[11]]=OxOa5a2[15];} else {inp_rows[OxOa5a2[11]]=element[OxOa5a2[16]];} ;} ;inp_Disabled[OxOa5a2[17]]=element[OxOa5a2[18]];inp_Readonly[OxOa5a2[17]]=element[OxOa5a2[19]];sel_Wrap[OxOa5a2[11]]=element[OxOa5a2[20]];if(element[OxOa5a2[21]]==0){inp_index[OxOa5a2[11]]=OxOa5a2[15];} else {inp_index[OxOa5a2[11]]=element[OxOa5a2[21]];} ;if(element[OxOa5a2[22]]){inp_access[OxOa5a2[11]]=element[OxOa5a2[22]];} ;} ;SyncTo=function SyncTo_Textarea(element){element[OxOa5a2[12]]=inp_name[OxOa5a2[11]];if(element[OxOa5a2[10]]){element[OxOa5a2[10]]=inp_name[OxOa5a2[11]];} else {if(element[OxOa5a2[12]]){element.removeAttribute(OxOa5a2[12],0);element[OxOa5a2[10]]=inp_name[OxOa5a2[11]];} else {element[OxOa5a2[10]]=inp_name[OxOa5a2[11]];} ;} ;element[OxOa5a2[13]]=inp_id[OxOa5a2[11]];element[OxOa5a2[11]]=inp_value[OxOa5a2[11]];if(!Browser_IsWinIE()){try{element[OxOa5a2[23]]=inp_value[OxOa5a2[11]];} catch(x){} ;} ;element[OxOa5a2[21]]=inp_index[OxOa5a2[11]];element[OxOa5a2[18]]=inp_Disabled[OxOa5a2[17]];element[OxOa5a2[19]]=inp_Readonly[OxOa5a2[17]];element[OxOa5a2[22]]=inp_access[OxOa5a2[11]];if(inp_cols[OxOa5a2[11]]==OxOa5a2[15]){element[OxOa5a2[14]]=20;} else {element[OxOa5a2[14]]=inp_cols[OxOa5a2[11]];} ;if(inp_rows[OxOa5a2[11]]==OxOa5a2[15]){element[OxOa5a2[16]]=2;} else {element[OxOa5a2[16]]=inp_rows[OxOa5a2[11]];} ;try{element[OxOa5a2[20]]=sel_Wrap[OxOa5a2[11]];} catch(e){element.removeAttribute(OxOa5a2[20]);} ;element[OxOa5a2[21]]=inp_index[OxOa5a2[11]];if(element[OxOa5a2[21]]==OxOa5a2[15]){element.removeAttribute(OxOa5a2[21]);} ;if(element[OxOa5a2[22]]==OxOa5a2[15]){element.removeAttribute(OxOa5a2[22]);} ;} ;