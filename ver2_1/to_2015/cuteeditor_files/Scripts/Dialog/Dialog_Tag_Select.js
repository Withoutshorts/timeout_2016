var OxO5230=["inp_name","inp_access","inp_id","inp_index","inp_size","inp_Multiple","inp_Disabled","inp_item_text","inp_item_value","btnInsertItem","btnUpdateItem","btnDeleteItem","btnMoveUpItem","btnMoveDownItem","list_options","list_options2","inp_item_forecolor","inp_item_forecolor_Preview","inp_item_backcolor_Preview","text","value","color","style","backgroundColor","selectedIndex","options","Please select an item first","length","ondblclick","onclick","OPTION","document","id","cssText","Name","name","size","checked","disabled","multiple","tabIndex","","accessKey"];var inp_name=Window_GetElement(window,OxO5230[0],true);var inp_access=Window_GetElement(window,OxO5230[1],true);var inp_id=Window_GetElement(window,OxO5230[2],true);var inp_index=Window_GetElement(window,OxO5230[3],true);var inp_size=Window_GetElement(window,OxO5230[4],true);var inp_Multiple=Window_GetElement(window,OxO5230[5],true);var inp_Disabled=Window_GetElement(window,OxO5230[6],true);var inp_item_text=Window_GetElement(window,OxO5230[7],true);var inp_item_value=Window_GetElement(window,OxO5230[8],true);var btnInsertItem=Window_GetElement(window,OxO5230[9],true);var btnUpdateItem=Window_GetElement(window,OxO5230[10],true);var btnDeleteItem=Window_GetElement(window,OxO5230[11],true);var btnMoveUpItem=Window_GetElement(window,OxO5230[12],true);var btnMoveDownItem=Window_GetElement(window,OxO5230[13],true);var list_options=Window_GetElement(window,OxO5230[14],true);var list_options2=Window_GetElement(window,OxO5230[15],true);var inp_item_forecolor=Window_GetElement(window,OxO5230[16],true);var inp_item_forecolor=Window_GetElement(window,OxO5230[16],true);var inp_item_forecolor_Preview=Window_GetElement(window,OxO5230[17],true);var inp_item_backcolor_Preview=Window_GetElement(window,OxO5230[18],true);function SetOption(Ox5){Ox5[OxO5230[19]]=inp_item_text[OxO5230[20]];Ox5[OxO5230[20]]=inp_item_value[OxO5230[20]];Ox5[OxO5230[22]][OxO5230[21]]=inp_item_forecolor[OxO5230[20]];Ox5[OxO5230[22]][OxO5230[23]]=inp_item_backcolor[OxO5230[20]];} ;function SetOption2(Ox5){Ox5[OxO5230[19]]=inp_item_value[OxO5230[20]];Ox5[OxO5230[20]]=inp_item_text[OxO5230[20]];Ox5[OxO5230[22]][OxO5230[21]]=inp_item_forecolor[OxO5230[20]];Ox5[OxO5230[22]][OxO5230[23]]=inp_item_backcolor[OxO5230[20]];} ;function Select(Ox5){var Ox483=Ox5[OxO5230[24]];list_options[OxO5230[24]]=Ox483;list_options2[OxO5230[24]]=Ox483;inp_item_text[OxO5230[20]]=list_options2[OxO5230[20]];inp_item_value[OxO5230[20]]=list_options[OxO5230[20]];} ;function Insert(){var Ox5= new Option();SetOption(Ox5);list_options[OxO5230[25]].add(Ox5);var Ox485= new Option();SetOption2(Ox485);list_options2[OxO5230[25]].add(Ox485);FireUIChanged();} ;function Update(){if(list_options[OxO5230[24]]==-1){alert(OxO5230[26]);return ;} ;var Ox5=list_options.options(list_options.selectedIndex);SetOption(Ox5);var Ox485=list_options2.options(list_options2.selectedIndex);SetOption2(Ox485);FireUIChanged();} ;function Move(Ox9){var Ox483=list_options[OxO5230[24]];if(Ox483<0){return ;} ;var Ox488=Ox483-Ox9;if(Ox488<0){Ox488=0;} ;if(Ox488>(list_options[OxO5230[25]][OxO5230[27]]-1)){Ox488=list_options[OxO5230[25]][OxO5230[27]]-1;} ;if(Ox483==Ox488){return ;} ;var Ox5=list_options.options(list_options.selectedIndex);var Ox6a=list_options2[OxO5230[20]];var Ox8a=list_options[OxO5230[20]];Delete();inp_item_text[OxO5230[20]]=Ox6a;inp_item_value[OxO5230[20]]=Ox8a;var Ox5= new Option();SetOption(Ox5);list_options[OxO5230[25]].add(Ox5,Ox488);var Ox485= new Option();SetOption2(Ox485);list_options2[OxO5230[25]].add(Ox485,Ox488);list_options[OxO5230[24]]=Ox488;list_options2[OxO5230[24]]=Ox488;FireUIChanged();} ;function Delete(){if(list_options[OxO5230[24]]==-1||list_options[OxO5230[24]]==-1){alert(OxO5230[26]);return ;} ;var Ox48a=list_options[OxO5230[24]];var Ox5=list_options.options(list_options.selectedIndex);Ox5.removeNode(true);Ox5=list_options2.options(list_options2.selectedIndex);Ox5.removeNode(true);if(list_options[OxO5230[25]][OxO5230[27]]>Ox48a){list_options[OxO5230[24]]=Ox48a;} else {if(list_options[OxO5230[25]][OxO5230[27]]){list_options[OxO5230[24]]=list_options[OxO5230[25]][OxO5230[27]]-1;} ;} ;list_options.ondblclick();if(list_options2[OxO5230[25]][OxO5230[27]]>Ox48a){list_options2[OxO5230[24]]=Ox48a;} else {if(list_options2[OxO5230[25]][OxO5230[27]]){list_options2[OxO5230[24]]=list_options2[OxO5230[25]][OxO5230[27]]-1;} ;} ;FireUIChanged();} ;list_options[OxO5230[28]]=function list_options_ondblclick(){if(list_options[OxO5230[24]]==-1){return ;} ;var Ox5=list_options.options(list_options.selectedIndex);inp_item_text[OxO5230[20]]=Ox5[OxO5230[19]];inp_item_value[OxO5230[20]]=Ox5[OxO5230[20]];inp_item_forecolor[OxO5230[20]]=inp_item_forecolor[OxO5230[22]][OxO5230[23]]=inp_item_forecolor_Preview[OxO5230[22]][OxO5230[23]]=Ox5[OxO5230[22]][OxO5230[21]];inp_item_backcolor[OxO5230[20]]=inp_item_backcolor[OxO5230[22]][OxO5230[23]]=inp_item_backcolor_Preview[OxO5230[22]][OxO5230[23]]=Ox5[OxO5230[22]][OxO5230[23]];} ;inp_item_forecolor[OxO5230[29]]=inp_item_forecolor_Preview[OxO5230[29]]=function inp_item_forecolor_onclick(){SelectColor(inp_item_forecolor,inp_item_forecolor_Preview);} ;inp_item_backcolor[OxO5230[29]]=inp_item_backcolor_Preview[OxO5230[29]]=function inp_item_backcolor_onclick(){SelectColor(inp_item_backcolor,inp_item_backcolor_Preview);} ;function CopySelect(Ox48f,Ox490){Ox490[OxO5230[25]][OxO5230[27]]=0;for(var i=0;i<Ox48f[OxO5230[25]][OxO5230[27]];i++){var Ox491=Ox48f[OxO5230[25]][i];var Ox485;if(Browser_IsWinIE()){Ox485=Ox490[OxO5230[31]].createElement(OxO5230[30]);} else {Ox485=document.createElement(OxO5230[30]);} ;if(Ox490[OxO5230[32]]!=OxO5230[15]){Ox485[OxO5230[20]]=Ox491[OxO5230[20]];Ox485[OxO5230[19]]=Ox491[OxO5230[19]];} else {Ox485[OxO5230[20]]=Ox491[OxO5230[19]];Ox485[OxO5230[19]]=Ox491[OxO5230[20]];} ;Ox485[OxO5230[22]][OxO5230[33]]=Ox491[OxO5230[22]][OxO5230[33]];Ox490[OxO5230[25]].add(Ox485);} ;Ox490[OxO5230[24]]=Ox48f[OxO5230[24]];} ;UpdateState=function UpdateState_Select(){} ;SyncToView=function SyncToView_Select(){if(element[OxO5230[34]]){inp_name[OxO5230[20]]=element[OxO5230[34]];} ;if(element[OxO5230[35]]){inp_name[OxO5230[20]]=element[OxO5230[35]];} ;inp_id[OxO5230[20]]=element[OxO5230[32]];inp_size[OxO5230[20]]=element[OxO5230[36]];CopySelect(element,list_options);CopySelect(element,list_options2);inp_Disabled[OxO5230[37]]=element[OxO5230[38]];inp_Multiple[OxO5230[37]]=element[OxO5230[39]];if(element[OxO5230[40]]==0){inp_index[OxO5230[20]]=OxO5230[41];} else {inp_index[OxO5230[20]]=element[OxO5230[40]];} ;if(element[OxO5230[42]]){inp_access[OxO5230[20]]=element[OxO5230[42]];} ;} ;SyncTo=function SyncTo_Select(element){element[OxO5230[35]]=inp_name[OxO5230[20]];if(element[OxO5230[34]]){element[OxO5230[34]]=inp_name[OxO5230[20]];} else {if(element[OxO5230[35]]){element.removeAttribute(OxO5230[35],0);element[OxO5230[34]]=inp_name[OxO5230[20]];} else {element[OxO5230[34]]=inp_name[OxO5230[20]];} ;} ;element[OxO5230[32]]=inp_id[OxO5230[20]];element[OxO5230[36]]=inp_size[OxO5230[20]];element[OxO5230[38]]=inp_Disabled[OxO5230[37]];element[OxO5230[39]]=inp_Multiple[OxO5230[37]];element[OxO5230[42]]=inp_access[OxO5230[20]];element[OxO5230[40]]=inp_index[OxO5230[20]];if(element[OxO5230[40]]==OxO5230[41]){element.removeAttribute(OxO5230[40]);} ;if(element[OxO5230[42]]==OxO5230[41]){element.removeAttribute(OxO5230[42]);} ;CopySelect(list_options,element);} ;