var OxO4f1c=["inp_src","inp_title","inp_target","sel_protocol","btnbrowse","editor","","href","value","title","target","onclick","length","options","://",":","others","selectedIndex"];var inp_src=Window_GetElement(window,OxO4f1c[0],true);var inp_title=Window_GetElement(window,OxO4f1c[1],true);var inp_target=Window_GetElement(window,OxO4f1c[2],true);var sel_protocol=Window_GetElement(window,OxO4f1c[3],true);var btnbrowse=Window_GetElement(window,OxO4f1c[4],true);var obj=Window_GetDialogArguments(window);var editor=obj[OxO4f1c[5]];SyncToView();function SyncToView(){var src=obj[OxO4f1c[7]].replace(/$\s*/,OxO4f1c[6]);Update_sel_protocol(src);inp_src[OxO4f1c[8]]=src;if(obj[OxO4f1c[9]]){inp_title[OxO4f1c[8]]=obj[OxO4f1c[9]];} ;if(obj[OxO4f1c[10]]){inp_target[OxO4f1c[8]]=obj[OxO4f1c[10]];} ;} ;btnbrowse[OxO4f1c[11]]=function btnbrowse_onclick(){function Ox260(Ox6){if(Ox6){inp_src[OxO4f1c[8]]=Ox6;} ;} ;editor.SetNextDialogWindow(window);if(Browser_IsSafari()){editor.ShowSelectFileDialog(Ox260,inp_src.value,inp_src);} else {editor.ShowSelectFileDialog(Ox260,inp_src.value);} ;} ;function sel_protocol_change(){var src=inp_src[OxO4f1c[8]].replace(/$\s*/,OxO4f1c[6]);for(var i=0;i<sel_protocol[OxO4f1c[13]][OxO4f1c[12]];i++){var Ox9=sel_protocol[OxO4f1c[13]][i][OxO4f1c[8]];if(src.substr(0,Ox9.length).toLowerCase()==Ox9){src=src.substr(Ox9.length,src[OxO4f1c[12]]-Ox9[OxO4f1c[12]]);break ;} ;} ;var Ox31f=src.indexOf(OxO4f1c[14]);if(Ox31f!=-1){src=src.substr(Ox31f+3,src[OxO4f1c[12]]-3-Ox31f);} ;var Ox31f=src.indexOf(OxO4f1c[15]);if(Ox31f!=-1){src=src.substr(Ox31f+1,src[OxO4f1c[12]]-1-Ox31f);} ;var Ox320=sel_protocol[OxO4f1c[8]];if(Ox320==OxO4f1c[16]){Ox320=OxO4f1c[6];} ;inp_src[OxO4f1c[8]]=Ox320+src;} ;function Update_sel_protocol(src){var Ox322=false;for(var i=0;i<sel_protocol[OxO4f1c[13]][OxO4f1c[12]];i++){var Ox9=sel_protocol[OxO4f1c[13]][i][OxO4f1c[8]];if(src.substr(0,Ox9.length).toLowerCase()==Ox9){if(sel_protocol[OxO4f1c[17]]!=i){sel_protocol[OxO4f1c[17]]=i;} ;Ox322=true;break ;} ;} ;if(!Ox322){sel_protocol[OxO4f1c[17]]=sel_protocol[OxO4f1c[13]][OxO4f1c[12]]-1;} ;} ;function insert_link(){var arr= new Array();arr[0]=inp_src[OxO4f1c[8]];if(inp_target[OxO4f1c[8]]){arr[1]=inp_target[OxO4f1c[8]];} ;if(inp_title[OxO4f1c[8]]){arr[2]=inp_title[OxO4f1c[8]];} ;Window_SetDialogReturnValue(window,arr);Window_CloseDialog(window);} ;