var OxO7fc9=["inp_src","inp_title","inp_target","sel_protocol","btnbrowse","editor","","href","value","title","target","onclick","length","options","://",":","others","selectedIndex"];var inp_src=Window_GetElement(window,OxO7fc9[0x0],true);var inp_title=Window_GetElement(window,OxO7fc9[0x1],true);var inp_target=Window_GetElement(window,OxO7fc9[0x2],true);var sel_protocol=Window_GetElement(window,OxO7fc9[0x3],true);var btnbrowse=Window_GetElement(window,OxO7fc9[0x4],true);var obj=Window_GetDialogArguments(window);var editor=obj[OxO7fc9[0x5]]; SyncToView() ; function SyncToView(){var src=obj[OxO7fc9[0x7]].replace(/$\s*/,OxO7fc9[0x6]); Update_sel_protocol(src) ; inp_src[OxO7fc9[0x8]]=src ;if(obj[OxO7fc9[0x9]]){ inp_title[OxO7fc9[0x8]]=obj[OxO7fc9[0x9]] ;} ;if(obj[OxO7fc9[0xa]]){ inp_target[OxO7fc9[0x8]]=obj[OxO7fc9[0xa]] ;} ;}  ; btnbrowse[OxO7fc9[0xb]]=function btnbrowse_onclick(){ function Ox2b3(Oxe6){if(Oxe6){ inp_src[OxO7fc9[0x8]]=Oxe6 ;} ;}  ; editor.SetNextDialogWindow(window) ;if(Browser_IsSafari()){ editor.ShowSelectFileDialog(Ox2b3,inp_src.value,inp_src) ;} else { editor.ShowSelectFileDialog(Ox2b3,inp_src.value) ;} ;}  ; function sel_protocol_change(){var src=inp_src[OxO7fc9[0x8]].replace(/$\s*/,OxO7fc9[0x6]);for(var i=0x0;i<sel_protocol[OxO7fc9[0xd]][OxO7fc9[0xc]];i++){var Oxe9=sel_protocol[OxO7fc9[0xd]][i][OxO7fc9[0x8]];if(src.substr(0x0,Oxe9.length).toLowerCase()==Oxe9){ src=src.substr(Oxe9[OxO7fc9[0xc]],src[OxO7fc9[0xc]]-Oxe9.length) ;break ;} ;} ;var Ox34e=src.indexOf(OxO7fc9[0xe]);if(Ox34e!=-0x1){ src=src.substr(Ox34e+0x3,src[OxO7fc9[0xc]]-0x3-Ox34e) ;} ;var Ox34e=src.indexOf(OxO7fc9[0xf]);if(Ox34e!=-0x1){ src=src.substr(Ox34e+0x1,src[OxO7fc9[0xc]]-0x1-Ox34e) ;} ;var Ox34f=sel_protocol[OxO7fc9[0x8]];if(Ox34f==OxO7fc9[0x10]){ Ox34f=OxO7fc9[0x6] ;} ; inp_src[OxO7fc9[0x8]]=Ox34f+src ;}  ; function Update_sel_protocol(src){var Ox351=false;for(var i=0x0;i<sel_protocol[OxO7fc9[0xd]][OxO7fc9[0xc]];i++){var Oxe9=sel_protocol[OxO7fc9[0xd]][i][OxO7fc9[0x8]];if(src.substr(0x0,Oxe9.length).toLowerCase()==Oxe9){if(sel_protocol[OxO7fc9[0x11]]!=i){ sel_protocol[OxO7fc9[0x11]]=i ;} ; Ox351=true ;break ;} ;} ;if(!Ox351){ sel_protocol[OxO7fc9[0x11]]=sel_protocol[OxO7fc9[0xd]][OxO7fc9[0xc]]-0x1 ;} ;}  ; function insert_link(){var arr= new Array(); arr[0x0]=inp_src[OxO7fc9[0x8]] ;if(inp_target[OxO7fc9[0x8]]){ arr[0x1]=inp_target[OxO7fc9[0x8]] ;} ;if(inp_title[OxO7fc9[0x8]]){ arr[0x2]=inp_title[OxO7fc9[0x8]] ;} ; Window_SetDialogReturnValue(window,arr) ; Window_CloseDialog(window) ;}  ;