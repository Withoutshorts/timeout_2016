var OxOcdfc=["INPUT","TEXTAREA","DIV","SPAN","","value","innerHTML","BODY","contentWindow","body","document","IFRAME","tagName","length","type","text","id","isContentEditable","showModalDialog","?","?Modal=true","\x26Modal=true","dialogHeight:340px; dialogWidth:395px; edge:Raised; center:Yes; help:No; resizable:Yes; status:No; scroll:No","newWindow","height=300,width=400,scrollbars=no,resizable=no,toolbars=no,status=no,menubar=no,location=no","ElementIndex","dialogArguments","window","opener","SpellMode","start","suggest","end","SpellingForm","checkElements","innerText","StatusText","Spell Checking Text ...","CurrentText","WordIndex","Spell Check Complete","selectedIndex","ReplacementWord","form","options"];var showCompleteAlert=true;var tagGroup= new Array(OxOcdfc[0],OxOcdfc[1],OxOcdfc[2],OxOcdfc[3]);var checkElements= new Array();function getText(Oxed){var Oxee=document.getElementById(checkElements[Oxed]);var Oxef=OxOcdfc[4];switch(Oxee[OxOcdfc[12]]){case OxOcdfc[0]:;case OxOcdfc[1]:Oxef=Oxee[OxOcdfc[5]];break ;;case OxOcdfc[2]:;case OxOcdfc[3]:;case OxOcdfc[7]:Oxef=Oxee[OxOcdfc[6]];break ;;case OxOcdfc[11]:var Oxf0=document.getElementById(Oxee.id);if(Oxf0[OxOcdfc[8]]){Oxef=Oxf0[OxOcdfc[8]][OxOcdfc[10]][OxOcdfc[9]][OxOcdfc[6]];} else {Oxef=Oxf0[OxOcdfc[10]][OxOcdfc[9]][OxOcdfc[6]];} ;;} ;return Oxef;} ;function setText(Oxed,Oxf2){var Oxee=document.getElementById(checkElements[Oxed]);switch(Oxee[OxOcdfc[12]]){case OxOcdfc[0]:;case OxOcdfc[1]:Oxee[OxOcdfc[5]]=Oxf2;break ;;case OxOcdfc[2]:;case OxOcdfc[3]:Oxee[OxOcdfc[6]]=Oxf2;break ;;case OxOcdfc[11]:var Oxf0=document.getElementById(Oxee.id);if(Oxf0[OxOcdfc[8]]){Oxf0[OxOcdfc[8]][OxOcdfc[10]][OxOcdfc[9]][OxOcdfc[6]]=Oxf2;} else {Oxf0[OxOcdfc[10]][OxOcdfc[9]][OxOcdfc[6]]=Oxf2;} ;break ;;} ;} ;function checkSpelling(){checkElements= new Array();for(var i=0;i<tagGroup[OxOcdfc[13]];i++){var Oxf4=tagGroup[i];var Oxf5=document.getElementsByTagName(Oxf4);for(var Oxf6=0;Oxf6<Oxf5[OxOcdfc[13]];Oxf6++){if((Oxf4==OxOcdfc[0]&&Oxf5[Oxf6][OxOcdfc[14]]==OxOcdfc[15])||Oxf4==OxOcdfc[1]){checkElements[checkElements[OxOcdfc[13]]]=Oxf5[Oxf6][OxOcdfc[16]];} else {if((Oxf4==OxOcdfc[2]||Oxf4==OxOcdfc[3])&&Oxf5[Oxf6][OxOcdfc[17]]){checkElements[checkElements[OxOcdfc[13]]]=Oxf5[Oxf6][OxOcdfc[16]];} ;} ;} ;} ;openSpellChecker();} ;function checkSpellingById(Oxaf,Oxf8){checkElements= new Array();checkElements[checkElements[OxOcdfc[13]]]=Oxaf;openSpellChecker(Oxf8);} ;function checkElementSpelling(Oxee){checkElements= new Array();checkElements[checkElements[OxOcdfc[13]]]=Oxee[OxOcdfc[16]];openSpellChecker();} ;function openSpellChecker(Oxf8){if(window[OxOcdfc[18]]){var Oxfb;if(Oxf8.indexOf(OxOcdfc[19])==-1){Oxfb=OxOcdfc[20];} else {Oxfb=OxOcdfc[21];} ;var Oxfc=window.showModalDialog(Oxf8+Oxfb,window,OxOcdfc[22]);} else {var Oxfd=window.open(Oxf8,OxOcdfc[23],OxOcdfc[24]);} ;} ;var iElementIndex=-1;var parentWindow;function initialize(){iElementIndex=parseInt(document.getElementById(OxOcdfc[25]).value);if(parent[OxOcdfc[27]][OxOcdfc[26]]){parentWindow=parent[OxOcdfc[27]][OxOcdfc[26]];} else {if(top[OxOcdfc[28]]){parentWindow=top[OxOcdfc[28]];} ;} ;var Ox101=document.getElementById(OxOcdfc[29])[OxOcdfc[5]];switch(Ox101){case OxOcdfc[30]:break ;;case OxOcdfc[31]:updateText();break ;;case OxOcdfc[32]:updateText();;default:if(loadText()){document.getElementById(OxOcdfc[33]).submit();} else {endCheck();} ;break ;;} ;} ;function loadText(){if(!parentWindow[OxOcdfc[10]]){return false;} ;for(++iElementIndex;iElementIndex<parentWindow[OxOcdfc[34]][OxOcdfc[13]];iElementIndex++){var Ox103=parentWindow.getText(iElementIndex);if(Ox103[OxOcdfc[13]]>0){updateSettings(Ox103,0,iElementIndex,OxOcdfc[30]);document.getElementById(OxOcdfc[36])[OxOcdfc[35]]=OxOcdfc[37];return true;} ;} ;return false;} ;function updateSettings(Ox105,Ox106,Ox107,Ox108){document.getElementById(OxOcdfc[38])[OxOcdfc[5]]=Ox105;document.getElementById(OxOcdfc[39])[OxOcdfc[5]]=Ox106;document.getElementById(OxOcdfc[25])[OxOcdfc[5]]=Ox107;document.getElementById(OxOcdfc[29])[OxOcdfc[5]]=Ox108;} ;function updateText(){if(!parentWindow[OxOcdfc[10]]){return false;} ;var Ox103=document.getElementById(OxOcdfc[38])[OxOcdfc[5]];parentWindow.setText(iElementIndex,Ox103);} ;function endCheck(){if(showCompleteAlert){alert(OxOcdfc[40]);} ;closeWindow();} ;function closeWindow(){window.close();top.close();} ;function changeWord(Oxee){var Ox10d=Oxee[OxOcdfc[41]];Oxee[OxOcdfc[43]][OxOcdfc[42]][OxOcdfc[5]]=Oxee[OxOcdfc[44]][Ox10d][OxOcdfc[5]];} ;