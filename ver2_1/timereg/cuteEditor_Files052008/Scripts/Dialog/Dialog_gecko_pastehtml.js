var OxOff97=["onload","contentWindow","idSource","innerHTML","body","document","","designMode","on","contentEditable","fontFamily","style","Tahoma","fontSize","11px","color","black","background","white","length","\x3C$1$3","\x26nbsp;","\x22","\x27","$1","\x26amp;","\x26lt;","\x26gt;","\x26#39;","\x26quot;"];var editor=Window_GetDialogArguments(window); function cancel(){ Window_CloseDialog(window) ;}  ; window[OxOff97[0x0]]=function (){var iframe=document.getElementById(OxOff97[0x2])[OxOff97[0x1]]; iframe[OxOff97[0x5]][OxOff97[0x4]][OxOff97[0x3]]=OxOff97[0x6] ; iframe[OxOff97[0x5]][OxOff97[0x7]]=OxOff97[0x8] ; iframe[OxOff97[0x5]][OxOff97[0x4]][OxOff97[0x9]]=true ; iframe[OxOff97[0x5]][OxOff97[0x4]][OxOff97[0xb]][OxOff97[0xa]]=OxOff97[0xc] ; iframe[OxOff97[0x5]][OxOff97[0x4]][OxOff97[0xb]][OxOff97[0xd]]=OxOff97[0xe] ; iframe[OxOff97[0x5]][OxOff97[0x4]][OxOff97[0xb]][OxOff97[0xf]]=OxOff97[0x10] ; iframe[OxOff97[0x5]][OxOff97[0x4]][OxOff97[0xb]][OxOff97[0x11]]=OxOff97[0x12] ; iframe.focus() ;}  ; function insertContent(){var iframe=document.getElementById(OxOff97[0x2])[OxOff97[0x1]];var Ox1ec=iframe[OxOff97[0x5]][OxOff97[0x4]][OxOff97[0x3]];if(Ox1ec&&Ox1ec[OxOff97[0x13]]>0x0){ Ox1ec=_CleanCode(Ox1ec) ;if(Ox1ec.match(/<*>/g)){ Ox1ec=String_HtmlEncode(Ox1ec) ;} ; editor.PasteHTML(Ox1ec) ; Window_CloseDialog(window) ;} ;}  ; function _CleanCode(Ox83){ Ox83=Ox83.replace(/<\\?\??xml[^>]>/gi,OxOff97[0x6]) ; Ox83=Ox83.replace(/<([\w]+) class=([^ |>]*)([^>]*)/gi,OxOff97[0x14]) ; Ox83=Ox83.replace(/<(\w[^>]*) lang=([^ |>]*)([^>]*)/gi,OxOff97[0x14]) ; Ox83=Ox83.replace(/\s*mso-[^:]+:[^;"]+;?/gi,OxOff97[0x6]) ; Ox83=Ox83.replace(/<o:p>\s*<\/o:p>/g,OxOff97[0x6]) ; Ox83=Ox83.replace(/<o:p>.*?<\/o:p>/g,OxOff97[0x15]) ; Ox83=Ox83.replace(/<\/?\w+:[^>]*>/gi,OxOff97[0x6]) ; Ox83=Ox83.replace(/<\!--.*-->/g,OxOff97[0x6]) ; Ox83=Ox83.replace(/[\\]/gi,OxOff97[0x16]) ; Ox83=Ox83.replace(/[\\]/gi,OxOff97[0x17]) ; Ox83=Ox83.replace(/<\\?\?xml[^>]*>/gi,OxOff97[0x6]) ; Ox83=Ox83.replace(/<(\w+)[^>]*\sstyle="[^"]*DISPLAY\s?:\s?none(.*?)<\/\1>/ig,OxOff97[0x6]) ; Ox83=Ox83.replace(/<span\s*[^>]*>\s*&nbsp;\s*<\/span>/gi,OxOff97[0x15]) ; Ox83=Ox83.replace(/<span\s*[^>]*><\/span>/gi,OxOff97[0x6]) ; Ox83=Ox83.replace(/\s*style="\s*"/gi,OxOff97[0x6]) ; Ox83=Ox83.replace(/<([^\s>]+)[^>]*>\s*<\/\1>/g,OxOff97[0x6]) ; Ox83=Ox83.replace(/<([^\s>]+)[^>]*>\s*<\/\1>/g,OxOff97[0x6]) ; Ox83=Ox83.replace(/<([^\s>]+)[^>]*>\s*<\/\1>/g,OxOff97[0x6]) ;while(Ox83.match(/<span\s*>(.*?)<\/span>/gi)){ Ox83=Ox83.replace(/<span\s*>(.*?)<\/span>/gi,OxOff97[0x18]) ;} ;while(Ox83.match(/<font\s*>(.*?)<\/font>/gi)){ Ox83=Ox83.replace(/<font\s*>(.*?)<\/font>/gi,OxOff97[0x18]) ;} ; Ox83=Ox83.replace(/<a name="?OLE_LINK\d+"?>((.|[\r\n])*?)<\/a>/gi,OxOff97[0x18]) ; Ox83=Ox83.replace(/<a name="?_Hlt\d+"?>((.|[\r\n])*?)<\/a>/gi,OxOff97[0x18]) ; Ox83=Ox83.replace(/<a name="?_Toc\d+"?>((.|[\r\n])*?)<\/a>/gi,OxOff97[0x18]) ; Ox83=Ox83.replace(/<p([^>])*>(&nbsp;)*\s*<\/p>/gi,OxOff97[0x6]) ; Ox83=Ox83.replace(/<p([^>])*>(&nbsp;)<\/p>/gi,OxOff97[0x6]) ;return Ox83;}  ; function String_HtmlEncode(Ox1d4){if(Ox1d4==null){return OxOff97[0x6];} ; Ox1d4=Ox1d4.replace(/&/g,OxOff97[0x19]) ; Ox1d4=Ox1d4.replace(/</g,OxOff97[0x1a]) ; Ox1d4=Ox1d4.replace(/>/g,OxOff97[0x1b]) ; Ox1d4=Ox1d4.replace(/'/g,OxOff97[0x1c]) ; Ox1d4=Ox1d4.replace(/\x22/g,OxOff97[0x1d]) ;return Ox1d4;}  ;