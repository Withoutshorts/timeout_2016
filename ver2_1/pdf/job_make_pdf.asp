<%'Response.buffer = true %>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/regular/header_hvd_inc.asp"-->
<%


lto = request("lto")

if len(request("id")) <> 0 then
id = request("id")
vis_indledning = request("FM_vis_indledning_pdf")
vis_footer = request("FM_vis_afslutning_pdf")
else
id = 0
end if
'kid = request("kid")



if len(trim(vis_indledning)) <> 0 then

strSQLpdfval = "INSERT INTO pdf_values (pdf_txt, pdf_footer) VALUES ('"& vis_indledning &"', '"& vis_footer &"')"
oConn.execute(strSQLpdfval)

strSQLsel = "SELECT id FROM pdf_values WHERE id <> 0 ORDER BY id DESC LIMIT 0, 1"
oRec.open strSQLsel, oConn, 3
if not oRec.EOF then

pdfvalid = oRec("id")

end if
oRec.close

'Response.Cookies("job")("pdfval") = vis_indledning
end if

'if request.Cookies("job")("pdfval") <> "" then
'Response.Write "cval:" & request.Cookies("job")("pdfval")
'end if

'Response.end


'Response.Write "vis_indledning " & vis_indledning
'Response.end


'dokid = request("FM_dokument")

'gemafs = request("FM_gem_afs")
'visJbesk = request("FM_vis_jbesk")
'visPris = request("FM_vis_pris")
'visjobstamdata = request("FM_vis_per")
'afslutningVal = request("FM_vis_afslutning")
'visfiler = request("FM_vis_filer")
'visMile = request("FM_vis_milepale")
'visAkt = request("FM_vis_akt")
'jobtilbudsnr = request("jobtilbudsnr")
'visDisseAkt = request("FM_vis_aktid")
'visAktPer = request("FM_vis_aktper")


if id <> 0 then

filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)


    
        'Response.redirect "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/job_print.asp?nosession=9999&func=print&lto="&lto&"&id="&id&"&"_
        '&"hideudskriftfunc=1&FM_gem_afs="&gemafs&"&FM_vis_jbesk="&visJbesk&"&"_
		'&"FM_vis_pris="&visPris&"&FM_vis_per="&visjobstamdata&"&FM_vis_afslutning="&afslutningVal&"&"_
		'&"FM_vis_filer="&visfiler&"&FM_vis_milepale="&visMile&"&FM_vis_akt="&visAkt&"&FM_vis_aktper="&visAktPer&"&jobtilbudsnr="&jobtilbudsnr&"&FM_vis_aktid="&visDisseAkt
        




        Set Pdf = Server.CreateObject("Persits.Pdf")
        Set Doc = Pdf.CreateDocument
        'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/erp_fak_godkendt_2007.asp?nosession=9999&print=pdf&jobid="&jobid&"&aftid="&aftid&"&id="&id&"&key="&session("lto")
        
        'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/job_print.asp?nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&"_
        '&"hideudskriftfunc=2&FM_gem_afs="&gemafs&"&FM_vis_jbesk="&visJbesk&"&"_
		'&"FM_vis_pris="&visPris&"&FM_vis_per="&visjobstamdata&"&FM_vis_afslutning="&afslutningVal&"&"_
		'&"FM_vis_filer="&visfiler&"&FM_vis_milepale="&visMile&"&FM_vis_akt="&visAkt&"&FM_vis_aktper="&visAktPer&"&jobtilbudsnr="&jobtilbudsnr&"&FM_vis_aktid="&visDisseAkt
        
        
        'Response.Write "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&FM_vis_indledning=<![CDATA["&vis_indledning&"]]"
        'Response.end
        
         Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin = 5, RightMargin = 5, TopMargin = 10, BottomMargin = 0"
         
         
         'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg/job_make_pdf.asp?lto="&lto
         'FM_vis_indledning=<![CDATA["&vis_indledning&"]]&
       
        
        
        'Doc.ImportFromUrl "http://localhost/timeout_xp/timereg/job_print.asp?func=print&nosession=9999&lto="&lto&"&id="&=id&"&_
        'hideudskriftfunc=1&FM_gem_afs="&gemafs&"&FM_vis_jbesk="&visJbesk&"&_
		'FM_vis_pris="&visPris&"&FM_vis_per="&visjobstamdata&"&FM_vis_afslutning="&afslutningVal&"&_
		'FM_vis_filer="&visfiler&"&FM_vis_milepale="&visMile&"&FM_vis_akt="&visAkt&"&FM_vis_aktper="&visAktPer&"&jobtilbudsnr="&jobtilbudsnr&"&FM_vis_aktid="&visDisseAkt
        
        
        'Filename = Doc.Save(Server.MapPath("../inc/upload/"&lto&"/pdf_"&lto&"_"&filnavnDato&"_"&filnavnKlok&".pdf"), True) 
        Filename = Doc.Save(Server.MapPath("../inc/upload/outz/pdf_"&lto&"_"&filnavnDato&"_"&filnavnKlok&".pdf"), True) 





Response.redirect "../inc/upload/outz/pdf_"&lto&"_"&filnavnDato&"_"&filnavnKlok&".pdf"


end if

%>


<!--
<br /><br /><br /><br />
<table cellspacing=0 cellpadding=0 border=0>
<tr><td style="padding:20px 20px 20px 20px;">
<h4>TimeOut PDF</h4>

Din PDF fil er produceret og gemt på serveren. <br />
Du kan hente din DPF faktura <a href="../inc/upload/"<%=lto%>"/jobtilbudprint_"<%=lto%>"_"<%= %>jobtilbudsnr&"_"&filnavnDato&"_"&filnavnKlok&".pdf">her</a>
</td></tr></table>
-->

<!--#include file="../inc/regular/footer_inc.asp"-->
