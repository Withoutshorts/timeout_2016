<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
 
<!--include file="../inc/regular/header_lysblaa_inc.asp"-->
<!--#include file="../inc/regular/header_hvd_inc.asp"-->

<%

'faknr = request("faknr")
lto = request("lto")

'jobid = request("jobid")
'aftid = request("aftid")
'id = request("id")
'kid = request("kid")

'Response.Write jobid
'Response.End 
if len(request("makepdf")) <> 0 then
makepdf = request("makepdf")
else
makepdf = 0
end if

if len(trim(request("doctype"))) <> 0 then
doctype = request("doctype")
else
doctype = ""
end if

if len(request("fakids")) <> 0 then
fakids = split(request("fakids"), ",")
else
fakids = 0
end if

if len(trim(request("kpid"))) <> 0 then
   kpid = request("kpid")

        strSQL = "SELECT kp.navn, kp.email FROM kontaktpers AS kp "_
        &" WHERE kp.id = "& kpid 
        oRec.open strSQL, oConn, 3
        if not oRec.EOF then

        emAf = oRec("navn")
        emAFmail = oRec("email")

        end if
        oRec.close 
		

else
   kpid = session("mid")

        '** emailmodtager / afsender
        strSQL = "SELECT mnavn, email FROM medarbejdere WHERE mid = "& kpid
        oRec.open strSQL, oConn, 3
        if not oRec.EOF then

        emAf = oRec("mnavn")
        emAFmail = oRec("email")

        end if
        oRec.close

end if 



select case lto
case "xexecon", "ximmenso"
emMo = "Bogholderi"
emMomail = "bogholderi@execon.dk"
case else
emMo = emAf
emMomail = emAFmail
end select


call TimeOutVersion()

    
    if request.servervariables("PATH_TRANSLATED") <> "C:\www\timeout_xp\wwwroot\ver2_1\timereg\erp_make_pdf_multi.asp" then
                               
    ''Opretter en instans af fil object **'
    Set fs=Server.CreateObject("Scripting.FileSystemObject")
    'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
    
                                 '' Sætter Charsettet til ISO-8859-1 
                                 'Mailer.CharSet = 2
                                 '' Afsenderens navn 
                                 'Mailer.FromName = "TimeOut Email Service"
                                 '' Afsenderens e-mail 
                                 'Mailer.FromAddress = "timeout_no_reply@outzource.dk"
                                 'Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
                                 'Mailer.ContentType = "text/html"
                                 
                                ''Mailens emne
                                 'Mailer.Subject = "Faktura PDF/XML mail"
                                 ''Modtagerens navn og e-mail
                                 'Mailer.AddRecipient emMo, emMomail



                        Set myMail=CreateObject("CDO.Message")
                        myMail.From="timeout_no_reply@outzource.dk"

                        if len(trim(emMo)) <> 0 then
                        myMail.To= ""& emMo &"<"& emMomail &">"
                        else
                        myMail.To= "TimeOut Support <support@outzource.dk>"
                        end if

                        
                       

                        
                        myMail.Configuration.Fields.Item _
                        ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                        'Name or IP of remote SMTP server
                                    
                       smtpServer = "formrelay.rackhosting.com" 
                                   
                       myMail.Configuration.Fields.Item _
                        ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                        'Server port
                        myMail.Configuration.Fields.Item _
                        ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                        myMail.Configuration.Fields.Update

                       
                    
    

    call meStamdata(session("mid"))
 

    f = 0
    for f = 1 to UBOUND(fakids)
        
        if len(trim(fakids(f))) <> 0 then
        fakids(f) = fakids(f)
        else
        fakids(f) = 0
        end if
        
        strSQL = "SELECT faknr FROM fakturaer WHERE fid = " & fakids(f)
        oRec.open strSQL, oConn, 3
        faknr = 0
        if not oRec.EOF then
        
        faknr = oRec("faknr")
        
        End if
        oRec.close


         
         strFaktypeNavn = "xx"
         if cdbl(fakids(f)) <> 0 then 
         
             strknavnSQL = "SELECT k.kkundenavn, fid, f.faktype, f.jobid, f.aftaleid, f.fakadr, f.afsender, "_
             &" afs.betbet as afsbetbet, afs.kkundenavn AS afskkundenavn, afs.adresse As afsadresse, afs.postnr as afspostnr, "_
             &" afs.city as afscity, afs.land AS afsland, afs.email As afsemail, afs.telefon AS afstelefon FROM fakturaer f"_ 
             &" LEFT JOIN kunder as k ON (k.kid = fakadr) "_
             &" LEFT JOIN kunder as afs ON (afs.kid = afsender) "_
             &" WHERE fid = " & fakids(f)
             
                'response.write "<br>" & strknavnSQL
                'response.flush
    
             oRec.open strknavnSQL, oConn, 3
             
             strknavn = "xx"
             if not oRec.EOF then
             strknavn = oRec("kkundenavn") 
             
             call kundenavnPDF(strknavn)
             strknavn = strKundenavnPDFtxt

             if oRec("faktype") = 1 then
             strFaktypeNavn = "kreditnota" 'erp_txt_002
             else
             strFaktypeNavn = "faktura" 'erp_txt_001
             end if

                        
                    jobid = oRec("jobid")
                    aftid = oRec("aftaleid")
                    kid = oRec("fakadr")         
       
                 

             afsenderTxt = "<br><br>Denne faktura er oprettet og afsendt fra timeOut (www.outzource.dk) på vegne af:<br><br>"_
             &""& oRec("afskkundenavn") &"<br>"& oRec("afsadresse") &"<br>"& oRec("afspostnr") &", "& oRec("afscity") &"<br>"& oRec("afsland") &"<br>"& oRec("afsemail") & ", tel: "& oRec("afstelefon") &""_
             &"<br><br>Standard betalingsbetingelser:<br>"& oRec("afsbetbet") & "<br><br>"_ 
             &"<br><br><br>Med venlig hilsen<br>"& oRec("afskkundenavn") &"<br>"& meNavn & "<" & meEmail& "><br><br>"& now

             end if
             oRec.close

             

         else
         strknavn = "oo"
         strFaktypeNavn = "oo"
         end if



        
            '*** Tilføjer PDF **'
            pdfurl = "d:\\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
	        
	       

        	If (fs.FileExists(pdfurl))=true Then
                    
            '        Mailer.AddAttachment "d:\\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
                     myMail.AddAttachment "d:\\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
                     bodyTXT = bodyTXT & vbcrlf &""&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
                    
                    
            else
                        
                        '*** Opretter PDF og afsender **'
                        '*** Ved godkend faktura, Execon, Immenso.. ***'
                        if cint(makepdf) = 1 then
                        
                        
                        Set Pdf = Server.CreateObject("Persits.Pdf")
                        Set Doc = Pdf.CreateDocument


                         select case lto
                         case "xdencker"
                         'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?&visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=0, RightMargin=0, TopMargin=0, BottomMargin=0, PageWidth=635, PageHeight=903"
                         Doc.ImportFromUrl "http://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=0, RightMargin=0, TopMargin=0, BottomMargin=0, PageWidth=635, PageHeight=903"

                         case "xsynergi1", "xintranet - local"
                         'Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/job_print.asp?media=pdf&nosession=9999&key="&session("lto")&"&func=print&lto="&lto&"&id="&id&"&kid="&kid&"&pdfvalid="&pdfvalid&"", "LeftMargin=20, RightMargin=0, TopMargin=0, BottomMargin=0, DrawBackground=True,"
         
                         
                         Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?&visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=40, RightMargin=0, TopMargin=40, BottomMargin=0, PageWidth=635, PageHeight=903, DrawBackground=True"
                         'Doc.ImportFromUrl "http://localhost/timeout_xp/timereg/erp_fak_godkendt_2007.asp?nosession=9999&media=pdf&jobid="&jobid&"&aftid="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"", "LeftMargin=40, RightMargin=0, TopMargin=40, BottomMargin=0, PageWidth=635, PageHeight=903, DrawBackground=True"

       
                         case else

                          
                             '*** NY SERVER ****
                              'Doc.ImportFromUrl "http://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")
                               Doc.ImportFromUrl "https://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&fakids(f)&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")
                               'Doc.ImportFromUrl Server.MapPath("../inc/upload/"&lto&"/"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf")
                        
                                 'response.write "http://timeout.cloud/timeout_xp/wwwroot/"&toVer&"/timereg/erp_opr_faktura_fs.asp?visminihistorik=1&visfaktura=2&visjobogaftaler=1&nosession=9999&media=pdf&FM_job="&jobid&"&FM_aftale="&aftid&"&id="&fakids(f)&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")
                                 'response.flush


                         end select


                        ''Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/erp_fak_godkendt_2007.asp?nosession=9999&print=pdf&jobid="&jobid&"&aftid="&aftid&"&id="&id&"&FM_usedatointerval="&request("FM_usedatointerval")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")&"&FM_fakint_ival="&request("FM_fakint_ival")&"&FM_aftnr="&sogaftnr&"&kid="&kid
                        ''Doc.ImportFromUrl "https://outzource.dk/timeout_xp/wwwroot/ver2_10/timereg/erp_fak_godkendt_2007.asp?nosession=9999&media=pdf&jobid="&jobid&"&aftid="&aftid&"&id="&id&"&key="&session("lto")&"&FM_start_dag_ival="&request("FM_start_dag_ival")&"&FM_start_mrd_ival="&request("FM_start_mrd_ival")&"&FM_start_aar_ival="&request("FM_start_aar_ival")&"&FM_slut_dag_ival="&request("FM_slut_dag_ival")&"&FM_slut_mrd_ival="&request("FM_slut_mrd_ival")&"&FM_slut_aar_ival="&request("FM_slut_aar_ival")
                        ''Doc.ImportFromUrl "http://localhost/timeout_xp/timereg/erp_fak_godkendt_2007.asp?nosession=9999&print=pdf&jobid="&jobid&"&aftid="&aftid&"&id="&id
                        
                        
                       


                        Filename = Doc.Save(Server.MapPath("../inc/upload/"&lto&"/"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"), True)  
                        'Mailer.AddAttachment "d:\\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
                        myMail.AddAttachment "d:\\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
                        bodyTXT = bodyTXT & vbcrlf &""&strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
               
                        
                        else
                        
                        Response.Write("<br>File not found on TimeOut server!<br>")
                        Response.Flush
                       
                        end if
                    
                    
                   
                    
            End If 'PDF URL Exist
            
            
             '*** Tilføjer XML **'
             'xmlurl = "d:\\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&strFaktypeNavn&"_xml_"&lto&"_"&faknr&"_"&strknavn&".xml"
	    
	 
        	'If (fs.FileExists(xmlurl))=true Then
            '       Mailer.AddAttachment "d:\\webserver\wwwroot\timeout_xp\wwwroot\"&toVer&"\inc\upload\"&lto&"\"&strFaktypeNavn&"_xml_"&lto&"_"&faknr&"_"&strknavn&".xml"
            '       bodyTXT = bodyTXT & vbcrlf &"faktura_xml_"&lto&"_"&faknr&"_"&strknavn&".xml"
                    
            'End If
       
       
      
    next
    
    
                 
          
                 
                 ''Selve teksten
                 'bodyTXT =  bodyTXT & vbcrlf & vbcrlf &"Denne mail er oprettet af:"& vbcrlf & emAf & ", "& emAFmail
                 
                 'Mailer.BodyText = "Din mail med PDF dokumenter og XML filer er klar. Følgende filer er vedhæftet:" &vbcrlf & bodyTXT 
                  
   
                        
                        'to_url = "https://timeout.cloud"
                        myMail.Subject=""& strFaktypeNavn &" " & faknr & ", " & strknavn
                        myMail.HTMLBody = "<HTML><HEAD></head><BODY>Kære kunde<br>Hermed fremsendes "& strFaktypeNavn &" " & faknr & ".<br>" & afsenderTxt & bodyTXT & "</Body></html>"
            
                        '"" & medarb_txt_009 &" " & strNavn & vbCrLf _ 
					    '& medarb_txt_118 & vbCrLf _ 
					    '& medarb_txt_012 &" "& strLogin &" "& medarb_txt_013 &" "& strPw & vbCrLf & vbCrLf _ 
					    '& medarb_txt_014  & vbCrLf _ 
					    '& medarb_txt_015 & vbCrLf & vbCrLf _ 
					    '& medarb_txt_119 &": "& to_url&"/"&lto&""& vbCrLf & vbCrLf _ 
					    '& medarb_txt_120 & vbCrLf & vbCrLf & strEditor & vbCrLf 

       
                    if len(trim(emMomail)) <> 0 then
                    myMail.Send
                    end if
                  

                  'If Mailer.SendMail Then
                        
                        
                   '     if makepdf = 1 then
                    '      Response.Write("<script language=""JavaScript"">window.opener.location.reload();</script>")
                    '        Response.Write("<script language=""JavaScript"">window.opener.top.frames['erp2_2'].location.reload();</script>")
                            
                   '         Response.Write("<script language=""JavaScript"">window.close();</script>")
                           
                    '       else
                     
                       %>
                       
                       
                  <div style="padding:10px;">
                  <table border=0 cellspacing=1 cellpadding=0 width="200">
	            <tr><td valign=top bgcolor="#ffffff" style="padding:5px;">
	            <img src="../ill/outzource_logo_200.gif" />
	            </td>
	            </tr>
	            <tr>
	            <td valign=top bgcolor="#ffffff" style="padding:5px 5px 5px 15px;">
	            <!-- Mail med vedhæftede PDF'er og XML filer er afsendt. -->
                Invoice mail with PDF attachment sent to seleted recipient.<br />
                    <br />
                    <% response.write strFaktypeNavn&"_"&lto&"_"&faknr&"_"&strknavn&".pdf"
                    response.flush %>
                    
	            </td></tr>
	            </table>
                </div>
                  <%    'end if
                            
                  'Else
                  'Response.Write "Fejl...<br>" & Mailer.Response
                  'End if
    
    
    set fs=nothing
    set myMail=nothing    

    end if '** C:
    
%>



<!--#include file="../inc/regular/footer_inc.asp"-->
