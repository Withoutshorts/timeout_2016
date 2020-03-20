<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/header_hvd_inc.asp"-->	

<!--#include file="../inc/regular/global_func.asp"--> 

<%  if len(trim(request("glt_pw_email"))) <> 0 then
    strEmail = request("glt_pw_email")
    else
    strEmail = "dsfwer23411"
    end if
    
    if len(trim(request("func"))) > 0 then
    func = request("func")
    else
    func = "red"
    end if
    

    if len(trim(request("lto"))) > 0 then
    lto = request("lto")
    else
    lto = "xxx"
    end if

    Private Function Encrypt(ByVal string)
	    Dim x, i, tmp
	    For i = 1 To Len( string )
	    x = Mid( string, i, 1 )
	    tmp = tmp & Chr( Asc( x ) + 1 )
	    Next
	    tmp = StrReverse( tmp )
	    Encrypt = tmp
    End Function
    
    public randomNumber
    Sub Test2()
        nLow = 10000000
        nHigh = 1000000000
        Set objRandom = CreateObject("System.Random")
        randomNumber = objRandom.Next_2(nLow, nHigh)

        randomNumber = replace(randomNumber, "78","")
        randomNumber = replace(randomNumber, "87","")

    End Sub

    
     %>

<div id="sindhold" style="position:absolute; left:20px; top:20px; width:400px; height:200px; visibility:visible;">


    <%if func <> "sendpw" then %>

    <form action="sendpw.asp?func=sendpw&lto=<%=lto %>" method="post">
        <h4>Glemt logind eller password</h4>
        Indtast din email adresse og du vil få tilsendt dit logind og password til TimeOut.<br />
        Bemærk venligst at dit password vil blive nulstillet.<br /><br />

                Email: <input type="text" name="glt_pw_email" style="width:150px;" /> <input type="submit" value="Send >>"/>

            </form><br /><br />



<% else

                        
                        strLogin = "-1"
                            strSQL = "SELECT mid, email, mnavn, login FROM medarbejdere WHERE email = '" & strEmail &"'"
						    oRec.open strSQL, oConn, 3
						    if not oRec.EOF then
							    strPw = "21"& right(oRec("mnavn"), 2) & "X2"& left(oRec("login"),2) & "C" & right(now,2)
                                strNavn = oRec("mnavn")
                                strLogin = oRec("login")
                                medarbejderid = oRec("mid")
                                    
                                    strSQLupd = "UPDATE medarbejdere SET pw = MD5('"& strPw &"') WHERE mid = "& oRec("mid")
                                    oConn.execute(strSQLupd)
                
						    end if
						    oRec.close


                
                        if strLogin = "-1" then
                        %>
    
                        <div id="sindhold" style="position:absolute; left:20px; top:20px; width:400px; height:200px; visibility:visible;">
                            <h4>Bruger ikke fundet!</h4>
                            Du har indtastet en email adresse (<%=strEmail %>) der ikke findes i systemet.<br />
                            Prøv igen eller kontakt din TimeOut administrator.<br /><br />
                            <a href="Javascript:history.back()"> << Tilbage</a>

                        </div>
                        <%



                        response.end
                        end if




    

                        'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
    					
					    ''Sætter Charsettet til ISO-8859-1 
					    'Mailer.CharSet = 2
					    '' Afsenderens navn 
					    'Mailer.FromName = "TimeOut"
					    '' Afsenderens e-mail 
					    'Mailer.FromAddress = "timeout_no_reply@outzource.dk"
					    'Mailer.RemoteHost = "webout.smtp.nu" '"webmail.abusiness.dk"
					    ' Modtagerens navn og e-mail
					    'Mailer.AddRecipient strNavn, strEmail
					    'Mailer.AddBCC "Support", "support@outzource.dk" 
					    ' Mailens emne
					    'Mailer.Subject = "TimeOut - Password"

                        
                        
                       Set myMail=CreateObject("CDO.Message")
                        myMail.Subject="TimeOut - Nyt Password"
                        myMail.From="timeout_no_reply@outzource.dk"

                        if len(trim(strEmail)) <> 0 then
                        myMail.To= ""& strNavn &"<"& strEmail &">"
                        end if

                        
                        call glemtpassword_visning

                        if cint(int_glemtpassword_visning) = 1 then
                            stampdate = year(now) &"-"& month(now) &"-"& day(now)
                            stamptime = hour(now) &":"& minute(now) & ":" & second(now)
                            strSQLForgotPassword_log = "INSERT INTO forgot_password SET medid ="& medarbejderid &", stampdate ='"& stampdate &"', stamptime ='"& stamptime &"'" 
                            'response.Write "strSQLForgotPassword_log " & strSQLForgotPassword_log    
                            'response.End
                            oConn.execute(strSQLForgotPassword_log)

                            'Henter licens kode
                            strlisenskey = ""
                            strSQLlicenskode = "SELECT * FROM licens"
                            oRec.open strSQLlicenskode, oConn, 3
                            if not oRec.EOF then
                                strlisenskey = oRec("key")
                            end if
                            oRec.close
                            

                            linkID = "2e60c62249a%01dba7%da=e797df7d4f436" 
 
                            call Test2()

                            linkID = linkID &"=0"& randomNumber&"&78="&Encrypt(medarbejderid)&"&"

                            call Test2()

                            linkID = linkID &"=0"& randomNumber & "&87="&lto&"&"
                            linkID = linkID & "key="&strlisenskey&"&"

                            call Test2()

                            linkID = linkID & "523af=537946b7%9c4f8369ed39%ba78605"

                            linkID = linkID & randomNumber

                            newpasswordlink = "<a href='https://timeout.cloud/timeout_xp/wwwroot/ver2_14/to_2015/newpassword.asp?func=newpassword&"&linkID&"'>Klik her</a>"

                            strmailtekst = "Hej " & strNavn & "<br>" & "<br>"
                            
                            strmailtekst = strmailtekst & newpasswordlink & " for at ændre dit password: " & "<br>"
                             
                            strmailtekst = strmailtekst & "<br>" & "<br>"                           
                            strmailtekst = strmailtekst & "Med venlig hilsen"& "<br>" & "<br>" & strEditor & "<br>" 
                            

                            myMail.HTMLBody = "<html><head></head><body>"& strmailtekst &"</body></html>" 'strmailtekst

                        else
                            
                            ' Selve teksten
					        myMail.TextBody = "" & "Hej "& strNavn & vbCrLf & vbCrLf _ 
					        & "Dit brugernavn er: " & strLogin & " og dit password er: " & strPw & vbCrLf & vbCrLf _ 
					        & "Gem disse oplysninger, til du skal logge ind i TimeOut."  & vbCrLf _ 
					        & "Du kan altid selv ændre dem når du er logget på systemet." & vbCrLf & vbCrLf _ 
					        & "Adressen til TimeOut er: https://timeout.cloud/"&lto&""& vbCrLf & vbCrLf _ 
					        & "Med venlig hilsen"& vbCrLf & vbCrLf & strEditor & vbCrLf 

                        end if

                       
                        myMail.Configuration.Fields.Item _
                        ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
                        'Name or IP of remote SMTP server
                                    
                                    if instr(request.servervariables("LOCAL_ADDR"), "195.189.130.210") <> 0 then
                                       smtpServer = "webout.smtp.nu"
                                    else
                                       smtpServer = "formrelay.rackhosting.com" 
                                    end if
                    
                                    myMail.Configuration.Fields.Item _
                                    ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= smtpServer

                        'Server port
                        myMail.Configuration.Fields.Item _
                        ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25
                        myMail.Configuration.Fields.Update

                        if len(trim(strEmail)) <> 0 then
                        myMail.Send

                              %>
                                <h4>Dit logind og password er tilsendt</h4>
                                Du password er nulstillet. Du kan altid ændre dit password ved at redigere din profil. <br /><br />
    
                                Emailen er sendt til: <b><%=strEmail %></b> og bør være fremme indenfor et par minutter. Husk evt. at tjekke din spam-mappe.<br /><br />

                                Du kan nu lukke dette vindue ned. <br /><br />
    
                                Med venlig hilsen <br />TimeOut
                                <%

                        else
                                %>
                                <h4>Der opstod en fejl</h4>
                                Kontakt Outzource support på <a href="mailto:support@outzource.dk">support@outzource.dk</a>
                                <%
                        end if
                        set myMail=nothing
                    
                           


						  

                           
					    
    					
					  
    
    
    end if
     %>	


</div>
<!--#include file="../inc/regular/footer_inc.asp"-->
 
