<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<%


    
	
	   '**** S�gekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_showjob"


        kundeid = request("kundeid")

        strJobs = "<option value='0'>V�lg folder...</option>" 

        kundeSQLkri = "fo.kundeid = "& kundeid & " OR fo.kundeid = 0 "
        jobidSQL = ""

    	strSQL = "SELECT fo.kundeid AS kundeid, fo.navn AS foldernavn, "_
	    &" fo.id AS foid, fo.kundese, fo.jobid"_
	    &" FROM foldere fo "_
	    &" WHERE ("& kundeSQLkri &" "& jobidSQL &") ORDER BY foldernavn"

        
        
    

        oRec.open strSQL, oConn, 3
        while not oRec.EOF 

        strJobs = strJobs & "<option value='"& oRec("foid") &"'>"& oRec("foldernavn") &"</option>" 


        oRec.movenext
        wend 
        oRec.close

       
        '*** ��� **'
        call jq_format(strJobs)
        strJobs = jq_formatTxt
        
        response.write strJobs

        end select
        response.end
        end if

if len(session("user")) = 0 then
	%>
	
	<%
	errortype = 5
	call showError(errortype)
	else
	
	func = request("func")
	
select case func
case "opret"

'*** Finder ud af om dokumentet skal tilknyttes en kunde ller et job ****
useid = request("type")

'if useid = "job" then
'jobid = request("id")
'kundeid = 0
'else
'jobid = 0
'kundeid = request("id")
'end if


'** Skal v�re nul da filer er afh�ngeig af folder og ikke job mere.
'** Jobafh�ngighed s�ttes p� folder niveau.
'** Sesseion jobid s�ttes i step 1 s� dervendes tilbage til den korrekte visning i filarkivet

if lto = "intanet - local" or lto = "nt" then
jobid = request("id")
filepath1 = request("filepath1")
else
jobid = 0 
kundeid = request("kundeid")
end if

if len(request("iid")) <> 0 then
iid = request("iid")
else
iid = 0
end if

if len(request("FM_adg_kunde")) <> 0 then
adg_kunde = 1
else
adg_kunde = 0
end if

if len(request("FM_adg_alle")) <> 0 then
adg_alle = 1
else
adg_alle = 0
end if

'Response.write "request(nomenu)" & nomenu
if lto = "intranet - local" or lto = "nt" then

        strSQLupdnt = "INSERT INTO filer (jobid, type, oprses) "_
        & " VALUES ("& jobid &", 1, 'x' )"

        'response.Write "tegn: " & strSQLupdnt
        'response.Flush
        oConn.Execute(strSQLupdnt)
%>
        <div class="wrapper">
            <div class="content">
                <div class="container">
                <div class="portlet">
                     <h3 class="portlet-title"><u>Upload fil</u></h3>
                    <div class="portlet-body">
                        <FORM ENCTYPE="multipart/form-data" ACTION="../timereg/upload_bin.asp" METHOD="POST">

                            <div class="row">
                                <div class="col-lg-2">
                                    
                                    <INPUT NAME="fileupload1" TYPE="file" class="btn btn-default btn-file">
			                        <!-- Der kan kun uploades 1 fil, pga x i ovenst�ende sql statement -->
			                        <!--
			                        <INPUT NAME="fileupload2" TYPE="file" style="width:200px;"><br>
			                        <INPUT NAME="fileupload3" TYPE="file" style="width:200px;"><br>-->
			                        <!--<INPUT NAME="Action"  type="image" src="../ill/opretpil.gif">-->
                                    <input type="submit" value=" Upload >> " />
                                </div>
                            </div>


                        </FORM>
                    </div>
                </div>
                </div>
            </div>
        </div>

       <div id="sindhold" style="position:absolute; left:20px; top:20px; visibility:hidden;">
		
			<h4>Upload fil (Step 2)</h4>
			V�lg den fil du vil uploade.<br>
			<font color="red"><b>!</b></font>&nbsp;Eksisterende filer overskrives altid n�r en ny fil, af samme navn, uploades.<br>
			<FORM ENCTYPE="multipart/form-data" ACTION="../timereg/upload_bin.asp" METHOD="POST">
            <table cellspacing="4" cellpadding="2" border="0">
			<tr>
			<td>
			<INPUT NAME="fileupload1" TYPE="file" style="width:400px;"><br><br />
			<!-- Der kan kun uploades 1 fil, pga x i ovenst�ende sql statement -->
			<!--
			<INPUT NAME="fileupload2" TYPE="file" style="width:200px;"><br>
			<INPUT NAME="fileupload3" TYPE="file" style="width:200px;"><br>-->
			<!--<INPUT NAME="Action"  type="image" src="../ill/opretpil.gif">-->
            <input type="submit" value=" Upload >> " /><br>&nbsp;
            </td>
            </tr>
			
			
            <tr><td><b>Note:</b><br>
			<br>
			<li>Klik p� "Browse"-knappen. <li>Find den �nskede fil p� din harddisk. Mark�r det, og klik p� "Open". <li>Klik p� "Upload". <br>
			
			<br><br><a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
			
            </table>
            </FORM>
            <br><br>	
		</div>


<%
        
else

if request("FM_folderid") <> "0" then
folderid = request("FM_folderid")

        
		strSQLupd = "INSERT INTO filer (type, oprses, jobid, kundeid, folderid, adg_kunde, adg_alle, incidentid) "_
		&" VALUES ('"& Request("FM_filtype")&"', 'x', "& jobid &", "& kundeid &", "_
		&" "& folderid &", "& adg_kunde &", "& adg_alle &", "& iid &")"
		'Response.write strSQLupd
		'response.flush
        oConn.Execute(strSQLupd)
		
		%>
		<!--#include file="../inc/regular/header_hvd_inc.asp"-->
	
		<div id="sindhold" style="position:absolute; left:20px; top:20px; visibility:visible;">
		
			<h4>Upload fil (Step 2)</h4>
			V�lg den fil du vil uploade.<br>
			<font color="red"><b>!</b></font>&nbsp;Eksisterende filer overskrives altid n�r en ny fil, af samme navn, uploades.<br>
			<FORM ENCTYPE="multipart/form-data" ACTION="../timereg/upload_bin.asp" METHOD="POST">
            <table cellspacing="4" cellpadding="2" border="0">
			<tr>
			<td>
			<INPUT NAME="fileupload1" TYPE="file" style="width:400px;"><br><br />
			<!-- Der kan kun uploades 1 fil, pga x i ovenst�ende sql statement -->
			<!--
			<INPUT NAME="fileupload2" TYPE="file" style="width:200px;"><br>
			<INPUT NAME="fileupload3" TYPE="file" style="width:200px;"><br>-->
			<!--<INPUT NAME="Action"  type="image" src="../ill/opretpil.gif">-->
            <input type="submit" value=" Upload >> " /><br>&nbsp;
            </td>
            </tr>
			
			
            <tr><td><b>Note:</b><br>
			<br>
			<li>Klik p� "Browse"-knappen. <li>Find den �nskede fil p� din harddisk. Mark�r det, og klik p� "Open". <li>Klik p� "Upload". <br>
			
			<br><br><a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
			
            </table>
            </FORM>
            <br><br>	
			</div>
			<%

 else
 %>
 
		<!--#include file="../inc/regular/header_hvd_inc.asp"-->
		<%
		useleftdiv = "m"
		errortype = 53
		call showError(errortype)
 
 end if

end if ' lto: nt 

case else




nomenu = request("nomenu")
if nomenu = "1" then
session("oprsesThis") = "open"
else
session("oprsesThis") = "close"
end if

'Response.write session("oprsesThis")

kundeid = request("kundeid")
'if kundeid <> 0 then
kundeSQLkri = "fo.kundeid = "& kundeid & " OR fo.kundeid = 0 "
'else
'kundeSQLkri = "fo.kundeid <> -1"
'end if

session("this_kid") = kundeid 

jobid = request("jobid")
session("this_jobid") = jobid

if request("sdsk") = "1" then
sdsk = 1
else
sdsk = 0
end if


'Response.write "sdsk" & sdsk

if len(request("iid")) <> 0 then
iid = request("iid")
else
iid = 0
end if
 
if jobid <> 0 then
jobidSQL = "AND fo.jobid = " & jobid
else
jobidSQL = ""
end if

%>
<!--#include file="../inc/regular/header_hvd_inc.asp"-->

<SCRIPT language="javascript" src="inc/upload_jav.js"></script>

	<!-------------------------------Sideindhold------------------------------------->
	
	<div id="sindhold" style="position:absolute; left:20px; top:20px; visibility:visible;">
	
	<h4>Upload fil (Step 1)</h4> 
		<table cellspacing="2" cellpadding="1" border="0">
		<tr><form action="upload.asp?menu=fob&func=opret&type=<%=request("type")%>&id=<%=request("id")%>&iid=<%=iid%>&jobid=<%=jobid%>" method="post">
		<td>
        
            <%if cint(kundeid) <> 0 then 
                %>
                 <br><b>Filtype</b>: Office dokument, billede ell. firma logo)<br>
		<select name="FM_filtype" style="width:300px;">
		<%select case request("type")
		case "kundelogo" 
		logoSel = "SELECTED"
		offSel = ""
		bilSel = ""
		case "pic"
		logoSel = ""
		offSel = ""
		bilSel = "SELECTED"
		case else
		logoSel = ""
		offSel = "SELECTED"
		bilSel = ""
		end select%>
		<option value="5" <%=offSel%>>Office Dokument (.docx, .doc, pdf, .xlsx, .xsl)</option>
		<option value="7" <%=bilSel%>>Billede (.gif, .jpg, .png)</option>
		<option value="1" <%=logoSel%>>Firma Logo</option>
		</select><br><br>

            <input type="hidden" name="kundeid" value="<%=kundeid %>" /> 
                
                <%
             else%>
            
             <br /><b>Kunde:</b><br />
        <select name="kundeid" id="kundeid" size="1" style="width:325px;">
		<option value="0">Alle</option>
		<%
				strSQL = "SELECT Kkundenavn, Kkundenr, Kid FROM kunder WHERE Kid <> 0 ORDER BY Kkundenavn"
				oRec.open strSQL, oConn, 3
				while not oRec.EOF
				
				if cint(kundeid) = cint(oRec("Kid")) then
				isSelected = "SELECTED"
				else
				isSelected = ""
				end if
				%>
				<option value="<%=oRec("Kid")%>" <%=isSelected%>><%=oRec("Kkundenavn")%></option>
				<%
				oRec.movenext
				wend
				oRec.close
				%>
		</select>    <br />
<br />
              <input type="hidden" name="FM_filtype" value="5" /> 
            <%end if %>
            
            
   
		
	
	
	<%if sdsk <> 1 then%>
   

	<b>Folder:</b><br />
	
	<%
	strSQL = "SELECT fo.kundeid AS kundeid, fo.navn AS foldernavn, "_
	&" fo.id AS foid, fo.kundese, fo.jobid"_
	&" FROM foldere fo "_
	&" WHERE ("& kundeSQLkri &" "& jobidSQL &") ORDER BY foldernavn"

   'Response.Write strSQL
   'Response.flush

   %>
   <select name="FM_folderid" id="FM_folderid" style="width:300px;"> 
	<option value="0">V�lg folder...</option>
   <%
	oRec.open strSQL, oConn, 3 
	while not oRec.EOF 
	
    if cdbl(jobid) = oRec("jobid") then
    foSEL = "SELECTED"
    else
    foSEL = ""
    end if
    %>
	<option value="<%=oRec("foid")%>" <%=foSEL %>><%=oRec("foldernavn")%></option>
	<%
	
    oRec.movenext
	wend
	oRec.close %>
	</select>&nbsp;
	<%else%>
	<input type="hidden" name="FM_folderid" id="FM_folderid" value="500">
	<%end if%>

	
	<br><br>
	<b>Adgangs rettigheder:</b><br>
	Hvilke brugergrupper skal have adgang til dette dokument:<br>
	<input type="checkbox" name="FM_adg_kunde" id="FM_adg_kunde" value="1">&nbsp;Kunder. <br>
	<img src="../ill/checkboxadmin.gif" width="19" height="19" alt="" border="0"> TimeOut System Administrator(er).<br>
	<input type="checkbox" name="FM_adg_alle" id="FM_adg_alle" value="1" CHECKED> TimeOut Alle brugere. 
	<br><br>
	<input type="image" src="../ill/naestepil.gif">
		
<br><br>&nbsp;</td></tr></form></table>
<a href="Javascript:history.back()"><img src="../ill/pil_left.gif" width="22" height="22" alt="" border="0"></a>
<br><br>	
</div>

<%end select%>
<%end if%>
<!--#include file="../inc/regular/footer_inc.asp"-->