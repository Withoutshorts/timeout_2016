

<%
response.buffer = true
%>
<!--#include file="../inc/connection/conn_db_inc.asp"-->
<%   
    
    
       '**** Søgekriterier AJAX **'
        'section for ajax calls
        if Request.Form("AjaxUpdateField") = "true" then
        Select Case Request.Form("control")
        case "FN_pic"

         
                jobnr = request("jq_jobnr") 

	                                strSQL = "SELECT id, filnavn FROM filer WHERE filertxt = "& jobnr
                                    
                                    filnavn = "Picture Not Found"
	                                oRec.open strSQL, oConn, 3
	                                j = 0
	                                if not oRec.EOF then
                                    if len(trim(oRec("filnavn"))) <> 0 then                       
	            
                                          
                                    filnavn = "<img src='../inc/upload/nt/"&oRec("filnavn")&"' alt='' border='0'>"                              


                                    j = j + 1                                    
	                                end if
                                    end if	                                
	                                oRec.close 


                                 response.write filnavn
                
        
        response.end

        case "FN_afd"
        
            if len(trim(request("jq_kid"))) <> 0 then
            kid = request("jq_kid")
            else 
            kid = 0
            end if
            strOptions = "<option value=0>Choose..</option>"
            

            strSQL = "SELECT id, afdeling FROM kontaktpers WHERE kundeid = "& kid
            
            oRec.open strSQL, oConn, 3
            while not oRec.EOF
        
              if len(trim(oRec("afdeling"))) <> 0 then
              strOptions = strOptions & "<option value="& oRec("id") &">"& oRec("afdeling") &"</option>"
              end if 
				
			oRec.movenext
			wend
			oRec.close

            '*** ÆØÅ **'
            call jq_format(strOptions)
            strOptions = jq_formatTxt

            response.write strOptions

         case "FN_betlev"
        
            if len(trim(request("jq_kid"))) <> 0 then
            kid = request("jq_kid")
            else 
            kid = 0
            end if
            

            strSQL = "SELECT betbetint, levbet FROM kunder WHERE kid = "& kid
            
            'response.Write strSQL
            'response.end

            oRec.open strSQL, oConn, 3
            if not oRec.EOF then
        
            betbetint = oRec("betbetint")

			
			end if
			oRec.close

           

            response.write betbetint
            response.end



       
        case "FN_sup"
        
            if len(trim(request("jq_origin"))) <> 0 then
            origin = request("jq_origin")
            else 
            origin = ""
            end if
            
            lndTxt = origin
        
            lndTxt2SQL = " OR land = 'xx'"

         
            if lcase(lndTxt) = "hongkong" then
            lndTxt2SQL = " OR land = 'Hong Kong'"
            end if

            strOptions = "<option value=0>Choose..</option>"
            strSQL = "SELECT kid, kkundenavn FROM kunder WHERE (land = '"& lndTxt &"' "& lndTxt2SQL &") AND useasfak = 6 ORDER BY kkundenavn" 'AND ketype = 'k'
            oRec.open strSQL, oConn, 3
            while not oRec.EOF
        
            
              strOptions = strOptions & "<option value="& oRec("kid") &">"& oRec("kkundenavn") &"</option>"
            
				
			oRec.movenext
			wend
			oRec.close

            '*** ÆØÅ **'
            call jq_format(strOptions)
            strOptions = jq_formatTxt

            response.write strOptions
            
         case "xFN_kpers" 'ændret til inter sales rep.
        

            if len(trim(request("jq_kid"))) <> 0 then
            kid = request("jq_kid")
            else 
            kid = 0
            end if

   

            strOptions = "<option value=0>Choose..</option>"

            strSQL = "SELECT mid, mnavn, email FROM medarbejdere WHERE mansat <> 2 ORDER BY mnavn"
            
            oRec.open strSQL, oConn, 3
            while not oRec.EOF

                if len(trim(oRec("email"))) <> 0 then
                emlTxt = " ("& oRec("email") &")"
                else
                emlTxt = ""
                end if           

              strOptions = strOptions & "<option value="& oRec("mid") &">"& oRec("mnavn") &" "& emlTxt &"</option>" 
				
			oRec.movenext
			wend
			oRec.close

             '*** ÆØÅ **'
            call jq_format(strOptions)
            strOptions = jq_formatTxt

            response.write strOptions
            
    
        end select

        response.end
    
        end if %>


<%if media = "" then %>
<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"--> 
<!--#include file="../inc/regular/topmenu_inc.asp"--> 
<%end if %>

<%if media = "print" then %>
<!--include file="../inc/regular/header_hvd_inc.asp"--> 
<%end if %>

<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->








<script src="js/job_nt_jav_201703.js"></script>


<%
'*** Sætter lokal dato/kr format.
 Session.LCID = 1030



    thisfile = "job_nt.asp"


  
    if len(session("user")) = 0 then

	errortype = 5
	call showError(errortype)

    response.end

	end if

   

func = request("func")    

 if len(trim(request("media"))) <> 0 then
    media = request("media")
    else
    media = ""
    end if
    


%>




<%
if media <> "print" then    
    call menu_2014
end if
    


 
    
%>

<style>

    table, tr, td, .tablecolor 
    {
        color:black;
        padding:0 15px 10px 0px;
    }
    
    
    
    /* The Modal (background) */
    .modal {
        display: none; /* Hidden by default */
        position: fixed; /* Stay in place */
        z-index: 1; /* Sit on top */
        padding-top: 100px; /* Location of the box */
        left: 0;
        top: 0;
        width: 100%; /* Full width */
        height: 100%; /* Full height */
        overflow: auto; /* Enable scroll if needed */
        background-color: rgb(0,0,0); /* Fallback color */
        background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
    }

    /* Modal Content */
    .modal-content {
        background-color: #fefefe;
        margin: auto;
        padding: 20px;
        border: 1px solid #888;
        width: 300px;
        height: 350px;
    }

    .picmodal:hover,
    .picmodal:focus {
    text-decoration: none;
    cursor: pointer;
}
   
</style>




<div id="wrapper">

<div class="content">
<div class="container" style="background-color:#FFFFFF; position:absolute; left:90px; top:-120px;">


<%        
    
select case func  
	case "slet"
	'*** Her spørges om det er ok at der slettes  ***
	
	
	
    id = request("id")

	slttxt = "<b>Delete order</b><br />"_
	&"You are aboute to delete an order. Are you sure you want do that?"
	slturl = "job_nt.asp?menu=job&func=sletok&id="&id
	
	call sltque(slturl,slttxt,slturlalt,slttxtalt,110,90)
	
	
	
	case "sletok"
	
        id = request("id")

	strSQL = "SELECT id, jobnavn, jobnr FROM job WHERE id = "& id &"" 
	oRec.open strSQL, oConn, 3
	if not oRec.EOF then
		strjobnr = oRec("jobnr")
		'*** Indsætter i delete historik ****'
	    call insertDelhist("job", id, oRec("jobnr"), oRec("jobnavn"), session("mid"), session("user"))
		
	
	end if
	oRec.close
	
    response.Write strjobnr & "<br>"
	
	strsqlfil = "SELECT filnavn FROM filer WHERE filertxt ="& strjobnr &""
    oRec.open strsqlfil, oConn, 3
    if not oRec.EOF then

    strfilnavn = oRec("filnavn")

    end if
    oRec.close
	
    strPath = "d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\upload\"&lto&"\" & strfilnavn
	Response.write strPath

    on Error resume Next 

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set fsoFile = FSO.GetFile(strPath)
	fsoFile.Delete

    oConn.execute("DELETE FROM filer WHERE filertxt = "& strjobnr &"")

	'Response.flush
	
	
	oConn.execute("DELETE FROM job WHERE id = "& id &"")
	
	
	
	Response.redirect "job_nt.asp?func=table"
	
case "sletfil"
'*** Her spørges om det er ok at der slettes en medarbejder ***
	%>
    
    
    <div class="container" style="width:500px;">
        <div class="porlet">
            
        
            
            <div class="portlet-body">
                <div style="text-align:center;"> You are about to <b>delete</b> a file. Is this correct?
                    <%if len(trim(slttxtb)) <> 0 then %>
                    <br /><br />&nbsp;
                    
                    <%end if %>
                </div><br />
                <div style="text-align:center;"><a  class="btn btn-primary btn-sm" role="button" href="job_nt.asp?func=sletfilok&id=<%=request("id")%>&filnavn=<%=request("filnavn")%>">&nbsp;Yes&nbsp;</a>&nbsp&nbsp&nbsp&nbsp<a class="btn btn-default btn-sm" role="button" href="Javascript:history.back()"><b>No</b></a>
                </div>
                <br /><br />
            </div>

        </div>
    </div>
    
    <%

case "sletfilok"
	'*** Her slettes en fil ***
	'ktv
	'strPath =  "E:\www\timeout_xp\wwwroot\ver2_1\upload\"&lto&"\" & Request("filnavn")
	'Qwert
	'strPath =  "d:\webserver\wwwroot\timeout_xp\wwwroot\ver2_1\inc\upload\"&lto&"\" & Request("filnavn")
    strPath = "d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\upload\"&lto&"\" & Request("filnavn")
	Response.write strPath
	
	on Error resume Next 

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set fsoFile = FSO.GetFile(strPath)
	fsoFile.Delete
	
	id = request("id")
    if len(trim(request("id"))) <> 0 then
    id = request("id")
    else
    id = 0
    end if

	oConn.execute("DELETE FROM filer WHERE id = "& id &"")
    response.Write "id er: " & id
	Response.redirect "job_nt.asp?func=table"

case "dbopr", "dbred"

    if len(trim(request("FM_kunde"))) <> 0 AND request("FM_kunde") <> "0" then
    kid = request("FM_kunde")
    else
    kid = 0
    end if

    if cdbl(kid) = 0 then
               
     errortype = 170
    call showError(errortype)
				
    response.End
    end if

    if len(trim(request("FM_kopierordre"))) <> 0 then
    kopier_ordre = 1
    func = "dbopr"
    else
    kopier_ordre = 0
    end if

    if len(trim(request("FM_jobnavn"))) <> 0 then
    jobnavn = replace(request("FM_jobnavn"), "'", "")
    else
    jobnavn = ""
    end if

    if cint(kopier_ordre) = 1 then
    jobnavn = jobnavn & " - COPY" 
    end if


    if len(trim(jobnavn)) = 0 then
               
     errortype = 171
    call showError(errortype)
				
    response.End
    end if


    jobid = request("FM_jobid")

    if len(trim(request("FM_jobnr"))) <> 0 then
    jobnr = request("FM_jobnr")
            
            jobnrFindes = 0        
            strSQL = "SELECT jobnr FROM job WHERE id <> "& jobid &" AND jobnr = "& jobnr &""
            oRec2.open strSQL, oConn, 3
            if not oRec2.EOF then
    
            jobnrFindes = 1

            end if
            oRec2.close

    else
    jobnr = 0
    end if


    if jobnr = 0 OR jobnrFindes = 1 then
               
     errortype = 172
    call showError(errortype)
				
    response.End
    end if


    jobstatus = request("FM_status")
  

    kpers = request("FM_kpers")
   
    department = request("FM_department") 
    origin = request("FM_origin") 

    if len(trim(request("FM_supplier"))) <> 0 then
    supplier = request("FM_supplier")
    else
    supplier = 0
    end if

    fastpris = request("FM_fastpris")

    collection = replace(request("FM_collection"), "'", "")
    composition = replace(request("FM_composition"), "'", "") 

    product_group = request("FM_product_group")

    scsq_note = replace(request("FM_scsq_note"), "'", "") 
    sample_note = replace(request("FM_sample_note"), "'", "") 

    job_internbesk = replace(request("FM_internnote"), "'", "") 

    beskrivelse = replace(request("FM_beskrivelse"), "'", "") 'invoice note

    rekvnr = replace(request("FM_rekvnr"), "'", "")

    if len(trim(request("FM_bruttooms"))) <> 0 then
    bruttooms = replace(request("FM_bruttooms"), ".","")
    bruttooms = replace(bruttooms, ",",".")
    else
    bruttooms = 0
    end if 

    if len(trim(request("FM_jo_udgifter_intern"))) <> 0 then
    jo_udgifter_intern = replace(request("FM_jo_udgifter_intern"), ".","")
    jo_udgifter_intern = replace(jo_udgifter_intern, ",",".")
    else
    jo_udgifter_intern = 0
    end if 


    
    if len(trim(request("FM_orderqty"))) <> 0 then 
    orderqty = request("FM_orderqty")
    else
    orderqty = 0
    end if
   
    
    if len(trim(request("FM_shippedqty"))) <> 0 then 
    shippedqty = request("FM_shippedqty")
    else
    shippedqty = 0
    end if

    supplier_invoiceno = request("FM_supplier_invoiceno") 

    transport = request("FM_transport")
    destination = request("FM_destination")


    kunde_betbetint = request("FM_betbetint") 
    kunde_levbetint = request("FM_levbetint")
    lev_betbetint =  request("FM_levbetbetint")
    lev_levbetint = request("FM_levlevbetint")
    'sup_betbetint = request("FM_betbetint")


    '**** ETD Buyer ***'
    dt_confb_etd = request("FM_dt_confb_etd")
    call fmatDate_fn(dt_confb_etd)
    dt_confb_etd = fmatDate

    
    '*** Bruges kun til Error required fileds, 
    '*** Dato sættes i orderAndProdDates func
    dt_actual_etd = request("FM_dt_actual_etd")
    call fmatDate_fn(dt_actual_etd)
    dt_actual_etd = fmatDate
    
    dt_confb_eta = request("FM_dt_confb_eta")
    call fmatDate_fn(dt_confb_eta)
    dt_confb_eta = fmatDate


    'Response.write "kopier_ordre: "& kopier_ordre & "<br>dt_actual_etd: " & dt_actual_etd & "<br>"


    '********* Requesred fields på ORDER jobstatus = 1 ***'


    if cint(jobstatus) <> 3 then 'inquery

         
        if cint(jobstatus) = 1 then 'active

            if cdbl(supplier) = 0 OR cdbl(orderqty) = 0 OR len(orderqty) = 0 OR dt_confb_etd = "2010-01-01" OR (cint(kunde_levbetint) = 2 AND dt_confb_eta  = "2010-01-01") then

            errortype = 173
            call showError(errortype)
				
            response.End

            end if

         end if



        if cint(jobstatus) = 2 then 'shipped

            if cint(supplier) = 0 OR cint(orderqty) = 0 OR len(orderqty) = 0 OR dt_confb_etd = "2010-01-01" OR cint(kpers) = 0 OR (dt_actual_etd = "2010-01-01" AND cint(kopier_ordre) <> 1) OR len(shippedqty) = 0 or shippedqty = 0 OR len(trim(supplier_invoiceno)) = 0 then

            errortype = 174
            call showError(errortype)
				
            response.End

         end if


        end if

    end if
    '*** Dato felter ****'
    
    '** enquery dates
    
    dt_jobstdato = request("FM_dt_jobstdato") 

    if isDate(dt_jobstdato) then
    dt_jobenddato = dateAdd("m", 4, dt_jobstdato)
    dt_jobenddato = year(dt_jobenddato)&"-"&month(dt_jobenddato)&"-"&day(dt_jobenddato)
    dt_jobstdato = year(dt_jobstdato)&"-"&month(dt_jobstdato)&"-"&day(dt_jobstdato)
    else
    dt_jobstdato = "2010-01-01"
    dt_jobenddato = "2010-01-01"
    end if

    

    dt_enq_st = request("FM_dt_enq_st")
    call fmatDate_fn(dt_enq_st)
    dt_enq_st = fmatDate

   

    dt_enq_end = request("FM_dt_enq_end")
    call fmatDate_fn(dt_enq_end)
    dt_enq_end = fmatDate


    dt_sour_dead = request("FM_dt_sour_dead")
    call fmatDate_fn(dt_sour_dead)
    dt_sour_dead = fmatDate
    
    
 
    dt_proto_dead = request("FM_dt_proto_dead")
    call fmatDate_fn(dt_proto_dead)
    dt_proto_dead = fmatDate
    
   
    
    
    dt_proto_sent = request("FM_dt_proto_sent")
    call fmatDate_fn(dt_proto_sent)
    dt_proto_sent = fmatDate
    
    

    dt_sms_dead = request("FM_dt_sms_dead")
    call fmatDate_fn(dt_sms_dead)
    dt_sms_dead = fmatDate

    dt_sms_sent = request("FM_dt_sms_sent")
    call fmatDate_fn(dt_sms_sent)
    dt_sms_sent = fmatDate

    
    dt_photo_dead = request("FM_dt_photo_dead")
    call fmatDate_fn(dt_photo_dead)
    dt_photo_dead = fmatDate

  
 
   dt_photo_sent = request("FM_dt_photo_sent")
    call fmatDate_fn(dt_photo_sent)
    dt_photo_sent = fmatDate
    
      dt_sup_photo_dead = request("FM_dt_sup_photo_dead")
    call fmatDate_fn(dt_sup_photo_dead)
    dt_sup_photo_dead = fmatDate

    dt_sup_sms_dead = request("FM_dt_sup_sms_dead")
    call fmatDate_fn(dt_sup_sms_dead)
    dt_sup_sms_dead = fmatDate


   call orderAndProdDates
     
    
    '*********** SLUT DATOER *************'


    lastid = 0


    if len(trim(request("FM_freight_pc"))) <> 0 then
    freight_pc = replace(request("FM_freight_pc"), ".","")
    freight_pc = replace(freight_pc, ",",".")
    else
    freight_pc = 0
    end if

    if len(trim(request("FM_tax_pc"))) <> 0 then
    tax_pc = replace(request("FM_tax_pc"), ".","")
    tax_pc = replace(tax_pc, ",",".") 
    else
    tax_pc = 0
    end if

    if len(trim(request("FM_comm_pc"))) <> 0  then
    comm_pc = replace(request("FM_comm_pc"), ".","")
    comm_pc = replace(comm_pc, ",",".")
    else
    comm_pc = 0
    end if


    

   
    if len(trim(request("FM_cost_price_pc"))) <> 0 then
    cost_price_pc = replace(request("FM_cost_price_pc"), ".","")
    cost_price_pc = replace(cost_price_pc, ",",".")
    else
    cost_price_pc = 0
    end if


    if len(trim(request("FM_cost_price_pc_base"))) <> 0 then
    cost_price_pc_base = replace(request("FM_cost_price_pc_base"), ".","")
    cost_price_pc_base = replace(cost_price_pc_base, ",",".")
    else
    cost_price_pc_base = 0
    end if


    
     if len(trim(request("FM_sales_price_pc"))) <> 0 then
    sales_price_pc = replace(request("FM_sales_price_pc"), ".","")
    sales_price_pc = replace(sales_price_pc, ",",".")
    else
    sales_price_pc = 0
    end if
       
     if len(trim(request("FM_tgt_price_pc"))) <> 0 then
     tgt_price_pc = replace(request("FM_tgt_price_pc"), ".","")
     tgt_price_pc = replace(tgt_price_pc, ",",".")
     else
     tgt_price_pc = 0
     end if

    if len(trim(request("FM_jo_dbproc"))) <> 0 then
    jo_dbproc = replace(request("FM_jo_dbproc"), ".","")
    jo_dbproc = replace(jo_dbproc, ",",".")
    else
    jo_dbproc = 0
    end if

    sales_price_pc_valuta = request("FM_valuta_sales_price_pc_valuta")
    cost_price_pc_valuta = request("FM_valuta_cost_price_pc_valuta") 
    tgt_price_pc_valuta = request("FM_valuta_tgt_price_pc_valuta") 

   

     

    if len(trim(request("FM_alert"))) <> 0 then
    alert = 1
    else
    alert = 0
    end if


    dd_dato = day(now) & ". "& left(monthname(month(now)), 3) &" "& year(now)
    editor = session("user")

    valuta = sales_price_pc_valuta

    'freight_price_pc_valuta, tgt_price_pc_valuta


    orderqty = replace(orderqty, ",", ".")


     jfak_moms = 0
     jfak_sprog = 2
     strSQLmomsK = "SELECT kfak_moms, kfak_sprog FROM kunder WHERE kid = "& kid 
     oRec.open strSQLmomsK, oConn, 3
     if not oRec.EOF then
    
        jfak_moms = oRec("kfak_moms")
        jfak_sprog = oRec("kfak_sprog") '2 UK

     end if
     oRec.close 

     


    if func = "dbopr" OR cint(kopier_ordre) = 1 then

        if cint(kopier_ordre) = 1 then

         call lastjobnr_fn()
        jobnr = nextjobnr

        end if


 

    strSQLjob = "INSERT INTO job (dato, editor, jobknr, jobnavn, jobstatus, jobnr, projektgruppe1, "_
    &" department, jobans1, origin, supplier, fastpris, collection, composition, product_group, scsq_note, sample_note, job_internbesk, beskrivelse, "_
    &" dt_enq_st, dt_enq_end, dt_sour_dead, dt_proto_dead, dt_proto_sent, dt_sms_dead, dt_sms_sent, dt_photo_dead, dt_photo_sent, dt_exp_order, "_
    &" dt_confb_etd, dt_confb_eta, dt_confs_etd, dt_confs_eta, dt_actual_etd, dt_actual_eta, rekvnr, jobstartdato, jobslutdato, "_
    &" dt_firstorderc, dt_ldapp, dt_sizeexp, dt_sizeapp, dt_ppexp, dt_ppapp, dt_shsexp, dt_shsapp, orderqty, shippedqty, supplier_invoiceno, transport, destination, jo_bruttooms, "_
    &" jo_udgifter_intern, dt_sup_photo_dead, dt_sup_sms_dead, freight_pc, tax_pc, comm_pc, cost_price_pc, sales_price_pc, tgt_price_pc, jo_dbproc, sales_price_pc_valuta, "_
    &" cost_price_pc_valuta, tgt_price_pc_valuta, cost_price_pc_base, kunde_betbetint, kunde_levbetint, lev_betbetint, lev_levbetint, valuta, jfak_moms, jfak_sprog, alert"_
    &" ) "_
    &" VALUES "_
    &" ('"& dd_dato &"', '"& editor &"', "& kid &", '"& jobnavn &"', "& jobstatus & ", '"& jobnr &"', 10, "_
    &"'"& department &"', "& kpers &", '"& origin &"', "& supplier &", "& fastpris &", '"& collection &"', '"& composition &"', "& product_group &", "_
    &"'"& scsq_note &"', '"& sample_note &"', '"& job_internbesk &"', '"& beskrivelse &"', '"& dt_enq_st &"', '"& dt_enq_end &"', '"& dt_sour_dead &"', "_
    &"'"& dt_proto_dead  &"', '"& dt_proto_sent  &"', '"& dt_sms_dead  &"', '"& dt_sms_sent &"', '"& dt_photo_dead  &"', '"& dt_photo_sent  &"', '"& dt_exp_order &"', "_
    &"'"& dt_confb_etd &"', '"& dt_confb_eta &"','"& dt_confs_etd &"', '"& dt_confs_eta &"', '"& dt_actual_etd &"', '"& dt_actual_eta &"', '"& rekvnr &"', '"& dt_jobstdato &"', '"& dt_jobenddato &"', "_
    &"'"& dt_firstorderc &"', '"& dt_ldapp &"', '"& dt_sizeexp &"', '"& dt_sizeapp &"', '"& dt_ppexp &"', '"& dt_ppapp &"', '"& dt_shsexp &"', '"& dt_shsapp &"', "_
    &""& orderqty &","& shippedqty &",'"& supplier_invoiceno &"', '"& transport &"', '"& destination &"', "& bruttooms &", "_
    &" "& jo_udgifter_intern &", '"& dt_sup_photo_dead &"', '"& dt_sup_sms_dead &"', "_
    &" "& freight_pc &","& tax_pc &","& comm_pc &", "& cost_price_pc &","& sales_price_pc &","& tgt_price_pc &", "& jo_dbproc &", "& sales_price_pc_valuta &","_
    &" "& cost_price_pc_valuta &", "& tgt_price_pc_valuta &", "& cost_price_pc_base &", "& kunde_betbetint &","& kunde_levbetint &","& lev_betbetint &","& lev_levbetint &", "& valuta &", "& jfak_moms &","& jfak_sprog &", "& alert &""_
    &" )" 


    'response.write strSQLjob
    'response.write "<br>dt_actual_etd:"& dt_actual_etd
    'response.end 

    oConn.execute(strSQLjob)


     

       strSQL = "SELECT id FROM job WHERE id <> 0 ORDER by id DESC limit 1"
       oRec.open strSQL, oConn, 3
       if not oRec.EOF then
        
        lastid = oRec("id")

       end if
       oRec.close

                
                
          
			
        

      '*** Opdater jobnr rækkefælge ***'
      strSQL = "UPDATE licens SET jobnr = "& jobnr &" WHERE id = 1"
	  oConn.execute(strSQL)

    else

     strSQLjob = "UPDATE job SET dato = '"& dd_dato &"', editor = '"& editor &"', jobknr = "& kid &", jobnavn = '"& jobnavn &"', "_
     &" jobstatus  = "& jobstatus & ", jobnr = '"& jobnr &"', projektgruppe1 = 10, "_
     &" department = '"& department &"', jobans1 = "& kpers &", origin = '"& origin &"', supplier = "& supplier &", "_
     &" fastpris = "& fastpris &", collection = '"& collection &"', composition = '"& composition &"', product_group = "& product_group &", "_
     &" scsq_note = '"& scsq_note &"', sample_note = '"& sample_note &"', job_internbesk = '"& job_internbesk &"', beskrivelse = '"& beskrivelse &"', "_
     &" dt_enq_st = '"& dt_enq_st &"', dt_enq_end = '"& dt_enq_end &"', dt_sour_dead = '"& dt_sour_dead &"', "_
     &" dt_proto_dead = '"& dt_proto_dead &"', dt_proto_sent = '"& dt_proto_sent &"', dt_sms_dead = '"& dt_sms_dead &"', dt_sms_sent = '"& dt_sms_sent &"', dt_photo_dead = '"& dt_photo_dead &"',"_
     &" dt_photo_sent = '"& dt_photo_sent &"', dt_exp_order = '"& dt_exp_order & "', "_
     &" dt_confb_etd = '"& dt_confb_etd &"', dt_confb_eta = '"& dt_confb_eta &"', dt_confs_etd = '"& dt_confs_etd &"', "_
     &" dt_confs_eta = '"& dt_confs_eta &"', dt_actual_etd = '"& dt_actual_etd &"', dt_actual_eta = '"& dt_actual_eta &"', rekvnr = '"& rekvnr &"', "_
     &" jobstartdato = '"& dt_jobstdato &"', jobslutdato = '"& dt_jobenddato &"', "_
     &" dt_firstorderc = '"& dt_firstorderc &"', dt_ldapp = '"& dt_ldapp &"', dt_sizeexp = '"& dt_sizeexp &"', dt_sizeapp = '"& dt_sizeapp &"', "_
     &" dt_ppexp = '"& dt_ppexp &"', dt_ppapp = '"& dt_ppapp &"', dt_shsexp = '"& dt_shsexp &"', dt_shsapp = '"& dt_shsapp &"', "_
     &" orderqty = "& orderqty &", shippedqty = "& shippedqty &", supplier_invoiceno = '"& supplier_invoiceno &"', transport = '"& transport &"', "_
     &" destination = '"& destination &"', jo_bruttooms = "& bruttooms &", jo_udgifter_intern = "& jo_udgifter_intern &", "_
     &" dt_sup_photo_dead = '"& dt_sup_photo_dead &"', dt_sup_sms_dead = '"& dt_sup_sms_dead &"', "_
     &" freight_pc = "& freight_pc &", tax_pc = "& tax_pc &", comm_pc = "& comm_pc &", "_
     &" cost_price_pc = "& cost_price_pc &", sales_price_pc = "& sales_price_pc &", tgt_price_pc = "& tgt_price_pc &", jo_dbproc = "& jo_dbproc &", "_
     &" sales_price_pc_valuta = "& sales_price_pc_valuta &", cost_price_pc_valuta = "& cost_price_pc_valuta &", tgt_price_pc_valuta = "& tgt_price_pc_valuta &", "_
     &" cost_price_pc_base = "& cost_price_pc_base &", "_
     &" kunde_betbetint = "& kunde_betbetint &", kunde_levbetint = "& kunde_levbetint &", lev_betbetint = "& lev_betbetint &", lev_levbetint = "& lev_levbetint &", valuta = "& valuta &", jfak_moms= "& jfak_moms &", jfak_sprog = "& jfak_sprog &", alert = "& alert &""_
     &" WHERE id = " & jobid
    
    'response.write strSQLjob
    'response.flush 
    
    oConn.execute(strSQLjob)

    lastid = jobid

    end if


    response.redirect "job_nt.asp?func=table&lastid="&lastid

case "bulk"


    dd_dato = day(now) & ". "& left(monthname(month(now)), 3) &" "& year(now)
    editor = session("user")

    if len(trim(request("bulk_jobids"))) > 3 then
    bulk_jobids = request("bulk_jobids")
    else
    bulk_jobids = 0
    end if



    bulk_jobidsArr = split(bulk_jobids, ",")


    for j = 0 TO UBOUND(bulk_jobidsArr)

     
    strSQLjobbulk = "UPDATE job SET dato = '"& dd_dato &"', editor = '"& editor &"'"

   

    '** Enq

    call enqDates

    if dt_sms_dead <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_sms_dead = '"& dt_sms_dead & "'"
    end if

    if dt_sup_sms_dead <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_sup_sms_dead = '"& dt_sup_sms_dead & "'"
    end if

     if dt_sms_sent <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_sms_sent = '"& dt_sms_sent & "'"
    end if

     if dt_photo_dead <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_photo_dead = '"& dt_photo_dead & "'"
    end if

      if dt_sup_photo_dead <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_sup_photo_dead = '"& dt_sup_photo_dead & "'"
    end if

   if dt_photo_sent <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_photo_sent = '"& dt_photo_sent & "'"
    end if
    
    
    '** Order 
   
      call orderAndProdDates

    if dt_exp_order <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_exp_order = '"& dt_exp_order & "'"
    end if


     if dt_confb_etd  <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_confb_etd  = '"& dt_confb_etd  & "'"
    end if

    if dt_confb_eta <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_confb_eta = '"& dt_confb_eta & "'"
    end if
    
    if dt_confs_etd <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_confs_etd = '"& dt_confs_etd & "'"
    end if

    if dt_confs_eta <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_confs_eta = '"& dt_confs_eta & "'"
    end if


    if dt_actual_etd <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_actual_etd = '"& dt_actual_etd & "'"
    end if

      if dt_actual_eta <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_actual_eta = '"& dt_actual_eta & "'"
    end if




     if dt_firstorderc <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_firstorderc = '"& dt_firstorderc & "'"
    end if


     if dt_firstorderc <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_firstorderc = '"& dt_firstorderc & "'"
    end if

    if dt_ldapp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_ldapp = '"& dt_ldapp & "'"
    end if

    if dt_sizeexp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_sizeexp = '"& dt_sizeexp & "'"
    end if

     if dt_sizeapp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_sizeapp = '"& dt_sizeapp & "'"
    end if

    if dt_ppexp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_ppexp = '"& dt_ppexp & "'"
    end if

    if dt_ppexp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_ppexp = '"& dt_ppexp & "'"
    end if

     if dt_ppapp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_ppapp = '"& dt_ppapp & "'"
    end if

    if dt_shsexp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_shsexp = '"& dt_shsexp & "'"
    end if

     if dt_shsapp <> "2010-01-01" then
    strSQLjobbulk = strSQLjobbulk & ", dt_shsapp = '"& dt_shsapp & "'"
    end if

   


   
    strSQLjobbulk = strSQLjobbulk &" WHERE id = " & bulk_jobidsArr(j)

   
    'response.write strSQLjob
    'response.flush 
    
    oConn.execute(strSQLjobbulk)


    next


    lastid = jobid

    response.redirect "job_nt.asp?func=table&lastid="&lastid


case "opret", "red"






if func = "red" then
    dbfunc = "dbred"

    if len(trim(request("jobid"))) <> 0 then
    id = request("jobid")
    else
    id = 0
    end if


    strSQLjob = "SELECT jobnavn, jobnr, jobknr, jobstatus,"_
    &" department, jobans1, origin, supplier, fastpris, collection, composition, product_group, "_
    &" scsq_note, sample_note, job_internbesk, beskrivelse, dt_enq_st, dt_enq_end, "_
    &" dt_proto_dead, dt_proto_sent, dt_sms_dead, dt_sms_sent, dt_photo_dead, dt_photo_sent, dt_exp_order, dt_sour_dead, "_ 
    &" dt_confb_etd, dt_confb_eta, dt_confs_etd, dt_confs_eta, dt_actual_etd, dt_actual_eta, rekvnr, jobstartdato, "_
    &" dt_firstorderc, dt_ldapp, dt_sizeexp, dt_sizeapp, dt_ppexp, dt_ppapp, dt_shsexp, dt_shsapp, orderqty, "_
    &" shippedqty, supplier_invoiceno, transport, destination, jo_bruttooms, jo_udgifter_intern, dt_sup_photo_dead, dt_sup_sms_dead, "_
    &" freight_pc, tax_pc, comm_pc, cost_price_pc, sales_price_pc, tgt_price_pc, jo_dbproc, sales_price_pc_valuta, cost_price_pc_valuta, tgt_price_pc_valuta, cost_price_pc_base, "_
    &" kunde_betbetint, kunde_levbetint, lev_betbetint, lev_levbetint, alert "_
    &" FROM job WHERE id = "& id

    'response.write strSQLjob
    'response.flush

    oRec.open strSQLjob, oConn, 3
    if not oRec.EOF then

    jobnavn = oRec("jobnavn")
    jobnr = oRec("jobnr")
    jobknr = oRec("jobknr")
    jobstatus = oRec("jobstatus")
    department = oRec("department")
    kpers = oRec("jobans1")
    origin = oRec("origin")
    supplier = oRec("supplier")
    fastpris = oRec("fastpris")
    collection = oRec("collection")
    composition = oRec("composition")
    product_group = oRec("product_group")
    scsq_note = oRec("scsq_note") 
    sample_note = oRec("sample_note")
    job_internbesk = oRec("job_internbesk") 
    beskrivelse = oRec("beskrivelse") 'inv. note
    
    dt_enq_st = oRec("dt_enq_st")
    dt_enq_end = oRec("dt_enq_end")

    dt_sour_dead = oRec("dt_sour_dead")

    dt_proto_dead = oRec("dt_proto_dead")
    dt_proto_sent = oRec("dt_proto_sent")
    dt_sms_dead = oRec("dt_sms_dead")
    dt_sms_sent = oRec("dt_sms_sent")
    dt_photo_dead = oRec("dt_photo_dead") 
    dt_photo_sent = oRec("dt_photo_sent")
    dt_exp_order = oRec("dt_exp_order")


    dt_confb_etd = oRec("dt_confb_etd") 
    dt_confb_eta = oRec("dt_confb_eta")
    dt_confs_etd = oRec("dt_confs_etd")
    dt_confs_eta = oRec("dt_confs_eta")
    dt_actual_etd = oRec("dt_actual_etd")
    dt_actual_eta = oRec("dt_actual_eta")


    rekvnr = oRec("rekvnr")

    dt_jobstdato = oRec("jobstartdato")


    dt_firstorderc = oRec("dt_firstorderc") 
    dt_ldapp = oRec("dt_ldapp") 
    dt_sizeexp = oRec("dt_sizeexp") 
    dt_sizeapp = oRec("dt_sizeapp")
    dt_ppexp = oRec("dt_ppexp")
    dt_ppapp = oRec("dt_ppapp")
    dt_shsexp = oRec("dt_shsexp")
    dt_shsapp = oRec("dt_shsapp")

    dt_sup_photo_dead = oRec("dt_sup_photo_dead")
    dt_sup_sms_dead = oRec("dt_sup_sms_dead")



    orderqty = oRec("orderqty")
    shippedqty = oRec("shippedqty")
    if shippedqty = 0 then
    shippedqty = ""
    end if
    
    supplier_invoiceno = oRec("supplier_invoiceno")


    transport = oRec("transport")
    destination = oRec("destination")

    bruttooms = oRec("jo_bruttooms")
    jo_udgifter_intern = oRec("jo_udgifter_intern")



    freight_pc = oRec("freight_pc")
    tax_pc = oRec("tax_pc") 
    comm_pc = oRec("comm_pc")

    cost_price_pc = oRec("cost_price_pc")
    cost_price_pc_base = oRec("cost_price_pc_base")

    sales_price_pc = oRec("sales_price_pc") 
    tgt_price_pc = oRec("tgt_price_pc") 

    jo_dbproc = oRec("jo_dbproc")

    sales_price_pc_valuta = oRec("sales_price_pc_valuta")
    cost_price_pc_valuta = oRec("cost_price_pc_valuta")

    tgt_price_pc_valuta = oRec("tgt_price_pc_valuta")


    kunde_betbetint = oRec("kunde_betbetint")
    kunde_levbetint = oRec("kunde_levbetint")
    lev_betbetint = oRec("lev_betbetint")
    lev_levbetint = oRec("lev_levbetint")

    alert = oRec("alert")

    if cint(alert) = 1 then
    alertCHK = "CHECKED"
    else
    alertCHK = ""
    end if


    end if
    oRec.close


    call alfanumerisk(jobnr)
    jobnr = alfanumeriskTxt
    jobnr = left(jobnr,20)



else

    dbfunc = "dbopr"
    id = 0

    call lastjobnr_fn()
    jobnr = nextjobnr


    '*** Opdater jobnr rækkefælge ***'
    strSQL = "UPDATE licens SET jobnr = "& jobnr &" WHERE id = 1"
	oConn.execute(strSQL)


    jobknr = 0
    jobnavn = ""

    jobstatus = 3
    kpers = 0
    origin = ""
    supplier = 0

    fastpris = 3

    collection = ""
    composition = ""

    product_group = 0

    sample_note = ""
    scsq_note = ""
    job_internbesk = ""
    beskrivelse = ""
    dt_enq_st = ""
    dt_enq_end = ""

     dt_sour_dead = ""

     dt_proto_dead = ""
    dt_proto_sent = ""
    dt_sms_dead = ""
    dt_sms_sent = ""
    dt_photo_dead = ""
    dt_photo_sent = ""
    dt_exp_order = ""

      dt_confb_etd = ""
    dt_confb_eta = ""
    dt_confs_etd = ""
    dt_confs_eta = ""
    dt_actual_etd = ""
    dt_actual_eta = ""

    rekvnr = ""

    dt_jobstdato = formatdatetime(now, 2)



     dt_firstorderc = ""
    dt_ldapp = "" 
    dt_sizeexp = "" 
    dt_sizeapp = ""
    dt_ppexp = ""
    dt_ppapp = ""
    dt_shsexp = ""
    dt_shsapp = ""
    dt_sup_photo_dead = ""
    dt_sup_sms_dead = ""


    orderqty = 0
    shippedqty = ""
    supplier_invoiceno = ""
    

    transport = ""
    destination = "Danmark"

    bruttooms = 0
    jo_udgifter_intern = 0

    cost_price_pc = 0
    cost_price_pc_base = 0
    sales_price_pc = 0
    tgt_price_pc = 0

    freight_pc = 0
    tax_pc = 0
    comm_pc = 0

    jo_dbproc = 0

    sales_price_pc_valuta = 3
    cost_price_pc_valuta = 3
    tgt_price_pc_valuta = 3


    kunde_betbetint = 0
    kunde_levbetint = 0
    lev_betbetint = 0
    lev_levbetint = 0

    alert = 0
    alertCHK = ""

    '*** lev & betbet kunde 
    'strSQLlevbetkunde = "SELECT levbet, betbet, betbetint FROM kunder WHERE kid = " & jobknr
    'oRec.open strSQLlevbetkunde, oConn, 3
    'if not oRec.EOF then

    ''kundeLevBet = oRec("levbet")
    'kunde_betbetint = oRec("betbetint")
    ''kundeBetBet = oRec("betbet")

    'end if
    'oRec.close


    
    '*** lev & betbet supplier 
    'strSQLlevbetkunde = "SELECT levbet, betbet, betbetint FROM kunder WHERE kid = " & supplier
    'oRec.open strSQLlevbetkunde, oConn, 3
    'if not oRec.EOF then

    ''supLevBet = oRec("levbet")
    'lev_betbetint = oRec("betbetint")

    'end if
    'oRec.close
  

end if 'Opret / rediger


   

    




%>

<!--#include file="inc/job_nt_inc.asp"-->



      <div class="portlet">
        <h3 class="portlet-title">
          <u>Order - Edit</u>
        </h3>
        <div class="portlet-body">




          <!-- Opret - Rediger Ordre form --->
           
             <form action="job_nt.asp?func=<%=dbfunc %>" method="post">
              <input type="hidden" name="FM_jobid" value="<%=id %>" />
               <input type="hidden" id="showfullscreen" value="0" />
              <%call valutaKurs(1) %>
            <input type="hidden" id="valuta_kurs_1" value="<%=dblKurs %>" /> 
            <%call valutaKurs(2) %>
            <input type="hidden" id="valuta_kurs_2" value="<%=dblKurs %>" />
            <%call valutaKurs(3) %>
            <input type="hidden" id="valuta_kurs_3" value="<%=dblKurs %>" />
            <%call valutaKurs(4) %>
            <input type="hidden" id="valuta_kurs_4" value="<%=dblKurs %>" />
            <%call valutaKurs(5) %>
            <input type="hidden" id="valuta_kurs_5" value="<%=dblKurs %>" />



             <section>
                <div class="well well-white">
                     <!--BASEDATA START -->
                         <div class="row">
                        <div class="col-lg-12">
                            <h4 class="panel-title-well">Basedata</h4>
                        </div>
                    </div>

                       

                         <div class="row">
                       
                               <div class="col-lg-6">
                                    Buyer <span style="color:red;">*</span>
                                   <select class="form-control input-small"id="FM_kunde" name="FM_kunde" class="form-control input-small"><%=strFil_Kon_Txt %></select>

                                </div>

                                  <div class="col-lg-2">
                                    Order No. <span style="color:red;">*</span>
                                   <input class="form-control input-small" type="text" name="" value="<%=jobnr %>" DISABLED />
                                   <input class="form-control input-small"type="hidden" id="FM_jobnr" name="FM_jobnr" value="<%=jobnr %>" />

                                </div>
                        
                       
                               <div class="col-lg-4 pad-r30">&nbsp;
                                  
                                </div>
                         </div>


                   
                    <div class="row">
                       
                             <div class="col-lg-4 pad-t10">
                             Place of origin
                             <select class="form-control input-small"id="FM_origin" name="FM_origin"><option value="0">Choose..</option>
                                <option value="Turkey" <%=originTurSel %>>Turkey</option>
                                <option value="China" <%=originChiSel %>>China</option>
                                 <option value="India" <%=originIndSel %>>India</option>
                                <option value="Bangladesh" <%=originBanSel %>>Bangladesh</option>
                                <option value="Vietnam" <%=originVieSel %>>Vietnam</option>
                                 <option value="Hongkong" <%=originHonSel %>>Hongkong</option>
                            </select>
                        </div>


                        
                        <div class="col-lg-4 pad-t10">
                            Supplier <span style="color:red;">(*)</span>
                            <select class="form-control input-small"id="FM_supplier" name="FM_supplier"><%=strFil_Sup_Txt %></select>
                        </div>

                             <div class="col-lg-2 pad-t10">
                                     Department
                                    <select class="form-control input-small"id="xFM_afd" name="FM_department"><option value="0">Choose..</option>
                                        <option value="Men" <%=departmentMenSEL %>>Men</option>
                                        <option value="Women" <%=departmentWomenSEL %>>Women</option>
                                        <option value="Kids" <%=departmentKidsSEL %>>Kids</option>
                                        <option value="Unisex" <%=departmentUnisexSEL %>>Unisex</option>
                                    </select>

                                </div>

                          <div class="col-lg-2 pad-t10 pad-r30">
                              &nbsp;
                              </div>

                    </div>

                    <div class="row">

                         <div class="col-lg-4 pad-t10">
                            Sales Rep. <span style="color:red;">(*)</span>

                            <%call salesreplist(kpers) %>

                             <select class="form-control input-small"id="FM_kpers" name="FM_kpers"><%=strFil_Kpers_Txt %></select>
                        </div>
                       
                         <div class="col-lg-2 pad-t10">
                            Status
                            <select class="form-control input-small"name="FM_status" id="status">
                            <option value="3" <%=jobstatus3SEL %>>Enquiry</option>
                            <option value="1" <%=jobstatus1SEL %>>Active</option>
                            
                            <option value="2" <%=jobstatus2SEL %>>Shipped</option>
                            <option value="4" <%=jobstatus4SEL %>>Cancelled</option>
                            <option value="0" <%=jobstatus0SEL %>>Invoiced/Closed</option>
                           
                          </select>
                        </div>

                        <div class="col-lg-4 pad-t10">
                            Order type
                            <select class="form-control input-small"name="FM_fastpris" id="fastpris">
                                <option value="2" <%=fastpris2SEL %>>Commission</option>
                                <option value="3" <%=fastpris3SEL %>>Salesorder</option>
                               <option value="1" disabled>Fixed price</option>
                                <option value="0" disabled>Time & Mat.</option>
                            </select>
                        </div>
                       
                         <div class="col-lg-2 pad-t10">
                        <%if func = "red" then %>
                        <br /><input type="checkbox" name="FM_kopierordre" value="1" /> Copy Order
                         <%else %>
                             &nbsp;

                          <%end if %>
                               </div>

                      
                     </div>


                        <div class="row">
                                <div class="col-lg-11 pad-t10">
                                    &nbsp;
                                </div>
                         <div class="col-lg-2 pad-t10 pad-r30 pull-right">
                        <input class="btn btn-success" type="submit" value="Submit >>">
                               </div>

                          </div>
                    

                  </div><!-- well well-white -->
            </section>
                  <!--BASEDATA END -->


            

            <br />
                          
                <div class="panel-group accordion-panel" id="accordion-paneled0">
                    <!-- ENQUIRY INFO TOP -->
                        <div class="panel panel-default"> 
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseOne">Enquiry info</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseOne" class="panel-collapse collapse">
                       
                            <div class="panel-body">

                            <div class="row">
    
                            <div class="col-lg-9">

                            <table class="tablecolor" style="width:100%">

                                <tr>
                                    <td colspan="2">Style <span style="color:red;">*</span> <br />
                                        <input class="form-control input-small" type="text" name="FM_jobnavn" value="<%=jobnavn %>" />
                                    </td>

                                    <td>Collection <br />
                                        <input class="form-control input-small" type="text" name="FM_collection" value="<%=collection %>" />
                                    </td>
                                                                                                        
                               </tr>
                                <tr>
                                    <td>Enquiry startup <br />
                                        <div class='input-group date'>
                                            <input class="form-control input-small" type="text" name="FM_dt_enq_st" value="<%=dt_enq_st %>" placeholder="dd-mm-yyyy" />
                                             <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                        </div>
                                    </td>
                                    <td>Enquiry end
                                        <div class='input-group date'>
                                            <input class="form-control input-small" type="text" name="FM_dt_enq_end" value="<%=dt_enq_end %>" placeholder="dd-mm-yyyy" />
                                            <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar">
                                            </span>
                                            </span>
                                        </div>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2">Product group <br />
                                        <select class="form-control input-small" name="FM_product_group"><%=strFil_PG_Txt %></select>
                                    </td>
                                    <td>
                                        Composition <br />
                                        <input class="form-control input-small" type="text" name="FM_composition" value="<%=composition %>" />
                                    </td>
                                </tr>

                                <tr>
                                    <td>
                                        Sourcing deadline <br />
                                        <div class='input-group date'>
                                                <input class="form-control input-small" type="text" name="FM_dt_sour_dead" value="<%=dt_sour_dead %>" placeholder="dd-mm-yyyy"/>
                                                <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                        </div>
                                    </td>
                                    <td>
                                        Proto deadline <br />
                                        <div class='input-group date'>
                                                <input class="form-control input-small" type="text" name="FM_dt_proto_dead" value="<%=dt_proto_dead %>" placeholder="dd-mm-yyyy" />
                                                <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                        </div>
                                    </td>
                                    <td>
                                        Proto sent
                                        <div class='input-group date'>
                                            <input class="form-control input-small" type="text" name="FM_dt_proto_sent" value="<%=dt_proto_sent %>" placeholder="dd-mm-yyyy" />
                                            <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar">
                                            </span>
                                            </span>
                                      </div>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        SMS Buyer deadline <br />                         
                                       <div class='input-group date'>
                                               <input class="form-control input-small" type="text" name="FM_dt_sms_dead" value="<%=dt_sms_dead %>" placeholder="dd-mm-yyyy" />
                                             <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                      </div>
                                    </td>
                                    <td>
                                        SMS Supplier DL
                              
                                      <div class='input-group date'>
                                           <input class="form-control input-small" type="text" name="FM_dt_sup_sms_dead" value="<%=dt_sup_sms_dead%>" placeholder="dd-mm-yyyy" />
                                         <span class="input-group-addon input-small">
                                            <span class="fa fa-calendar">
                                            </span>
                                        </span>
                                      </div>
                                    </td>
                                    <td>
                                        SMS sent  <br />                             
                                        <div class='input-group date'>
                                                <input class="form-control input-small" type="text" name="FM_dt_sms_sent" value="<%=dt_sms_sent %>" placeholder="dd-mm-yyyy" />
                                                <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                        </div>
                                    </td>
                                </tr>

                                <tr>
                                    <td>
                                        Photo Buyer Deadline <br />
                               
                                        <div class='input-group date'>
                                            <input class="form-control input-small" type="text" name="FM_dt_photo_dead" value="<%=dt_photo_dead %>" placeholder="dd-mm-yyyy" />
                                             <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                      </div>
                                    </td>
                                    <td>
                                        Photo Supplier DL <br />
                                
                                          <div class='input-group date'>
                                           <input class="form-control input-small" type="text" name="FM_dt_sup_photo_dead" value="<%=dt_sup_photo_dead%>" placeholder="dd-mm-yyyy" />
                                             <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                          </div>
                                    </td>
                                    <td>
                                        Photo sent
                                            <div class='input-group date'>
                                            <input class="form-control input-small" type="text" name="FM_dt_photo_sent" value="<%=dt_photo_sent %>" placeholder="dd-mm-yyyy" />
                                                <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                          </div>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Exp order <br />
                                       <div class='input-group date'>
                                          <input class="form-control input-small" type="text" name="FM_dt_exp_order" value="<%=dt_exp_order %>" placeholder="dd-mm-yyyy"/>
                                             <span class="input-group-addon input-small">
                                                <span class="fa fa-calendar">
                                                </span>
                                            </span>
                                      </div>
                                    </td>
                                </tr>
                            </table>
                            </div>

                        
                            <div class="col-lg-3">
                                <table class="tablecolor">
                                    <%
	                                strSQL = "SELECT id, filnavn FROM filer WHERE filertxt = "& jobnr
                                    
	                                oRec.open strSQL, oConn, 3
	                                j = 0
	                                if not oRec.EOF then
                                    if len(trim(oRec("filnavn"))) <> 0 then                       
	                                %>
                                    <tr>
                                        <td>
                                            Picture <br />
                                            <div class="fileinput fileinput-new" data-provides="fileinput">
                                                <div class="fileinput-preview thumbnail" style="width: 250px; height: 300px;">
                                                    <img src="../inc/upload/<%=lto%>/<%=oRec("filnavn")%>" alt='' border='0'>                                                   
                                                </div>
                                            </div>
                                            <a href="job_nt.asp?func=sletfil&id=<%=oRec("id")%>&filnavn=<%=oRec("filnavn")%>" class="btn btn-default btn-sm">Remove image</a>
                                        </td>                                        
                                    </tr>
                                    <%
                                    j = j + 1                                    
	                                end if
                                    end if	                                
	                                oRec.close 

                                    if j = 0 then
                                    %>
                                    <tr>
                                        <td>
                                            Picture <br />
                                            <div class="fileinput fileinput-new" data-provides="fileinput">
                                                <div class="fileinput-new thumbnail" style="width: 250px; height: 300px;" id="nt_file">                                                  
                                                </div>
                                            </div>
                                            <a onclick="Javascript:window.open('upload.asp?menu=fob&func=opret&jobnr=<%=jobnr%>', '', 'width=650,height=600,resizable=yes,scrollbars=yes')" class="btn btn-default btn-sm">Add image</a>
                                            <span id="sp_updatepic" style="color:#5582d2; float:right; padding-top:5px; font-size:125%; color:dimgrey" class="fa fa-refresh"></span>
                                        </td>
                                    </tr>
                                    <%end if %>
                                </table>
                            </div>
                            
                               
                            </div>
                            

                            <div class="row">
                               <div class="col-lg-6 pad-t10">
                                Sample colour &amp Sample Qty.
                                <textarea name="FM_scsq_note" class="form-control input-small" placeholder="Write your comment here"><%=scsq_note %></textarea>
                            </div>

                            <div class="col-lg-6 pad-t10 pad-r30">
                                Sample Note
                                <textarea name="FM_sample_note" class="form-control input-small" placeholder="Write your comment here"><%=sample_note %></textarea>
                            </div>
                         </div>
                           

            </div>
            </div>
            </div>
            </div>
          <!--ENQUIRY END - Top -->



            <!--PRICE TERMS TOP START -->
               <div class="panel-group accordion-panel" id="accordion-paneled3">
                    <!-- ENQUIRY INFO TOP -->
                        <div class="panel panel-default"> 
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTwo">Price terms</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseTwo" class="panel-collapse collapse">
                       
                            <div class="panel-body">



             <!-- PRICE TERMS START -->
             <section>
             <!--<div class="well well-white">-->



                       <div class="row">
                               <div class="col-lg-2 pad-t10">
                                Order Qty. <span style="color:red;">(*)</span>
                                <input class="form-control input-small" type="text" name="FM_orderqty" id="orderqty" value="<%=orderqty%>" />
                            </div>
                             <div class="col-lg-10 pad-t10">&nbsp;</div>

                         </div> 

                         <div class="row">
                            <div class="col-lg-12 pad-t10 pad-r30">
                                 Price, Invoice or Production Note
                                <textarea class="form-control input-small" style="height:150px;" name="FM_internnote" placeholder="Write your comment here"><%=job_internbesk %></textarea>
                                <br /> <input type="checkbox" name="FM_alert" value="1" <%=alertCHK %> /> Alert (attention needed, add comment to production note)
                            </div>
                        </div> 

                        <div class="row">
                            <div class="col-lg-4 pad-t10">
                            Cost price PC (<label id="cost_price_pc_label"><%=formatnumber(cost_price_pc_base, 4) %></label>) base amount here
                          
                                <input class="form-control input-small" type="text" name="FM_cost_price_pc" id="cost_price_pc" value="<%=formatnumber(cost_price_pc, 4) %>" />
                                <input class="form-control input-small"type="hidden" name="FM_cost_price_pc_base" id="cost_price_pc_base" value="<%=formatnumber(cost_price_pc_base, 4) %>" />
                             </div>
                              <div class="col-lg-2"><br /><br />
                                <%call valutakoder("cost_price_pc_valuta", cost_price_pc_valuta, 1) %>
                                   </div>
                              <div class="col-lg-6 pad-t10">&nbsp;</div>


                         </div> 
                    

                       <div class="row">

                              <div class="col-lg-4 pad-t10">
                                Sales price PC <span style="color:#999999;">(invoice currency)</span>
                                    <input class="form-control input-small" type="text" name="" id="sales_price_pc_label" value="<%=formatnumber(sales_price_pc, 4) %>" DISABLED/>
                                    <input class="form-control input-small" type="text" name="FM_sales_price_pc" id="sales_price_pc" value="<%=formatnumber(sales_price_pc, 4) %>"/>
                              </div>

                               <div class="col-lg-2 pad-t10"><br />
                                    <%call valutakoder("sales_price_pc_valuta", sales_price_pc_valuta, 1) %>
                              </div>

                            <div class="col-lg-4 pad-t10"><br /><input type="checkbox" id="cc" disabled /> Add sales price manually (and calc. comi.)</div>

                            <div class="col-lg-2 pad-t10">&nbsp;</div>
                        </div> 

                        <div class="row">
                          <div class="col-lg-4 pad-t10">Target price PC 
                              <input class="form-control input-small" type="text" name="FM_tgt_price_pc" id="tgt_price_pc" value="<%=formatnumber(tgt_price_pc, 4) %>"/>
                            </div>
                             <div class="col-lg-2 pad-t10"><br />
                                <%
                                    
                                  call valutakoder("tgt_price_pc_valuta", tgt_price_pc_valuta, 1) %>
                               </div>
                                <div class="col-lg-6 pad-t10">&nbsp;</div>
                         </div>

                       
                      <div class="row">
                          <div class="col-lg-4 pad-t10">Comission PC %
                            
                                <input class="form-control input-small" type="text" name="" id="comm_pc_label" value="<%=comm_pc%>" DISABLED/>
                                <input class="form-control input-small" type="text" name="FM_comm_pc" id="comm_pc" value="<%=comm_pc %>" />
                          </div>
                           <div class="col-lg-8 pad-t10">&nbsp;</div>
                       </div>
                     
                          
                      <div class="row">
                          <div class="col-lg-4 pad-t10">Freight PC
                          
                                 <input class="form-control input-small" type="text" name="" id="freight_pc_label" value="0" DISABLED/>
                            <input class="form-control input-small" type="text" name="FM_freight_pc" id="freight_pc" value="<%=freight_pc %>" />
                            </div>
                            <div class="col-lg-2 pad-t10"><br />

                                 <%freight_price_pc_valuta = cost_price_pc_valuta 'følger altid %>
                             <%call valutakoder("freight_price_pc_valuta", freight_price_pc_valuta, 1) %>
                                </div>
                                  <div class="col-lg-8 pad-t10">&nbsp;</div>
                               
                        </div>


                       <div class="row">
                          <div class="col-lg-4 pad-t10">TAX % 
                          
                                <input class="form-control input-small" type="text" name="" id="tax_pc_label" value="<%=tax_pc %>" DISABLED/>
                                <input class="form-control input-small" type="text" name="FM_tax_pc" id="tax_pc" value="<%=tax_pc %>" />
                            </div>
                             <div class="col-lg-8 pad-t10">&nbsp;</div>


                       </div>


                         <div class="row">
                          <div class="col-lg-3 pad-t10">Total cost price DKK <span style="color:#999999;">(Order Qty. * Cost. PC * TAX PC) + (Freight PC * Order Qty.)</span>
                          
                                <input class="form-control input-small" type="text" id="udgifter_intern_label" value="<%=formatnumber(jo_udgifter_intern, 2) %>" DISABLED/>
                                <input class="form-control input-small"type="hidden" name="FM_jo_udgifter_intern" id="udgifter_intern" value="<%=formatnumber(jo_udgifter_intern, 2) %>"/>
                           </div>

                          <div class="col-lg-3 pad-t10">Total sales price DKK <span style="color:#999999;">(Order Qty. * Sales PC)</span>
                          
                                <input class="form-control input-small" type="text" id="bruttooms_label" value="<%=formatnumber(bruttooms, 2) %>" DISABLED/>
                                <input class="form-control input-small"type="hidden" name="FM_bruttooms" id="bruttooms" value="<%=formatnumber(bruttooms, 2) %>"/>
                          </div>
                             
                         
                                <div class="col-lg-3 pad-t10"><br />Total Profit app. %
                          
                                    <input class="form-control input-small" type="text" id="jo_dbproc_label" value="<%=jo_dbproc %>" DISABLED />
                                    <input class="form-control input-small"type="hidden" name="FM_jo_dbproc" id="jo_dbproc" value="<%=jo_dbproc %>" />
                                </div>
                             
                                
                        
                            

                             <%jo_dbproc_bel = formatnumber((bruttooms - jo_udgifter_intern), 2) %>
                                 
                             
                                <div class="col-lg-3 pad-r30 pad-t10"><br />DKK
                                    <input class="form-control input-small" type="text" name="" id="jo_dbproc_bel" value="<%=jo_dbproc_bel %>" disabled />
                                </div>

                               
                                
                            
                          </div>
                        <!--
                        <div class="form-group">
                            Profit/PC
                            <input class="form-control input-small" type="text" value="DKK" disabled />
                        </div>
                        -->
                       

                        <div class="row">
                              <div class="col-lg-6 pad-t10"><input type="hidden" id="FM_t5" value="0" />
                            Payment Term Buyer
                             <%
                            disa = 0
                            lang = 1
                            nameid = "FM_betbetint"
                            call betalingsbetDage(kunde_betbetint, disa, lang, nameid) %>
                            </div>

                        
                                 <div class="col-lg-6 pad-t10 pad-r30">
                                Delivery Term Buyer 
                                <select class="form-control input-small" name="FM_levbetint">
                                    <option value="0" <%=kunde_levbetint0SEL %>>Choose..</option>
                                    <option value="1" <%=kunde_levbetint1SEL %>>FOB</option>
                                    <option value="2" <%=kunde_levbetint2SEL %>>DDP</option>
                                    <option value="3" <%=kunde_levbetint3SEL %>>CIF</option>

                                </select></div>
                             
                            
                        </div>

                        
                         <div class="row">
                                   <div class="col-lg-6 pad-t10">
                                Payment Term Supplier
                                <%
                                disa = 0
                                lang = 1
                                nameid = "FM_levbetbetint"
                                call betalingsbetDage(lev_betbetint, disa, lang, nameid) %>
                                    </div>
                               
                                   <div class="col-lg-6 pad-t10 pad-r30"">
                            Delivery Term Supplier
                         
                              <select class="form-control input-small" name="FM_levlevbetint">
                                <option value="0" <%=levlevbetint0SEL %>>Choose..</option>
                                <option value="1" <%=levlevbetint1SEL %>>FOB</option>
                                <option value="2" <%=levlevbetint2SEL %>>DDP</option>
                                <option value="3" <%=levlevbetint3SEL %>>CIF</option>

                            </select>
                        </div>
                          
                   </div>


                         
                       
               </section>
            
             <!--PRICE TERMS END -->

            </div>
            </div>
            </div>
            </div>
          
            <!--PRICE TERMS TOP END -->






            <!--ORDER INFO START -->
              
               <div class="panel-group accordion-panel" id="accordion-paneled1">
                    <!-- ORDER INFO TOP -->
                        <div class="panel panel-default"> 
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseTree">Order info</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseTree" class="panel-collapse collapse">
                       
                            <div class="panel-body">



        
             <section>
          
                       <div class="row">
                            <div class="col-lg-3 pad-t10">
                            Buyers Order No. (PO no.)
                            <input class="form-control input-small" type="text" name="FM_rekvnr" value="<%=rekvnr %>" />
                        </div>


                       <div class="col-lg-3 pad-t10">
                            Order placement
                            <input class="form-control input-small" type="text" name="FM_dt_jobstdato" value="<%=dt_jobstdato%>" placeholder="dd-mm-yyyy"/>
                        </div>

                       <div class="col-lg-3 pad-t10">
                            Destination
                            <select class="form-control input-small"name="FM_destination">
                                <option>Denmark</option>
                                 <option disabled>------------------------</option>
                                 <option value="0" disabled>Choose destination</option>
                                <%if func = "red" then%>
		                        <option SELECTED><%=destination%></option>
                                <%end if %>

                               
                               <!--#include file="../timereg/inc/inc_option_land.asp"-->
                            </select>
                        </div>

                       <div class="col-lg-3 pad-t10 pad-r30">
                            Transport
                            <select class="form-control input-small"name="FM_transport">
                                <option value="0">Choose transport</option>
                                <option value="By Air" <%=transportFlSel %>>By Air</option>
                                <option value="By Sea" <%=transportShSel %>>By Sea (payment terms +5 weeks)</option>
                                <option value="By Train" <%=transportTrSel%>>By Train</option>
                                <option value="By Truck" <%=transportRdSel %>>By Truck (payment terms +7 days)</option>
                                <option value="By Currier" <%=transportCurSel %>>By Currier</option>
                            </select>
                        </div>

                        </div>


                        <div class="row">

                         <div class="col-lg-2 pad-t10">
                            Confirmed buyer ETD <span style="color:red;">(*)</span>
                           
                              <div class='input-group date'>
                                  <input class="form-control input-small" type="text" name="FM_dt_confb_etd" value="<%=dt_confb_etd%>" placeholder="dd-mm-yyyy" />
                            
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                        </div>

                            <div class="col-lg-2 pad-t10">
                            Confirmed supplier ETD
                               <div class='input-group date'>
                                 <input class="form-control input-small" type="text" name="FM_dt_confs_etd" value="<%=dt_confs_etd%>" placeholder="dd-mm-yyyy" />
                            
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                        </div>

                          <div class="col-lg-2 pad-t10">
                            Conf. buyer ETA <span style="color:red;">(* <span style="font-size:9px;">on DDP</span>)</span>
                               <div class='input-group date'>
                                 <input class="form-control input-small" type="text" name="FM_dt_confb_eta" value="<%=dt_confb_eta%>" placeholder="dd-mm-yyyy" />
                                     <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                        </div>
                         
                          <div class="col-lg-2 pad-t10">
                            Confirmed supplier ETA
                           
                               <div class='input-group date'>
                                  <input class="form-control input-small" type="text" name="FM_dt_confs_eta" value="<%=dt_confs_eta%>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                        </div>
                       
                          <div class="col-lg-2 pad-t10">
                            Actual ETD <span style="color:red;">(*)</span> <span style="font-size:10px;">(= invoice date)</span>
                            
                               <div class='input-group date'>
                                 <input class="form-control input-small" type="text" name="FM_dt_actual_etd" value="<%=dt_actual_etd%>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                        </div>
                         <div class="col-lg-2 pad-t10 pad-r30">
                            Act. ETA <span style="font-size:10px;">(= inv.date DDP)</span>
                              <div class='input-group date'>
                                 <input class="form-control input-small" type="text"  name="FM_dt_actual_eta" value="<%=dt_actual_eta%>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            
                        </div>

                        </div>
                        

                        <div class="row">
                        <div class="col-lg-3 pad-t10">
                            Shipped Qty. <span style="color:red;">(*)</span>
                            <input class="form-control input-small" type="text" id="shippedqty" name="FM_shippedqty" value="<%=shippedqty%>" />
                        </div>
                            <div class="col-lg-5 pad-t10 pad-r30">
                                &nbsp;
                                </div>
                                 <div class="col-lg-4 pad-t10 pad-r30">
                                 &nbsp;<!-- Exp. invoice duedate: -->
                                </div>
                             
                            
                                
                        </div>

                      

                     <div class="row">
                       <div class="col-lg-3 pad-t10">
                            Supplier invoice No. <span style="color:red;">(*)</span>
                            <input class="form-control input-small" type="text" name="FM_supplier_invoiceno" value="<%=supplier_invoiceno%>"  />
                        </div>
                                <div class="col-lg-9 pad-t10 pad-r30">
                                &nbsp;
                                </div>
                       </div>
                        <!--
                        <div class="form-group">
                            Invoiced
                            <select class="form-control input-small" type="text">
                                <option>Invoiced</option>
                                <option>1</option>
                                <option>2</option>
                            </select>
                        </div>
                        <div class="form-group">
                            Payed
                            <select class="form-control input-small" type="text">
                                <option>Payed</option>
                                <option>1</option>
                                <option>2</option>
                            </select>
                        </div>
                        -->
                            <div class="row">
                            <div class="col-lg-12 pad-t10 pad-r30">

                       
                            Invoice Note (visible on invoice)
                            <textarea class="form-control input-small" name="FM_beskrivelse" style="height:100px;" placeholder="Write your comment here"><%=beskrivelse %></textarea>
                        </div>
                                </div>
            </section>
            
              <!--ORDER INFO END -->

            </div>
            </div>
            </div>
            </div>
          
               

            <!--ORDER INFO TOP END -->
           


            <!--PRODUCTION INFO START -->
            <!--PRICE TERMS TOP START -->
               <div class="panel-group accordion-panel" id="accordion-paneled2">
                    <!-- ENQUIRY INFO TOP -->
                        <div class="panel panel-default"> 
                        <div class="panel-heading">
                          <h4 class="panel-title">
                            <a class="accordion-toggle" data-toggle="collapse" data-target="#collapseFour">Production info</a>
                          </h4>
                        </div> <!-- /.panel-heading -->
                        <div id="collapseFour" class="panel-collapse collapse">
                       
                            <div class="panel-body">



             <!-- PRICE TERMS START -->
             <section>
             <!--<div class="well well-white">-->

                   <div class="row">
                        <div class="col-lg-3 pad-t10">
                            1st. order comment
                             <div class='input-group date'>
                                 <input class="form-control input-small" type="text" name="FM_dt_firstorderc" value="<%=dt_firstorderc%>" placeholder="dd-mm-yyyy"/>
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                        </div>
                        <div class="col-lg-3 pad-t10">
                            LD app
                             <div class='input-group date'>
                                  <input class="form-control input-small" type="text" name="FM_dt_ldapp" value="<%=dt_ldapp%>" placeholder="dd-mm-yyyy"/>
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            
                        </div>
                        <div class="col-lg-3 pad-t10">
                            Exp Sizeset
                            <div class='input-group date'>
                                  <input class="form-control input-small" type="text" name="FM_dt_sizeexp" value="<%=dt_sizeexp%>" placeholder="dd-mm-yyyy"/>
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                           
                        </div>

                        <div class="col-lg-3 pad-t10 pad-r30">
                            Sizeset app
                            <div class='input-group date'>
                                 <input class="form-control input-small" type="text" name="FM_dt_sizeapp" value="<%=dt_sizeapp%>" placeholder="dd-mm-yyyy"/>
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                        </div>

                    </div>

                    <div class="row">
                          <div class="col-lg-3 pad-t10">
                            Exp PP
                              <div class='input-group date'>
                                 <input class="form-control input-small" type="text" name="FM_dt_ppexp" value="<%=dt_ppexp%>" placeholder="dd-mm-yyyy"/>
                                    <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                           
                        </div>
                          <div class="col-lg-3 pad-t10">
                            PP app
                                <div class='input-group date'>
                                <input class="form-control input-small" type="text" name="FM_dt_ppapp" value="<%=dt_ppapp%>" placeholder="dd-mm-yyyy"/>
                                    <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            
                        </div>
                          <div class="col-lg-3 pad-t10">
                            Exp SHS
                                    <div class='input-group date'>
                                <input class="form-control input-small" type="text" name="FM_dt_shsexp" value="<%=dt_shsexp%>" placeholder="dd-mm-yyyy"/>
                                    <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            
                        </div>
                          <div class="col-lg-3 pad-t10 pad-r30">
                            SHS app
                                <div class='input-group date'>
                                <input class="form-control input-small" type="text" name="FM_dt_shsapp" value="<%=dt_shsapp%>" placeholder="dd-mm-yyyy"/>
                                    <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                          </div>
                     
                    </div>
                 
                        <div class="row">
                                <div class="col-lg-11 pad-t10">
                                    &nbsp;
                                </div>
                         <div class="col-lg-2 pad-t10 pad-r30 pull-right">
                        <input class="btn btn-success" type="submit" value="Submit">
                               </div>

                          </div>

                </section>

               
                     </form>
         
                
               </section>
            
              <!--PRODUCTION INFO END -->

            </div>
            </div>
            </div>
            </div>
          
           
            <!--PRODUCTION INFO TOP END -->


        </div> <!-- portlet BODY -->
    </div> <!-- portlet -->



<%
 
    case else   
    
    if len(trim(request("post"))) <> 0 then
    post = 1
    else
    post = 0
    end if
    if len(trim(request("lastid"))) <> 0 then
    lastid = request("lastid")
    else
    lastid = 0
    end if
   if len(trim(request("rapporttype"))) <> 0 then
   rapporttype = request("rapporttype")
   response.cookies("orders")("rapporttype") = rapporttype 
   else
        if request.cookies("orders")("rapporttype") <> "" then
        rapporttype = request.cookies("orders")("rapporttype")
        else
        rapporttype = 0
        end if
    end if
   if len(trim(request("supplier"))) <> 0 then
   supplier = request("supplier")
   response.cookies("orders")("supplier") = supplier 
   else
        if request.cookies("orders")("supplier") <> "" then
        supplier = request.cookies("orders")("supplier")
        else
        supplier = 0
        end if
    end if
    if len(trim(request("buyer"))) <> 0 then
   buyer = request("buyer")
   response.cookies("orders")("buyer") = buyer
   else
        if request.cookies("orders")("buyer") <> "" then
        buyer = request.cookies("orders")("buyer")
        else
        buyer = 0
        end if
    end if
    
   if len(trim(request("salesrep"))) <> 0 then
   salesrep = request("salesrep")
   response.cookies("orders")("salesrep") = salesrep 
   else
        if request.cookies("orders")("salesrep") <> "" then
        salesrep = request.cookies("orders")("salesrep")
        else
        salesrep = 0
        end if
    end if
    if salesrep <> 0 then
    salesrepSQL = " AND j.jobans1 = "& salesrep
    else
    salesrepSQL = ""
    end if
    if supplier <> 0 then
    supplierSQL = " AND j.supplier = "& supplier
    else
    supplierSQL = ""
    end if
    if buyer <> 0 then
    buyerSQL = " AND k.kid = "& buyer
    else
    buyerSQL = ""
    end if
    'Rapporttyper
    '0: overview 
    '1: enquery  / case else
    '2: production
    '3: production no eco
    '4: Sales budget
    select case rapporttype
    case 1
    rapporttypeTxt = "Production (Enq. Overview)"
    case 3
    rapporttypeTxt = "Production Orders (Ext. Overview)"
    'case 4
    'rapporttypeTxt = "Production (simple)"
    'case 5
    'rapporttypeTxt = "Sales budget"
    case else
    rapporttypeTxt = "Orders (Overview)"
    end select
    if cint(post) = 1 then
    sogVal = request("FM_sog")
    response.cookies("orders")("sog") = sogVal
    else
        if request.cookies("orders")("sog") <> "" then
        sogVal = request.cookies("orders")("sog")
        else
        sogVal = ""
        end if
    end if
    if cint(post) = 1 then' fra Sog submit
        if len(trim(request("FM_status1"))) <> 0 then
        strStatusSQL = " AND (jobstatus = 1"
        statusCHK1 = "CHECKED"
        response.cookies("orders")("status1") = "1" 
        else
        strStatusSQL = " AND (jobstatus = -1"
        statusCHK1 = ""
        response.cookies("orders")("status1") = ""
        end if
    else
            if len(trim(request("rapporttype"))) <> 0 then 'altid slået til når der vælges en rapport fra menu, uanset hvilken          
            strStatusSQL = " AND (jobstatus = 1"
            statusCHK1 = "CHECKED"
            response.cookies("orders")("status1") = "1"
            else 'Cookie
                if request.cookies("orders")("status1") <> "" then
                strStatusSQL = " AND (jobstatus = 1"
                statusCHK1 = "CHECKED"
                else
                strStatusSQL = " AND (jobstatus = -1"
                statusCHK1 = ""
                response.cookies("orders")("status1") = "" 
                end if
            
            end if
    end if
    if len(trim(request("FM_status2"))) <> 0 then
    strStatusSQL = strStatusSQL & " OR jobstatus = 2"
    statusCHK2 = "CHECKED"
    response.cookies("orders")("status2") = "1" 
    else
        if cint(post) = 1 then
        statusCHK2 = ""
        response.cookies("orders")("status2") = ""
        else
            if request.cookies("orders")("status2") <> "" then
            strStatusSQL = strStatusSQL & " OR jobstatus = 2"
            statusCHK2 = "CHECKED"
            else
            statusCHK2 = ""
            end if
        end if
    end if
  if len(trim(request("FM_status3"))) <> 0 then
    strStatusSQL = strStatusSQL & " OR jobstatus = 3"
    statusCHK3 = "CHECKED"
    response.cookies("orders")("status3") = "1" 
    else
        if cint(post) = 1 then
        statusCHK3 = ""
        response.cookies("orders")("status3") = ""
        else
            if request.cookies("orders")("status3") <> "" then
            strStatusSQL = strStatusSQL & " OR jobstatus = 3"
            statusCHK3 = "CHECKED"
            else
                
                  if request("rapporttype") = "1" then 'altid slået til når der vælges en Enqueries fra menu  
                  strStatusSQL = strStatusSQL & " OR jobstatus = 3"
                  statusCHK3 = "CHECKED"  
                  else
                  statusCHK3 = ""
                  end if
            end if
        end if
    end if
   if len(trim(request("FM_status4"))) <> 0 then
    strStatusSQL = strStatusSQL & " OR jobstatus = 4"
    statusCHK4 = "CHECKED"
    response.cookies("orders")("status4") = "1" 
    else
        if cint(post) = 1 then
        statusCHK4 = ""
        response.cookies("orders")("status4") = ""
        else
            if request.cookies("orders")("status4") <> "" then
            strStatusSQL = strStatusSQL & " OR jobstatus = 4"
            statusCHK4 = "CHECKED"
            else
            statusCHK4 = ""
            end if
        end if
    end if
       if len(trim(request("FM_status0"))) <> 0 then
    strStatusSQL = strStatusSQL & " OR jobstatus = 0"
    statusCHK0 = "CHECKED"
    response.cookies("orders")("status0") = "1" 
    else
        if cint(post) = 1 then
        statusCHK0 = ""
        response.cookies("orders")("status0") = ""
        else
            if request.cookies("orders")("status0") <> "" then
            strStatusSQL = strStatusSQL & " OR jobstatus = 0"
            statusCHK0 = "CHECKED"
            else
            statusCHK0 = ""
            end if
        end if
    end if
    
    strStatusSQL = strStatusSQL & ")"
    if len(trim(request("sort"))) <> 0 then
    sort = request("sort")
    response.cookies("orders")("orderby") = sort
    else
         if request.cookies("orders")("orderby") <> "" then
         sort = request.cookies("orders")("orderby")
         else
         sort = 1
         end if 
    end if
     select case sort
        case "1"
        strSQLOdrBy = "kkundenavn, jobnavn"
        case "2"
        strSQLOdrBy = "jobstatus, kkundenavn, jobnavn"
        case "3"
        strSQLOdrBy = "dt_sour_dead"
        case "4"
        strSQLOdrBy = "dt_proto_dead"
        case "5"
        strSQLOdrBy = "dt_photo_dead"
        case "6"
        strSQLOdrBy = "dt_sms_dead"
        case "7"
        strSQLOdrBy = "jobnavn"
        case "8"
        strSQLOdrBy = "dt_sizeapp"
        case "9"
        strSQLOdrBy = "dt_ppapp"
        case "10"
        strSQLOdrBy = "dt_shsapp"
        case "11"
        strSQLOdrBy = "jobstartdato"
        case "14"
        strSQLOdrBy = "dt_confb_etd"
            case "15"
                strSQLOdrBy = "rekvnr"
            case "16"
                strSQLOdrBy = "supplier"
            case "17"
                strSQLOdrBy = "product_group"
            case "18"
                strSQLOdrBy = "jobnr"
        case "19"
                strSQLOdrBy = "dt_actual_etd"
        case "20"
                strSQLOdrBy = "dt_confs_etd"
        case else 
        strSQLOdrBy = "kkundenavn, jobnavn"
        end select
     
    strSQLOdrBy = strSQLOdrBy & ", kkundenavn, supplier, jobnavn"
    
    if len(trim(request("FM_dt_from"))) <> 0 then
    dt_from = request("FM_dt_from")
        
    dt_from = replace(dt_from, ".", "-")
    dt_from = replace(dt_from, "/", "-")
    dt_from = replace(dt_from, ":", "-")
    dt_from = replace(dt_from, " ", "")
     if isDate(dt_from) = false then
    dt_from = "2010-01-01"
    end if
    response.cookies("orders")("dt_from") = dt_from
    else
        if request.cookies("orders")("dt_from") <> "" then
        dt_from = request.cookies("orders")("dt_from")
        else
        dt_from = dateAdd("m", -3, now)
        end if
    end if
    dt_fromTxt = formatdatetime(dt_from, 2) 'day(dt_from) &"-"& month(dt_from) &"-"& year(dt_from)
    dt_fromSQL = year(dt_from) &"-"& month(dt_from) &"-"& day(dt_from)
    if len(trim(request("FM_dt_to"))) <> 0 then
    dt_to = request("FM_dt_to")
    dt_to = replace(dt_to, ".", "-")
    dt_to = replace(dt_to, "/", "-")
    dt_to = replace(dt_to, ":", "-")
    dt_to = replace(dt_to, " ", "")
    if isDate(dt_to) = false then
    dt_to = "2010-01-01"
    end if
     
    response.cookies("orders")("dt_to") = dt_to
    else
        if request.cookies("orders")("dt_to") <> "" then
        dt_to = request.cookies("orders")("dt_to")
        else
        dt_to = dateAdd("m", -3, now)
        end if
    end if
     dt_toTxt = formatdatetime(dt_to, 2) '&"-"& month(dt_to) &"-"& year(dt_to)
    dt_toSQL = year(dt_to) &"-"& month(dt_to) &"-"& day(dt_to)
    if len(trim(request("FM_append_to"))) <> 0 then
    append_to = request("FM_append_to")
    response.cookies("orders")("apend_to") = append_to 
    else
        if request.cookies("orders")("apend_to") <> "" then
        append_to = request.cookies("orders")("apend_to")
        else
        append_to = "-1"
        end if
    end if
  
    appto_1Sel = ""
    appto_2Sel = ""
    appto_3Sel = ""
    appto_4Sel = ""
    appto_5Sel = ""
    appto_6Sel = ""
    appto_7Sel = ""
    appto_8Sel = ""
     appto_9Sel = ""
     appto_10Sel = ""
     appto_11Sel = ""
     appto_12Sel = ""
     appto_13Sel = ""
     appto_14Sel = ""
    appto_15Sel = ""
    select case append_to
    case "-1"
    strSQLdtKri = ""
    case "1"
    strSQLdtKri = " AND (dt_sour_dead " 
    appto_1Sel = "SELECTED"
    case "2"
    strSQLdtKri = " AND (dt_proto_dead "
    appto_2Sel = "SELECTED"
    case "3"
    strSQLdtKri = " AND (dt_photo_dead "
    appto_3Sel = "SELECTED"
    case "4"
    strSQLdtKri = " AND (dt_sms_dead "
    appto_4Sel = "SELECTED"
     case "5"
    strSQLdtKri = " AND (dt_sizeapp "
    appto_5Sel = "SELECTED"
    case "6"
    strSQLdtKri = " AND (dt_ppapp "
    appto_6Sel = "SELECTED"
    case "7"
    strSQLdtKri = " AND (dt_shsapp "
    appto_7Sel = "SELECTED"
     case "8"
    strSQLdtKri = " AND (jobstartdato "
    appto_8Sel = "SELECTED"
    'case "9"
    'strSQLdtKri = " AND (dt_ppapp "
    'appto_9Sel = "SELECTED"
    'case "10"
    'strSQLdtKri = " AND (dt_shsapp "
    'appto_10Sel = "SELECTED"
     case "11"
    strSQLdtKri = " AND (jobstartdato "
    appto_11Sel = "SELECTED"
     case "12"
    strSQLdtKri = " AND (dt_sup_photo_dead "
    appto_12Sel = "SELECTED"
     case "13"
    strSQLdtKri = " AND (dt_sup_sms_dead "
    appto_13Sel = "SELECTED"
    case "14"
    strSQLdtKri = " AND (IF(kunde_levbetint != 2, dt_confb_etd, dt_confb_eta)"
    appto_14Sel = "SELECTED"
    case "15"
    strSQLdtKri = " AND (dt_confs_etd "
    appto_15Sel = "SELECTED"
    end select
    if append_to <> "-1" then
    strSQLdtKri = strSQLdtKri & " BETWEEN '"& dt_fromSQL &"' AND '"& dt_toSQL &"' )"
    end if
    rapporttype0SEL = ""
    rapporttype1SEL = ""
    rapporttype3SEL = ""
    select case rapporttype
    case 0
    rapporttype0SEL = "SELECTED"
    case 1
    rapporttype1SEL = "SELECTED"
    case 3
    rapporttype3SEL = "SELECTED"
    case else
    rapporttype0SEL = "SELECTED"
    end select
    'if media = "" then
    'call menu_2014()
    'end if
    %>
    



<%if media = "" then %>





      <div class="portlet">
        <h3 class="portlet-title">
          <u>Orders</u>
        </h3>
        <div class="portlet-body">


<%end if %>

            <%if media = "" then %>
            <!--SEARCH START -->


              <form class="panel-body" method="post" action="job_nt.asp?post=1">
             <section>
                <div class="well well-white">
                    
                        <input type="hidden" id="showfullscreen" value="0" />

                       

                         <div class="row">
                       
                           <div class="col-lg-4">
                                Search: <input type="search" name="FM_sog" class="form-control input-small" value="<%=sogVal%>" placeholder="Search"/>
                               <span style="color:#999999; font-size:9px;">Style, Order No, PO no. or Sup. Invoice NO. or NT Invoice NO., Buyer
                                   <br />OR <b>Buyer</b> followed by, Style, Style, Style etc.<br />
                                   Order No > 1000 </span>
                               </div>


                                <div class="col-lg-4">Tools & Functions

                           
                                    <br /><a href="job_nt.asp?media=exp" target="_blank">Export .CSV</a>
                                    <br /><a href="job_nt.asp?media=print" target="_blank">Print</a>
                                     <!--<br /><a href="job_nt.asp?func=opret" target="_blank"><b>New Order</b></a>-->
                                    
                                    <br /><a class="btn btn-success btn-sm" href="job_nt.asp?func=opret" role="button" target="_blank">New Order</a>
                                     
                                     </div>

                             <div class="col-lg-4" id="td_listtotals">&nbsp;
                                 </div>

                             
                        
                        </div>


                          <div class="row">
                       
                           <div class="col-lg-4 pad-t15">
                            Date from:
                            <input type="text" name="FM_dt_from" value="<%=dt_fromTxt%>" placeholder="dd-mm-yyyy" class="form-control input-small" />
                        </div>
                         <div class="col-lg-4 pad-t15">
                            To:
                            <input type="text" name="FM_dt_to" value="<%=dt_toTxt%>" placeholder="dd-mm-yyyy" class="form-control input-small" />
                       

                          </div>
                         <div class="col-lg-4 pad-t15 pad-r30">
                            Append to:
                             <select name="FM_append_to" onchange="submit();" class="form-control input-small">
                                 <option value="-1">Ignore dates</option>

                                 <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then  %>
                           <option value="8" <%=appto_8Sel %>>Order date</option>
                            <option value="14" <%=appto_14Sel %>>Conf. ETD Buyer (ETA on DDP orders)</option>
                              
                                 <%end if %>

                                 <%if cint(rapporttype) = 1 then  %>
                        
                            <option value="2" <%=appto_2Sel %>>Proto DL</option>
                            <option value="3" <%=appto_3Sel %>>Photo Buyer DL</option>
                            <option value="12" <%=appto_12Sel %>>Photo Supp. DL</option>
                            <option value="4" <%=appto_4Sel %>>SMS Buyer DL</option>
                            <option value="13" <%=appto_13Sel %>>SMS Supp. DL</option>

                             <option value="6" <%=appto_6Sel %>>PP App</option>
                             <option value="7" <%=appto_7Sel %>>SHS App</option>
                               <option value="14" <%=appto_14Sel %>>Conf. ETD Buyer (ETA on DDP orders)</option>

                                 <%end if %>

                                 <%if cint(rapporttype) = 3 then  %>
                            <option value="15" <%=appto_15Sel %>>ETD Suppl.</option>

                                 <%end if %>


                                 </select>
                               </div>

                 </div><!-- row -->

                        <div class="row">

                               <div class="col-lg-4 pad-t15">Buyer:<select name="buyer" class="form-control input-small" onchange="submit();">
                                  <option value="0">Choose..</option>
                               <%   strSQL = "SELECT kid, kkundenavn FROM kunder WHERE useasfak = 0 ORDER BY kkundenavn"
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF
        
                                    if cdbl(buyer) = oRec("kid") then
                                    ssel = "SELECTED"
                                    else
                                    ssel = ""
                                    end if
                             %>
                                  <option value="<%=oRec("kid")%>" <%=ssel %>><%=oRec("kkundenavn") %></option>
                                  <%
            
				
			                oRec.movenext
			                wend
			                oRec.close
                                    %>
                                  </select>
                               </div>

                             <div class="col-lg-4 pad-t15">Supplier:<select name="supplier" class="form-control input-small" onchange="submit();">
                                  <option value="0">Choose..</option>
                               <%   strSQL = "SELECT kid, kkundenavn FROM kunder WHERE useasfak = 6 ORDER BY kkundenavn"
                                    oRec.open strSQL, oConn, 3
                                    while not oRec.EOF
        
                                    if cdbl(supplier) = oRec("kid") then
                                    ssel = "SELECTED"
                                    else
                                    ssel = ""
                                    end if
                             %>
                                  <option value="<%=oRec("kid")%>" <%=ssel %>><%=oRec("kkundenavn") %></option>
                                  <%
            
				
			                oRec.movenext
			                wend
			                oRec.close
                                    %>
                                  </select>
                               </div>
                                 

                              <div class="col-lg-4 pad-t15 pad-r30 ">  <%  '**** Sales rep ***'
                                    call salesreplist(salesrep)
                                %>
                        
                         
                            Sales Rep.: <select name="salesrep" class="form-control input-small" onchange="submit();">
                           <%=strFil_Kpers_Txt %>
                            </select>


                               </div>

                        </div>


                           <div class="row">
                       
                               <div class="col-lg-8 pad-t15">Status:<br />
                               <input type="checkbox" name="FM_status3" value="3" class="checkbox-inline" <%=statusCHK3 %>> Enquiry
                               <input type="checkbox" name="FM_status1" value="1" class="checkbox-inline" <%=statusCHK1 %>> Active orders
                               <input type="checkbox" name="FM_status2" value="2" class="checkbox-inline" <%=statusCHK2 %>> Shipped orders
                               <input type="checkbox" name="FM_status0" value="0" class="checkbox-inline" <%=statusCHK0 %>> Invoiced/closed
                               <input type="checkbox" name="FM_status4" value="4"  class="checkbox-inline" <%=statusCHK4 %>> Cancelled


                               </div>

                               
                                     <div class="col-lg-4 pad-t15 pad-r30"> Order view:<select name="rapporttype" class="form-control input-small" onchange="submit();">
                            <option value="0" <%=rapporttype0SEL %>>Orders (Overview)</option>
                            <option value="1" <%=rapporttype1SEL %>>Production (Enq. Overview)</option>
                            <option value="3" <%=rapporttype3SEL %>>Production Orders (Ext. Overview)</option>
                            </select>
                               </div>

                                    

                               </div><!-- row -->

                                 
                       
                             

                          



                          

     <div class="row">

          <div class="col-lg-12 pad-r30 pad-t20"><input type="submit" class="btn btn-secondary pull-right" value="Search"></div>
            </div><!-- row -->

                </div>
                </section>
               </form>


                         

                        
                      
              

                    

            <!--SEARCH END -->

            <%end if 'media %>

        



            <% 
                
             
            antal_orders = 0    
                
             if media <> "exp" then %>
            <!--TABLE START -->
            
                
            
        <!--
          <form class="mainnav-form pull-right" role="search" style="margin-top:-15px;margin-bottom:15px;">
          <input type="text" class="form-control input-md mainnav-search-query" placeholder="Search">
          <button class="btn btn-sm mainnav-form-btn"><i class="fa fa-search"></i></button>
        </form>
        -->
            <!--
        <button class="btn btn-sm btn-success pull-right" style="margin-top:-15px;margin-bottom:15px;">Create</button>
            -->

            <%select case cint(rapporttype) 
            case 0
            tblListWdt = 1510
            case 1 'Prod / Enquery
            tblListWdt = 1650 
            case 3 'Overview
            tblListWdt = 2150
            end select    %>
        
       
            <input type="hidden" id="rapporttype" value="<%=rapporttype%>" />
             <div class="portlet-body" id="xtbl_orderlist" style="width:<%=tblListWdt%>px; padding:0px 0px 0px 0px; background-color:#ffffff;">          
                <%
                
                call tableheader %>
              
                           
            <%end if %>



<%
'if len(trim(sogVal)) <> 0 then
'    sogValSQL = " AND (k.kkundenavn LIKE '"& sogVal &"%' OR jobnavn LIKE '"& sogVal &"%' OR jobnr LIKE '"& sogVal &"%' )"
'else
'    sogValSQL = ""
'end if
call basisValutaFN()
basisValKursUse = replace(basisValKurs, ".", ",")
'************************************************************************
'****************** Søge funktion , Trim ********************************
'************************************************************************
if len(trim(sogVal)) <> 0 then 
                                    
                 
                    if instr(sogVal, ",") <> 0 then
                                    
                    sogValArr = split(sogVal, ",")
                                    
                                   
                    for j = 0 TO UBOUND(sogValArr)
                    sogValTxt = trim(sogValArr(j)) 
                    if j = 0 then
                        strsogValKri = " AND ((k.kkundenavn LIKE '%"& sogValTxt &"%' OR k.kkundenr = '"& sogValTxt &"') "
                    else
                        if j = 1 then
                        strsogValKri = strsogValKri & " AND ((jobnr LIKE '"& sogValTxt &"%' OR jobnavn LIKE '%"& sogValTxt &"%' OR supplier_invoiceno LIKE '"& sogVal &"%' OR rekvnr LIKE '"& sogVal &"%') "
                        else
                        strsogValKri = strsogValKri & " OR (jobnr LIKE '"& sogValTxt &"%' OR jobnavn LIKE '%"& sogValTxt &"%' OR supplier_invoiceno LIKE '"& sogVal &"%' OR rekvnr LIKE '"& sogVal &"%') "
                        end if
                    end if
                    next
                    strsogValKri = strsogValKri & "))"
                                    
                    else
                        if instr(sogVal, ">") <> 0 then
                        sogValTxt = trim(replace(sogVal, ">", "")) 
                            call erDetInt(SQLBless(trim(sogValTxt)))
                            if len(trim(sogValTxt)) > 0 AND isInt = 0 then
                            strsogValKri = " AND (jobnr > "& sogValTxt &")"
                            else
                                strsogValKri = " AND (jobnr < 0)"
                            end if 
                        else
                            if instr(sogVal, "<") <> 0 then
                                sogValTxt = trim(replace(sogVal, "<", ""))
                                            
                                call erDetInt(SQLBless(trim(sogValTxt)))
				                           
                                      
                                if len(trim(sogValTxt)) > 0 AND isInt = 0 then
                                strsogValKri = " AND (jobnr < "& sogValTxt &")"
                                else
                                strsogValKri = " AND (jobnr < 0)"
                                end if 
                            
                            else	
                                    '**** Finder jobnr på fakturaer vhsi der er faktureret *****
                                    fakfundet = 0
                                    strSogFaknrJobids = " OR (jobnr = -1"
                                    strSQLfakjob = "SELECT fid, j.jobnr FROM fakturaer AS f LEFT JOIN job AS j ON (j.id = f.jobid) WHERE faknr LIKE '"& sogVal &"%' AND j.jobnr IS NOT NULL "
                    
                                    'response.Write strSQLfakjob
                                    'response.flush
    
    
                                    oRec4.open strSQLfakjob, oConn, 3
                                    while not oRec4.EOF 
                                    strSogFaknrJobids = strSogFaknrJobids & " OR jobnr = "& oRec4("jobnr")
                                    fakfundet = 1
                                    oRec4.movenext
                                    wend
                                    oRec4.close
                 
                                    if cint(fakfundet) = 1 then
                                    strSogFaknrJobids = strSogFaknrJobids & ")"
                                    else
                                    strSogFaknrJobids = ""
                                    end if
        
                                    'response.end
                                    	
		                    strsogValKri = " AND (jobnr LIKE '"& sogVal &"%' "& strSogFaknrJobids &" OR jobnavn LIKE '%"& sogVal &"%' OR k.kkundenavn LIKE '%"& sogVal &"%' OR k.kkundenr = '"& sogVal &"' OR supplier_invoiceno LIKE '"& sogVal &"%' OR rekvnr LIKE '"& sogVal &"%') "
                                    
                            end if
                    end if
                    end if
		else
		strsogValKri = ""
		end if				
    sogValSQL = strsogValKri 
    '******************************************************
strSQLjobsel = "SELECT j.id, jobnavn, jobnr, k.kkundenavn, k.kid, k2.kkundenavn AS suppliername, jobstatus, supplier, mg.navn AS mgruppenavn, rekvnr, supplier_invoiceno, fastpris, collection, kunde_levbetint, "
if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then
strSQLjobsel = strSQLjobsel &" jobstartdato, orderqty, shippedqty, product_group, kundekpers, m.mnavn AS salesrep, m.init,"
strSQLjobsel = strSQLjobsel &" comm_pc, jo_dbproc, destination, dt_confb_etd, dt_confb_eta, dt_confs_etd, dt_actual_etd, sales_price_pc, sales_price_pc_valuta, cost_price_pc, cost_price_pc_valuta, freight_pc, tax_pc,"
end if
if cint(rapporttype) = 1 then
strSQLjobsel = strSQLjobsel &" dt_proto_dead, dt_sms_dead, dt_photo_dead, dt_sour_dead, dt_proto_dead, dt_ppapp, dt_shsapp, dt_sup_photo_dead, dt_sup_sms_dead, dt_confb_etd, dt_confb_eta, dt_confs_etd,"
end if
   
     
strSQLjobsel = strSQLjobsel &" jo_bruttooms, jo_udgifter_intern, alert"_
&" FROM job AS j "_
&" LEFT JOIN kunder AS k ON (k.kid = j.jobknr)"_
&" LEFT JOIN kunder AS k2 ON (k2.kid = j.supplier)"
strSQLjobsel = strSQLjobsel &" LEFT JOIN materiale_grp AS mg ON (mg.id = j.product_group)"
if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then
strSQLjobsel = strSQLjobsel &" LEFT JOIN medarbejdere AS m ON (m.mid = j.jobans1)"
end if
if post <> 0 OR lastid <> 0 OR media = "exp" then
strSQLjobsel = strSQLjobsel &" WHERE j.id <> 0 "& sogValSQL &" "& strStatusSQL &" "& strSQLdtKri &" "& salesrepSQL &" "& supplierSQL &" "& buyerSQL &" ORDER BY "& strSQLOdrBy &" LIMIT 2000" 
else
strSQLjobsel = strSQLjobsel &" WHERE j.id = -1" 
end if
'if session("mid") = 1 then
'response.write strSQLjobsel
'response.Flush    
'end if
 
  jo_bruttooms_tot = 0
  jo_cost_tot = 0
  orderqtyTot = 0 
  shippedqtyTot = 0
  ordertypeTxt = ""
  strExpTxt = ""
oRec.open strSQLjobsel, oConn, 3
while not oRec.EOF
   
    'if cint(antal_orders) = 0 then
    %>
                 <!--<input type="hidden" id="fakhref_<%=oRec("id") %>" value="../timereg/erp_opr_faktura_fs.asp?visfaktura=1&visjobogaftaler=1&visminihistorik=1&FM_job=<%=oRec("id") %>&FM_kunde=<%=oRec("kid")%>&FM_aftale=0&reset=1&FM_usedatokri=1&FM_start_dag=<%=day(dt_from)%>&FM_start_mrd=<%=month(dt_from)%>&FM_start_aar=<%=year(dt_from)%>&FM_slut_dag=<%=day(dt_to)%>&FM_slut_mrd=<%=month(dt_to)%>&FM_slut_aar=<%=year(dt_to)%>" />-->
                  <input type="text" id="fakhref_<%=oRec("id") %>" value="../timereg/erp_opr_faktura_fs.asp?visfaktura=1&visjobogaftaler=1&visminihistorik=1&FM_job=<%=oRec("id")%>&FM_kunde=<%=oRec("kid")%>&FM_aftale=0&reset=1&FM_usedatokri=1&FM_start_dag=<%=day(dt_from)%>&FM_start_mrd=<%=month(dt_from)%>&FM_start_aar=<%=year(dt_from)%>&FM_slut_dag=<%=day(dt_to)%>&FM_slut_mrd=<%=month(dt_to)%>&FM_slut_aar=<%=year(dt_to)%>" style="visibility:hidden; display:none;" />
                                                        
    <%'end if
      if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then
    
     if oRec("orderqty") <> 0 then
    orderqty = oRec("orderqty")
    orderqtyTot = orderqtyTot + oRec("orderqty")  
    else
    orderqty = ""
    end if
    
     if oRec("shippedqty") <> 0 then
    shippedqty = oRec("shippedqty")
    shippedqtyTot = shippedqtyTot + oRec("shippedqty")
    else
    shippedqty = ""
    end if
    if instr(oRec("jobstartdato"), "2010") <> 0 then
    dt_jobstartdato = ""
    else
    dt_jobstartdato = year(oRec("jobstartdato"))&"/"&month(oRec("jobstartdato"))&"/"&day(oRec("jobstartdato"))  'replace(formatdatetime(oRec("jobstartdato"), 2), "-", ".")
    'dt_jobstartdato = formatdatetime(dt_jobstartdato, 2)
    end if
    
     end if
    if cint(rapporttype) = 1 then
    
    if instr(oRec("dt_sour_dead"), "2010") <> 0 then
    dt_sour_dead = ""
    else
    'dt_sour_dead = replace(formatdatetime(oRec("dt_sour_dead"), 2), "-", ".") ' oRec("dt_sour_dead")
    dt_sour_dead = year(oRec("dt_sour_dead"))&"/"&month(oRec("dt_sour_dead"))&"/"&day(oRec("dt_sour_dead")) 
    'dt_sour_dead = formatdatetime(dt_sour_dead, 2)
    end if
     if instr(oRec("dt_proto_dead"), "2010") <> 0 then
    dt_proto_dead = ""
    else
    'dt_proto_dead = replace(formatdatetime(oRec("dt_proto_dead"), 2), "-", ".")
    dt_proto_dead = year(oRec("dt_proto_dead"))&"/"&month(oRec("dt_proto_dead"))&"/"&day(oRec("dt_proto_dead")) 
    'dt_proto_dead = formatdatetime(dt_proto_dead, 2)
    end if
     
     if instr(oRec("dt_photo_dead"), "2010") <> 0 then
    dt_photo_dead = ""
    else
    'dt_photo_dead = replace(formatdatetime(oRec("dt_photo_dead"), 2), "-", ".") 
    dt_photo_dead = year(oRec("dt_photo_dead"))&"/"&month(oRec("dt_photo_dead"))&"/"&day(oRec("dt_photo_dead")) 
    'dt_photo_dead = formatdatetime(dt_photo_dead, 2)
    end if
    
    
    if instr(oRec("dt_sms_dead"), "2010") <> 0 then
    dt_sms_dead = ""
    else
    'dt_sms_dead = replace(formatdatetime(oRec("dt_sms_dead"), 2), "-", ".")
    dt_sms_dead = year(oRec("dt_sms_dead"))&"/"&month(oRec("dt_sms_dead"))&"/"&day(oRec("dt_sms_dead")) 
    'dt_sms_dead = formatdatetime(dt_sms_dead, 2)
    end if
   
    if instr(oRec("dt_proto_dead"), "2010") <> 0 then
    dt_proto_dead = ""
    else
    'dt_proto_dead = replace(formatdatetime(oRec("dt_proto_dead"), 2), "-", ".") 
    dt_proto_dead = year(oRec("dt_proto_dead"))&"/"&month(oRec("dt_proto_dead"))&"/"&day(oRec("dt_proto_dead")) 
    'dt_proto_dead = formatdatetime(dt_proto_dead, 2)
    end if
     
     if instr(oRec("dt_ppapp"), "2010") <> 0 then
    dt_ppapp = ""
    else
    'dt_ppapp = replace(formatdatetime(oRec("dt_ppapp"), 2), "-", ".")
    dt_ppapp = year(oRec("dt_ppapp"))&"/"&month(oRec("dt_ppapp"))&"/"&day(oRec("dt_ppapp")) 
    'dt_ppapp = formatdatetime(dt_ppapp, 2)
    end if
    
    
     if instr(oRec("dt_shsapp"), "2010") <> 0 then
    dt_shsapp = ""
     else
    'dt_shsapp = replace(formatdatetime(oRec("dt_shsapp"), 2), "-", ".") 'oRec("dt_shsapp")
    dt_shsapp = year(oRec("dt_shsapp"))&"/"&month(oRec("dt_shsapp"))&"/"&day(oRec("dt_shsapp")) 
    'dt_shsapp = formatdatetime(dt_shsapp, 2)
    end if
   
    if instr(oRec("dt_sup_photo_dead"), "2010") <> 0 then
    dt_sup_photo_dead = ""
    else
    'dt_sup_photo_dead = replace(formatdatetime(oRec("dt_sup_photo_dead"), 2), "-", ".") 'oRec("dt_sup_photo_dead")
    dt_sup_photo_dead = year(oRec("dt_sup_photo_dead"))&"/"&month(oRec("dt_sup_photo_dead"))&"/"&day(oRec("dt_sup_photo_dead")) 
    'dt_sup_photo_dead = formatdatetime(dt_sup_photo_dead, 2)
    end if
     if instr(oRec("dt_sup_sms_dead"), "2010") <> 0 then
     dt_sup_sms_dead = ""
     else
        if isDate(dt_sup_sms_dead) = true then
        'dt_sup_sms_dead = replace(formatdatetime(oRec("dt_sup_sms_dead"), 2), "-", ".") 'oRec("dt_sup_sms_dead")
        dt_sup_sms_dead = year(oRec("dt_sup_sms_dead"))&"/"&month(oRec("dt_sup_sms_dead"))&"/"&day(oRec("dt_sup_sms_dead")) 
        'dt_sup_sms_dead = formatdatetime(dt_sup_sms_dead, 2)
        else
        dt_sup_sms_dead = oRec("dt_sup_sms_dead")
        end if
    end if
       
    
    end if
    
     if cint(rapporttype) = 1 OR cint(rapporttype) = 3 then
      if instr(oRec("dt_confs_etd"), "2010") <> 0 then
    dt_confs_etd = ""
     else
        'dt_confs_etd = replace(formatdatetime(oRec("dt_confs_etd"), 2), "-", ".") 'oRec("dt_confs_etd")
        dt_confs_etd = year(oRec("dt_confs_etd"))&"/"&month(oRec("dt_confs_etd"))&"/"&day(oRec("dt_confs_etd")) 
        'dt_confs_etd = formatdatetime(dt_confs_etd, 2)
    end if
    end if
 
    if oRec("jo_udgifter_intern") <> 0 then
    jo_udgifter_intern = formatnumber(oRec("jo_udgifter_intern"), 2) 
    jo_udgifter_internTxt = formatnumber(oRec("jo_udgifter_intern"), 2) & " DKK"
    else
    jo_udgifter_intern = ""
    jo_udgifter_internTxt = ""
    end if
    if oRec("jo_bruttooms") <> 0 then
    jo_bruttoomsTxt = formatnumber(oRec("jo_bruttooms"), 2) & " DKK"
    jo_bruttooms = formatnumber(oRec("jo_bruttooms"), 2)
    else
    jo_bruttoomsTxt = ""
    jo_bruttooms = ""
    end if
    if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then
        
        if oRec("sales_price_pc") <> 0 then
        call valutakode_fn(oRec("sales_price_pc_valuta"))
        sales_price_pc_val = valutaKode_CCC 
        sales_price_pc = formatnumber(oRec("sales_price_pc"), 2) 
        else
        sales_price_pc = ""
        sales_price_pc_val = ""
        end if
        
    
    
         if oRec("cost_price_pc") <> 0 then
        
        call valutakode_fn(oRec("cost_price_pc_valuta"))
        cost_price_pc_val = valutaKode_CCC
        cost_price_pc = formatnumber(oRec("cost_price_pc"), 2) 
        else
        cost_price_pc = ""
        cost_price_pc_val = ""
        end if
        
        if oRec("comm_pc") <> 0 then
        comm_pc = formatnumber(oRec("comm_pc"), 2) ' & " %" 
        else
        comm_pc = ""
        end if
        if oRec("cost_price_pc") <> 0 then
        valutaKursOR3(oRec("sales_price_pc_valuta"))
        salgsprisKurs = dblKursOR3
        valutaKursOR3(oRec("cost_price_pc_valuta"))
        kostprisKurs = dblKursOR3
        tax_pc_calc = 0
        tax_pc_calc = (oRec("cost_price_pc")/1 * (oRec("tax_pc")/100))
        profit_pc = formatnumber((oRec("sales_price_pc") * (salgsprisKurs/100)) - ((oRec("cost_price_pc") + tax_pc_calc + oRec("freight_pc")) * (kostprisKurs/100)) * basisValKursUse/100, 2)  '& " "& basisValISO '& " ("& basisValKursUse &")"
        else
        profit_pc = ""
        end if
       
        
        if oRec("jo_dbproc") <> 0 then
        jo_dbproc = formatnumber(oRec("jo_dbproc"), 2) '& " %"
        else
        jo_dbproc = ""
        end if
        
        if oRec("destination") <> "0" then
        destination = oRec("destination")
        else
        destination = ""
        end if
        
       
        if instr(oRec("dt_actual_etd"), "2010") <> 0 then
        dt_actual_etd = ""
        else
        'dt_actual_etd = oRec("dt_actual_etd")
        dt_actual_etd = year(oRec("dt_actual_etd"))&"/"&month(oRec("dt_actual_etd"))&"/"&day(oRec("dt_actual_etd")) 
        'dt_actual_etd = formatdatetime(dt_actual_etd, 2)
        end if
        
    end if
     if instr(oRec("dt_confb_etd"), "2010") <> 0 then
    dt_confb_etd = ""
     else
    'dt_confb_etd = replace(formatdatetime(oRec("dt_confb_etd"), 2), "-", ".") 'oRec("dt_confb_etd")
    dt_confb_etd = year(oRec("dt_confb_etd"))&"/"&month(oRec("dt_confb_etd"))&"/"&day(oRec("dt_confb_etd")) 
    'dt_confb_etd = formatdatetime(dt_confb_etd, 2)
    end if
     if instr(oRec("dt_confb_eta"), "2010") <> 0 then
     dt_confb_eta = ""
     else
     dt_confb_eta = year(oRec("dt_confb_eta"))&"/"&month(oRec("dt_confb_eta"))&"/"&day(oRec("dt_confb_eta")) 
     end if
    'jo_dbproc_bel_tot = 0
     if jo_bruttooms <> "" AND jo_udgifter_intern <> "" then
     jo_dbproc_bel = formatnumber((oRec("jo_bruttooms")/1 - oRec("jo_udgifter_intern")/1), 2)
     jo_dbproc_bel_tot = jo_dbproc_bel_tot/1 + jo_dbproc_bel/1
     jo_dbproc_bel = jo_dbproc_bel 
     jo_dbproc_belTxt = jo_dbproc_bel & " DKK"
     else
     jo_dbproc_bel = ""
     jo_dbproc_belTxt = "" 
     end if
  
    jo_cost_tot = jo_cost_tot + oRec("jo_udgifter_intern")
    jo_bruttooms_tot = jo_bruttooms_tot + oRec("jo_bruttooms")
    
   
    select case oRec("jobstatus")
    case 0
    jobstatusTxt = "Closed"
    case 1
    jobstatusTxt = "Active"
    case 2
    jobstatusTxt = "Shipped"
    case 3
    jobstatusTxt = "Enquiry"
    case 4
    jobstatusTxt = "Review"
    end select
    select case oRec("fastpris")
    case 2
    ordertypeTxt = "Commission"
    case 3
    ordertypeTxt = "Salesorder"
    end select
                                  
    if media = "exp" then
    strExpTxt = strExpTxt & jobstatusTxt & ";" & oRec("kkundenavn") & ";"& oRec("rekvnr") &";"& oRec("suppliername") &";"& oRec("supplier_invoiceno") &";"& oRec("mgruppenavn") &";"& oRec("jobnavn") &";" & oRec("jobnr") & ";"& ordertypeTxt &";" 
    
    if cint(oRec("kunde_levbetint")) <> 2 then
        dt_confb_etd_a_txt = dt_confb_etd
        else
        dt_confb_etd_a_txt = dt_confb_eta
        end if
    if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 
        strExpTxt = strExpTxt & dt_jobstartdato &";"& oRec("destination") &";"& oRec("collection") &";"& oRec("salesrep") &";"& dt_confb_etd_a_txt &";"& dt_actual_etd &";" 
    end if
     if cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt  & dt_confs_etd & ";"
    end if
      if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & orderqty & ";" & shippedqty & ";" 
    end if
    if cint(rapporttype) = 1 then 
    strExpTxt = strExpTxt  & dt_proto_dead & ";" & dt_photo_dead & ";"& dt_sup_photo_dead  &";" & dt_sms_dead & ";" & dt_sup_sms_dead & ";" & dt_ppapp & ";" & dt_shsapp &";" & dt_confb_etd_a_txt & ";" & dt_confs_etd & ";"
    end if
    
     if cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & cost_price_pc & ";" & cost_price_pc_val & ";"
    end if
    if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & sales_price_pc & ";" & sales_price_pc_val & ";"
    end if
    if cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & comm_pc & ";"& profit_pc & ";" & jo_udgifter_intern & ";DKK;"
    end if
    if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & jo_bruttooms & ";DKK;"
    end if
    if cint(rapporttype) = 3 then 
    strExpTxt = strExpTxt & jo_dbproc_bel & ";DKK;" & jo_dbproc & ";"
    end if
    strExpTxt = strExpTxt & ";;;xx99123sy#z"
    end if
                             if media <> "exp" then
                             if cdbl(lastid) = oRec("id") then
                             trbgcol = "#FFFFe1"
                             else
                             trbgcol = ""
                             end if
                             %>
                           
                             <tr style="background-color:<%=trbgcol%>;">
                                 <td >
                                    <%=left(jobstatusTxt, 10) %>
                                 </td>
                                 
                                 <td style="text-align:left;">
                                    <%if media <> "print" then %>
                                    <a href="../to_2015/kunder.asp?func=red&id=<%=oRec("kid") %>"><%=left(oRec("kkundenavn"), 12) %></a>
                                    <%else %>
                                    <%=oRec("kkundenavn") %>
                                    <%end if %>
                                     <br /><span style="font-size:9px; color:#999999;"><%=left(oRec("collection"), 15) %></span> 
                                </td>
                                 <td ><%=oRec("rekvnr") %>
                                     <br /><span style="font-size:9px; color:#999999;"><%=left(destination, 3)%></span> 
                                     
                                 </td>

                                  <td ><%=left(oRec("suppliername"), 10) %>

                                      <%if oRec("jobstatus") = 2 AND len(trim(oRec("supplier_invoiceno"))) <> 0 then %>
                                      <br /><span style="font-size:9px; color:#999999;"><%=oRec("supplier_invoiceno") %></span>
                                      <%end if %>
                                  </td>

                                 
                                <!-- <td style="white-space:nowrap; width:<%=tdWdthTBC%>px; text-align:left;"><%=left(oRec("mgruppenavn"),10) %></td>-->

                                <td style="text-align:left;">
                                    <%if cint(oRec("alert")) = 1 then %>
                                    <span style="color:red;"><b>!</b>&nbsp;</span>
                                    <%end if %>

                                    <%if media <> "print" then %>
                                    <a href="job_nt.asp?func=red&jobid=<%=oRec("id")%>"><%=left(oRec("jobnavn"), 30) %></a>
                                     <%else %>
                                    <%=oRec("jobnavn") %>
                                    <%end if %>
                                     <br /><span style="font-size:9px; color:#999999;"><%=left(oRec("mgruppenavn"),15) %></span> 
                                    

                                </td>
                                
                                
                                <td >
                                     <%if media <> "print" then %>
                                    <a href="job_nt.asp?func=red&jobid=<%=oRec("id")%>"><%=oRec("jobnr") %></a>

                                    <%
                                        jobnr = oRec("jobnr")

                                        strSQLimg = "SELECT id, filnavn FROM filer WHERE filertxt = "& jobnr
                                        oRec2.open strSQLimg, oConn, 3
	                                    j = 0
	                                    if not oRec2.EOF then
                                        if len(trim(oRec2("filnavn"))) <> 0 then
                                        %>
                                          
                                            <span id="modal_<%=jobnr %>" style="color:cornflowerblue;" class="fa fa-image pull-right picmodal"></span>
                                           <!-- <button id="modal_<%=jobnr %>" class="picmodal">Open Modal</button> -->
                                            
                                            <!-- The Modal -->
                                            <div id="myModal_<%=jobnr %>" class="modal">

                                              <!-- Modal content -->
                                              <div class="modal-content">
                                                <img src="../inc/upload/<%=lto%>/<%=oRec2("filnavn")%>" alt='' border='0' width="100%" height="100%">
                                              </div>

                                            </div>

                                        <%
                                        j = j + 1                                    
	                                    end if
                                        end if	                                
	                                    oRec2.close 
                                    %>

                                    <br /><span style="font-size:9px; color:#999999;"><%=ordertypeTxt %></span> 
                                    <%else %>
                                    <%=oRec("jobnr") %>
                                    <%end if %>
                                </td>
                                
                                

                                     <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 'Overview%>
                                 <td style="white-space:nowrap;"><%=dt_jobstartdato%></td>
                                 <!--<td style="white-space:nowrap; width:<%=tdWdthTBB%>px;"><%=left(destination, 3)%></td>-->
                                 
                                 <!--<td style="white-space:nowrap;"><%=oRec("collection") %></td>-->
                                 
                                 <td style="white-space:nowrap; width:<%=tdWdthTBD%>px;"><%=oRec("init") %></td>
                                 
                                
                            
                                <%end if %>


                                   <%if cint(rapporttype) = 1 then %>
                                
                                 <td style="white-space:nowrap;"><%=dt_proto_dead %></td>
                                <td style="white-space:nowrap;"><%=dt_photo_dead %></td>
                                  <td style="white-space:nowrap;"><%=dt_sup_photo_dead %></td>
                                 <td style="white-space:nowrap;"><%=dt_sms_dead%></td>
                                 <td style="white-space:nowrap;"><%=dt_sup_sms_dead%></td>
                                  <td style="white-space:nowrap;"><%=dt_ppapp%></td>
                                   <td style="white-space:nowrap;"><%=dt_shsapp%></td>
                                  <%end if 'rapporttype %>

                                     <%if cint(rapporttype) = 0 OR cint(rapporttype) = 1 OR cint(rapporttype) = 3 then%>
                                 <td style="white-space:nowrap;">
                                     
                                     <%if cint(oRec("kunde_levbetint")) <> 2 then%>
                                     <%=dt_confb_etd %>
                                     <%else %>
                                     
                                         <%if len(trim(dt_confb_etd)) <> 0 then %>
                                         <%=dt_confb_etd %> <span style="font-size:8px;">ETD</span><br />
                                         <%end if %>

                                         <%if len(trim(dt_confb_eta)) <> 0 then %>
                                         <%=dt_confb_eta %> <span style="font-size:8px;">ETA</span>
                                         <%end if %>

                                     <%end if %>
                                 </td>
                                  <%end if 'rapporttype %>

                                  <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then%>
                                 <td style="white-space:nowrap;"><%=dt_actual_etd %></td>
                                  <%end if 'rapporttype %>


                                  <%if cint(rapporttype) = 1 OR cint(rapporttype) = 3 then %>
                                 <td style="white-space:nowrap;"><%=dt_confs_etd %></td>
                                 <%end if 'rapporttype %>
                                 
                                  <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 'Overview%>
                                 <td style="white-space:nowrap;"><%=orderqty %></td>
                                 <td style="white-space:nowrap;"><%=shippedqty %></td>
                                 <%end if %>


                                

                       

                                  <%if cint(rapporttype) = 3 then  %>
                                 <td style="text-align:right;"><%=cost_price_pc%></td><!-- &" "& cost_price_pc_val -->
                                 <%end if %>

                                     <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then  %>
                                <td style="text-align:right;"><%=sales_price_pc%></td> <!--&" "& sales_price_pc_val  -->
                                 <%end if %>

                                  <%if cint(rapporttype) = 3 then  %>
                                  <td style="text-align:right;"><%=comm_pc%></td>
                                 <td style="text-align:right;" ><%=profit_pc%></td>
                                <td style="text-align:right;"><%=jo_udgifter_intern%></td><!-- jo_udgifter_internTxt -->
                                 <%end if %>

                                 <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then  %>
                                <td style="text-align:right;"><%=jo_bruttooms%></td><!-- jo_bruttoomsTxt -->
                                 <%end if %>

                                  <%if cint(rapporttype) = 3 then  %>
                                
                                <td style="text-align:right;"><%=jo_dbproc_bel%></td><!-- jo_dbproc_belTxt -->
                                 <td style="text-align:right;"><%=jo_dbproc%></td>
                                   <%end if %>


                                  


                                    <td style="white-space:nowrap; font-size:9px; overflow:hidden;"><%if media <> "print" then %> 


                                        
                                                        <%'if ((level <= 2 OR level = 6)) then 'AND cint(oRec("jobstatus")) = 2) then 'shipped 
                                                            jobids = jobids & ","& oRec("id")
                                                            jobnrs = jobnrs & ","& oRec("jobnr")
                                                            %>
                                                      
                                                        <!--<a href="../timereg/erp_opr_faktura_fs.asp?visfaktura=1&visjobogaftaler=1&visminihistorik=1&FM_kunde=<%=oRec("kid")%>&FM_job=<%=oRec("id")%>&FM_aftale=0&reset=1&FM_usedatokri=1&FM_start_dag=<%=day(dt_from)%>&FM_start_mrd=<%=month(dt_from)%>&FM_start_aar=<%=year(dt_from)%>&FM_slut_dag=<%=day(dt_to)%>&FM_slut_mrd=<%=month(dt_to)%>&FM_slut_aar=<%=year(dt_to)%>" target="_blank" class=rmenu>Create Invoice >> </a>-->
                                                            
                                                            <%strSQLFakhist = "SELECT faknr, beloeb, fid AS fid, betalt, shadowcopy FROM fakturaer WHERE jobid = "& oRec("id") &" LIMIT 30"
                                                                
                                                                oRec6.open strSQLFakhist, oConn, 3
                                                                f = 0
                                                                while not oRec6.EOF
                                                                if f = 0 then
                                                                %>

                                                                <!--
                                                                <br />Invoices:
                                                                <span style="color:#999999; font-size:9px;">Not approved invoices<br /> will be deleted on create new invoice.</span>
                                                                -->
                                                                <%
                                                                end if
                                                                
                                                                     if oRec6("shadowcopy") <> 1 then 

                                                                         if oRec6("betalt") <> 1 then %> 
                                                                        <br /><a href="../timereg/erp_opr_faktura_fs.asp?func=red&id=<%=oRec6("fid")%>&visfaktura=2&visjobogaftaler=1&visminihistorik=1&FM_kunde=<%=oRec("kid")%>&FM_job=<%=oRec("id")%>&FM_aftale=0&reset=1&FM_usedatokri=1&FM_start_dag=<%=day(dt_from)%>&FM_start_mrd=<%=month(dt_from)%>&FM_start_aar=<%=year(dt_from)%>&FM_slut_dag=<%=day(dt_to)%>&FM_slut_mrd=<%=month(dt_to)%>&FM_slut_aar=<%=year(dt_to)%>" target="_blank" style="color:#999999; font-size:9px;"><%=oRec6("faknr")%></a> 
                                                                                              
                                                                        <%else %>
                                                                        <br />
                                                                        <a href="../timereg/erp_opr_faktura_fs.asp?func=red&id=<%=oRec6("fid")%>&visfaktura=2&visjobogaftaler=1&visminihistorik=1&FM_kunde=<%=oRec("kid")%>&FM_job=<%=oRec("id")%>&FM_aftale=0&reset=1&FM_usedatokri=1&FM_start_dag=<%=day(dt_from)%>&FM_start_mrd=<%=month(dt_from)%>&FM_start_aar=<%=year(dt_from)%>&FM_slut_dag=<%=day(dt_to)%>&FM_slut_mrd=<%=month(dt_to)%>&FM_slut_aar=<%=year(dt_to)%>" target="_blank" style="color:green; font-size:9px;"><i>V</i>&nbsp;&nbsp;<%=oRec6("faknr")%></a> 
                                                                   
                                                                        <%end if%>


                                                                     <%else 
                                                                         
                                                                        if oRec6("betalt") <> 1 then  %>
                                                                        <br /><span style="color:#999999; font-size:10px;">(<%=oRec6("faknr")%>)</span> 
                                                                        <%else %>
                                                                        
                                                                         <br />
                                                                             
                                                                           
                                                                              
                                                                              <% 
                                                                                  strSQLFakorg = "SELECT jobnr, f.jobid, f.fid AS fakid, fakadr FROM fakturaer f "_
                                                                                  &" LEFT JOIN job j ON (j.id = f.jobid) WHERE faknr = "& oRec6("faknr") &" AND shadowcopy <> 1 "

                                                                                  'response.write strSQLFakorg
                                                                                  'response.flush

                                                                                  fak_ordrenr = 0
                                                                                  fakid = 0
                                                                                  fak_jobid = 0
                                                                                  fak_kid = 0
                                                                                  oRec8.open strSQLFakorg, oConn, 3
                                                                                  if not oRec8.EOF then

                                                                                  fak_ordrenr = oRec8("jobnr")
                                                                                  fakid = oRec8("fakid")
                                                                                  fak_jobid = oRec8("jobid") 
                                                                                  fak_kid = oRec8("fakadr")

                                                                                  end if
                                                                                  oRec8.close

                                                                                  %>
                                                                                   

                                                                                    <a href="../timereg/erp_opr_faktura_fs.asp?func=red&id=<%=fakid%>&visfaktura=2&visjobogaftaler=1&visminihistorik=1&FM_kunde=<%=fak_kid%>&FM_job=<%=fak_jobid%>&FM_aftale=0&reset=1&FM_usedatokri=1&FM_start_dag=<%=day(dt_from)%>&FM_start_mrd=<%=month(dt_from)%>&FM_start_aar=<%=year(dt_from)%>&FM_slut_dag=<%=day(dt_to)%>&FM_slut_mrd=<%=month(dt_to)%>&FM_slut_aar=<%=year(dt_to)%>" target="_blank" style="color:dodgerblue; font-size:9px;">
                                                                                        <i>V</i>&nbsp; <%=oRec6("faknr")%> 
                                                                                        <br />(invoiced on: <%=fak_ordrenr %>)</a>
                                                                   

                                                                               
                                                                        <%end if%>


                                                                    <%end if%>
                                                                    
                                                                     <%if oRec6("shadowcopy") <> 1 then %>
                                                                     <%'formatnumber(oRec6("beloeb"), 0) %>
                                                                     <%end if %>
                                                                
                                                                <%
                                                                f = f + 1
                                                                oRec6.movenext
                                                                wend 
                                                                oRec6.close%>            

                                        
                                                        <%'else %>
                                                        &nbsp;
                                                        <%'end if %>

                                                        

                                     <%else %>
                                     &nbsp;
                                     <%end if %>
                                 </td>

                             <%if media <> "print" then %>
                                 <!--<td ><a href="job_nt.asp?func=slet&id=<%=oRec("id") %>" style="color:red;">X</a></td>-->
                                 
                                 <td ><input type="checkbox" value="<%=oRec("id") %>" id="bulk_jobid_<%=oRec("id") %>" name="FM_bulk_jobid" class="bulk_jobid" />
                                     &nbsp;<a href="job_nt.asp?func=slet&id=<%=oRec("id") %>" style="color:red;">X</a>
                                 </td>
                                <%end if %>

                            </tr>
                            
                            <%end if 'media %>
                       



<%
		
    antal_orders = antal_orders + 1
oRec.movenext
wend
oRec.close
    
    if media <> "exp" then%>  

                        </tbody>
                    <tfoot>
                        <tr>
                              <td><b>Total:</b></td>
                            <td><%=antal_orders %></td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                           <!-- prod grp  <td>&nbsp;</td>-->
                            <!-- DEST <td>&nbsp;</td>-->

                            <%if cint(rapporttype) = 0 then %>
                             <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <%end if %>

                        
                            <!-- Collection <td>&nbsp;</td>-->
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                          
                            <!--<td>&nbsp;</td>-->
                            

                
                             <%if cint(rapporttype) = 0 then %>
                            <td>&nbsp;</td>
                            <%end if %>   
                         

                             <%if cint(rapporttype) = 1 then %>
                             <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            
                            <%end if %>

                            
                               <%if cint(rapporttype) = 3 then %>
                             <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                           <td>&nbsp;</td>

                            
                            <%end if %>


                            <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then %>
                            <td>&nbsp;</td>
                            <td style="text-align:right;"><%=formatnumber(orderqtyTot, 0) %></td>
                            <td style="text-align:right;"><%=formatnumber(shippedqtyTot, 0) %></td>
                            <%end if %>

                            
                                   <%if cint(rapporttype) = 3 then %>
                             <td>&nbsp;</td>
                             <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                           
                            
                            <%end if %>

                            
                             <%if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then %>
                          
                            <td style="text-align:right;"><%=formatnumber(jo_cost_tot, 2)%></td>
                            <td style="text-align:right;"><%=formatnumber(jo_bruttooms_tot, 2)%></td>
                            <%end if %>

                             

                             <%if cint(rapporttype) = 3 then 
                                 
                                 if len(trim(jo_dbproc_bel_tot)) <> 0 then
                                 jo_dbproc_bel_tot = formatnumber(jo_dbproc_bel_tot, 2)
                                 else
                                 jo_dbproc_bel_tot = "0"
                                 end if
                                 %>
                            <td style="text-align:right;"><%=jo_dbproc_bel_tot%></td>
                            <td style="text-align:right;">
                                <%
                                    if (jo_bruttooms_tot) <> 0 AND len(trim(jo_bruttooms_tot)) <> 0 then
                                    jo_dbproc_tot = formatnumber(100-(jo_cost_tot/jo_bruttooms_tot) * 100, 2)
                                    else
                                    jo_dbproc_tot = 0
                                    end if %>

                                <%=jo_dbproc_tot %> 

                            </td>
                             
                            
                            <%end if %>
                    
                          

                           

                            <%if cint(rapporttype) = 1 then %>
                              <td>&nbsp;</td>
                              <!--<td>&nbsp;</td>-->
                            <%end if %>

                         <%if media <> "print" then %>
                         <td>&nbsp;</td> 
                        <td>&nbsp;</td>
                         <%end if %>

                       </tr>
                           
                           
                       </tfoot>
                    </table>
                 
            <br />&nbsp;
                 </div><!-- datatables wrapper --> 
               
              <!--  </section> -->
            
            </div><!--Table body div -->   
               
               
            </div><!-- list div -->
            <!--TABLE END -->


          <!--
               <div id="dv_grandtotal" style="position:absolute; top:100px;left:1580px; width:400px;">
                 <section class="panel">
                    <header class="panel-heading">Totals for orders on list:</header>
                    <div class="panel-body">
                    <div id="dv_topGt">
                        -->
                        
                    <form>

                        <textarea id="listtotals" style="visibility:hidden; display:none;">
                            Totals for orders on list:
                            <table cellpadding="2" cellspacing="1" border="0">
                            <tr>
                                <td style="vertical-align:bottom; white-space:nowrap;">Order Qty.:</td> <td style="text-align:right; white-space:nowrap;"><%=formatnumber(orderqtyTot, 0) %></td>
                                </tr><tr>
                                <td style="vertical-align:bottom; white-space:nowrap;">Shipped Qty.:</td> <td style="text-align:right; white-space:nowrap;"><%=formatnumber(shippedqtyTot, 0) %></td>
                                </tr>
                                 <%if cint(rapporttype) = 3 then %>
                                <tr>
                                    <td style="vertical-align:bottom; white-space:nowrap;">Total Costprice:</td><td style="text-align:right; white-space:nowrap;"><%=formatnumber(jo_cost_tot, 2) & " DKK" %></td>
                                </tr>
                                <%end if %>
                                <tr>
                                    <td style="vertical-align:bottom; white-space:nowrap; width:120px;">Total Salesprice:</td><td style="text-align:right; white-space:nowrap;"><%=formatnumber(jo_bruttooms_tot, 2) & " DKK" %></td>
                                <%if cint(rapporttype) = 3 then %>
                                </tr>
                                <tr>
                                <td style="vertical-align:bottom; white-space:nowrap;">Total Profit:</td> 
                                <td style="text-align:right; white-space:nowrap;"><%=formatnumber(jo_dbproc_bel_tot, 2) & " DKK" %></td>

                                 </tr>
                                <tr>
                                <td style="vertical-align:bottom; white-space:nowrap;">Profit %:</td>   
                                <td style="text-align:right; white-space:nowrap;">
                                    
                                    <% 'jo_cost_tot = replace(jo_cost_tot, ".", "") %>
                                    <% 'jo_bruttooms_tot = replace(jo_bruttooms_tot, ".", "") %>


                                    <%if jo_bruttooms_tot <> 0 AND jo_cost_tot <> 0 then %>
                                    <%=formatnumber(100-(jo_cost_tot/jo_bruttooms_tot)*100, 2) & " %" %>
                                    <%else %>
                                    0 %
                                    <%end if %>
                                </td>
                                
                                <%end if %>
                            </tr>
                         
                        </table></textarea>

                        </form>

            <!--
                    </div>
                        </div>
                     </div>
                </section>
                </div>
            -->


        <%if media = "" then %>

                    

 </div>
<!-- </div> -->

</div> <!--portlet-body -->
<!--</div>--> <!-- hybrid -->
</div> <!-- container -->
</div><!-- content --->
</div><!-- Wrapper -->
  
 
                          
        <%end if %>





                             <div id="dv_invoice" style="position:absolute; float:left; left:1200px; top:500px; z-index:2000; border:10px yellowgreen solid; padding:20px; visibility:hidden; display:none; background-color:#FFFFFF;">
                              
                               
                                     <table border="0" cellpadding="0" cellspacing="0" width="100%">
                                     <tr><td style="border:0px;" align="left"><b>Create invoice on selected orders? </b>
                                 <br />
                                   <a id="ainvlink" href="#" target="_blank" class=vmenu>Yes, create Invoice on selected orders >> </a>

                                         <br /><br />
                                         <span style="color:#999999; font-size:9px;">
                                        
                                         Actual ETD, Buyer info, Bankaccount, currency, Invoice note etc.
                                         will always be obtained from the first selected order on the list.
                                             </span>

                                         </td><td valign="top" style="width:40px; border:0px;"><span id="sp_dv_invoice" style="color:#999999; font-size:14px;"><b>X</b></span></td></tr>
                                        </table>               


                                 </div>


                <div id="dv_bulk" style="position:absolute; width:800px; left:100px; top:200px; z-index:2000; border:10px #CCCCCC solid; padding:20px; visibility:hidden; display:none; background-color:#FFFFFF;">
               
                     <div class="portlet">
        <h3 class="portlet-title">
          <u>Bulk update production info</u>
        </h3>
        <div class="portlet-body">
             <form class="panel-body" id="joblist" method="post" action="job_nt.asp?func=bulk">
                 <input type="hidden" id="bulk_jobids" name="bulk_jobids" value="0" />

                    <div class="form-group" style="float:right;">
                              <span style="color:red;" id="bulk_close">[X]</span>
                        </div>

             <section>
                <div class="well well-white">
                    <!--BASEDATA START -->
                         <div class="row">
                        <div class="col-lg-12">
                            <h4 class="panel-title-well">Enquery info</h4>
                        </div>
                    </div>
                    
                      
                       

                         <div class="row">
                        
                            <div class="col-lg-4">
                                SMS Buyer deadline
                                 <div class='input-group date'>
                                     <input type="text" class="form-control input-small" name="FM_dt_sms_dead" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            
                            </div>
                              <div class="col-lg-4">
                                SMS Supplier DL
                                   <div class='input-group date'>
                                      <input type="text" class="form-control input-small" name="FM_dt_sup_sms_dead" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                               
                            </div>
                             <div class="col-lg-4">
                                SMS sent
                                 <div class='input-group date'>
                                      <input type="text" class="form-control input-small" name="FM_dt_sms_sent" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                                
                            </div>

                        </div>

                         <div class="row">
                             <div class="col-lg-4">
                                Photo Buyer Deadline
                                  <div class='input-group date'>
                                     <input type="text" class="form-control input-small" name="FM_dt_photo_dead" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                                
                            </div>
                               <div class="col-lg-4">
                                Photo Supplier DL
                                   <div class='input-group date'>
                                     <input type="text" class="form-control input-small" name="FM_dt_sup_photo_dead" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                                
                            </div>

                            <div class="col-lg-4 pad-r30">
                                Photo sent
                                <div class='input-group date'>
                                      <input type="text" class="form-control input-small" name="FM_dt_photo_sent" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                               
                            </div>
                         </div>

                        
                          <div class="row">
                        <div class="col-lg-12 pad-t20"><br />
                            <h4 class="panel-title-well">Order Dates</h4>
                        </div>
                    </div>

                       <div class="row">

                               <div class="col-lg-4">
                                Confirmed buyer ETD
                                   <div class='input-group date'>
                                      <input type="text" class="form-control input-small" name="FM_dt_confb_etd" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                                
                            </div>

                            <div class="col-lg-4 pad-r30">
                                Confirmed supplier ETD
                                   <div class='input-group date'>
                                       <input type="text" class="form-control input-small" name="FM_dt_confs_etd" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                                
                            </div>

                             <div class="col-lg-4">
                                Confirmed buyer ETA
                                  <div class='input-group date'>
                                       <input type="text" class="form-control input-small" name="FM_dt_confb_eta" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                               
                            </div>
                             

                        </div>

                         <div class="row">

                         <div class="col-lg-4">
                            Conf. suppl. ETA
                              <div class='input-group date'>
                                       <input type="text" class="form-control input-small" name="FM_dt_confs_eta" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            
                        </div>
                       
                        <div class="col-lg-4">
                            Actual ETD
                            <div class='input-group date'>
                                       <input type="text" class="form-control input-small" name="FM_dt_actual_etd" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            
                        </div>
                        <div class="col-lg-4 pad-r30">
                            Actual ETA
                             <div class='input-group date'>
                                       <input type="text" class="form-control input-small"  name="FM_dt_actual_eta" value="" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            
                        </div>

                        </div>

                      

                         <div class="row">
                            <div class="col-lg-12 pad-t20"><br />
                                <h4 class="panel-title-well">Production dates</h4>
                            </div>
                         </div>

                        <div class="row">

                             <div class="col-lg-4">
                                1st. order comment
                                  <div class='input-group date'>
                                       <input type="text" class="form-control input-small" name="FM_dt_firstorderc" value="" placeholder="dd-mm-yyyy"/>
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                                
                            </div>
                             <div class="col-lg-4">
                                LD app
                               
                                  <div class='input-group date'>
                                        <input type="text" class="form-control input-small" name="FM_dt_ldapp" value="" placeholder="dd-mm-yyyy"/>
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            </div>
                             <div class="col-lg-4 pad-r30">
                                Exp Sizeset
                                  <div class='input-group date'>
                                        <input type="text" class="form-control input-small" name="FM_dt_sizeexp" value="" placeholder="dd-mm-yyyy"/>
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                                
                            </div>

                         </div>

                         <div class="row">

                         <div class="col-lg-4">
                            Sizeset app
                             <div class='input-group date'>
                                        <input type="text" class="form-control input-small" name="FM_dt_sizeapp" value="" placeholder="dd-mm-yyyy"/>
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            
                        </div>
                         <div class="col-lg-4">
                            Exp PP
                             <div class='input-group date'>
                                         <input type="text" class="form-control input-small" name="FM_dt_ppexp" value="" placeholder="dd-mm-yyyy"/>
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                           
                        </div>
                         <div class="col-lg-4 pad-r30">
                            PP app
                           
                              <div class='input-group date'>
                                         <input type="text" class="form-control input-small" name="FM_dt_ppapp" value="" placeholder="dd-mm-yyyy"/>
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                        </div>

                         </div>

                        <div class="row">
                         <div class="col-lg-4">
                            Exp SHS
                             <div class='input-group date'>
                                         <input type="text" class="form-control input-small" name="FM_dt_shsexp" value="" placeholder="dd-mm-yyyy"/>
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            
                        </div>
                         <div class="col-lg-4 pad-r30">
                            SHS app
                            
                              <div class='input-group date'>
                                        <input type="text" class="form-control input-small" name="FM_dt_shsapp" value="" placeholder="dd-mm-yyyy"/>
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                        </div>

                         </div>

                         <div class="row">
                              <div class="col-lg-10">&nbsp;</div>
                                 <div class="col-lg-2 pad-r30">
                                <input type="submit" value="Submit" />
                            </div>
                          </div>
                       
               
                </section>

                     </form> <!-- bulk -->



                               
      
                    </div><!-- class="well well-white" -->
             </section>
                  </div>
                             </div>
                    
                      
                       

                       

                </div>

    <%end if %>
   


            <%
               
                if media = "exp" then
                strExpTxt = replace(strExpTxt, "xx99123sy#z", vbcrlf)
	
	         
	            call TimeOutVersion()
	
	            filnavnDato = year(now)&"_"&month(now)& "_"&day(now)
	            filnavnKlok = "_"&datepart("h", now)&"_"&datepart("n", now)&"_"&datepart("s", now)
	
				Set objFSO = server.createobject("Scripting.FileSystemObject")
				
                'response.write request.servervariables("PATH_TRANSLATED")
                'response.flush
                                                                  
				if request.servervariables("PATH_TRANSLATED") = "C:\www\timeout_xp\wwwroot\ver2_1\to_2015\job_nt.asp" then
					Set objNewFile = objFSO.createTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\ordersexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("c:\www\timeout_xp\wwwroot\ver2_1\inc\log\data\ordersexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				else
					Set objNewFile = objFSO.createTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ordersexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", True, False)
					Set objNewFile = nothing
					Set objF = objFSO.OpenTextFile("d:\webserver\wwwroot\timeout_xp\wwwroot\"& toVer &"\inc\log\data\ordersexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv", 8)
				end if
				
				
				
				file = "ordersexp_"&filnavnDato&"_"&filnavnKlok&"_"&lto&".csv"
				
                
                strOskrifter = "Status;Buyer;Buyer PO no.;Supplier;Sup. Inv. no.;Product grp.;Style;Order No.;Order type;"
                
                if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 
                strOskrifter = strOskrifter & "Order date;Destination;Collection;Sales Rep.;ETD Buyer;Actual ETD;"
                end if
                 if cint(rapporttype) = 3 then 
                strOskrifter = strOskrifter & "ETD Suppl.;"
                end if
                  if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then 
                strOskrifter = strOskrifter & "Order Qty.;Shipped Qty.;"
                end if
                
                if cint(rapporttype) = 1 then 
                strOskrifter = strOskrifter & "Proto DL;Photo Buyer DL;Photo Suppl. DL;SMS Buyer DL;SMS Suppl. DL;PP App;SHS App;ETD Buyer;ETD Suppl.;"
                end if
                
               
                 
                if cint(rapporttype) = 3 then
                strOskrifter = strOskrifter & "Cost Price PC;Val;"
                end if
                if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then
                strOskrifter = strOskrifter & "Sales Price PC;Val;"
                end if
                 if cint(rapporttype) = 3 then
                strOskrifter = strOskrifter & "Commision PC;Profit PC;Total Cost Price;Val;"
                end if
                  if cint(rapporttype) = 0 OR cint(rapporttype) = 3 then
                strOskrifter = strOskrifter & "Total Sales Price;Val;"
                end if
                 if cint(rapporttype) = 3 then
                strOskrifter = strOskrifter & "Total Profit;Val;Total Profit %;"
                end if
                
				objF.WriteLine(strOskrifter)
				objF.WriteLine(strExpTxt)
				objF.close
                response.redirect "../inc/log/data/"&file
                %>
                 
                 <!--
                <div style="position:absolute; left:90px; top:100px; width:400px; padding:20px; background-color:#FFFFFF; border:0px;">
                       <img src="../ill/outzource_logo_200.gif" /><br /><br /><br />
	              <a href="../inc/log/data/<%=file%>" class=vmenu target="_blank" onClick="Javascript:window.close()">Your CSV. file is ready >></a>
                -->

                   

                   
                        </div>
                    
                <%
                end if'media
                if media = "print" then
                Response.Write("<script language=""JavaScript"">window.print();</script>")
                end if
    
 end select %>



<script type="text/javascript">


$(".picmodal").click(function() {

    var modalid = this.id
    var idlngt = modalid.length
    var idtrim = modalid.slice(6, idlngt)

    //var modalidtxt = $("#myModal_" + idtrim);
    var modal = document.getElementById('myModal_' + idtrim);

    modal.style.display = "block";

    window.onclick = function (event) {
        if (event.target == modal) {
            modal.style.display = "none";
        }
    }

});


</script>











<!--#include file="../inc/regular/footer_inc.asp"-->