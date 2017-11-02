<%@ Page Language="vb" Debug="true" %>

<%@ Import Namespace="System.Web.Services.Description" %>

<%@  Import Namespace="Microsoft.VisualBasic" %>
<%@  Import Namespace="System.Data"%>
<%@  Import Namespace="System.Data.Odbc" %>
<%@  Import Namespace="System.Web.Services" %> 
<%@  Import Namespace="System.Diagnostics" %>
<%@  Import Namespace="System.Web.UI" %> 
<%@  Import Namespace="System.Web.UI.Webcontrols" %>
<!--@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_job_wilke.cs" Inherits="importer_job_wilke" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<script runat="server" language="vbscript">


    Sub Page_load()




        meMTxt.Text = "HENT DATA"

        datofrom_d.Text = Day(Now)
        datofrom_m.Text = Month(Now)
        datofrom_y.Text = Year(Now)
        datoto_d.Text = Day(Now)
        datoto_m.Text = Month(Now)
        datoto_y.Text = Year(Now)



    End Sub




    Function hentData()

        Dim ddf As String = Request.Form("datofrom_d")
        Dim dmf As String = Request.Form("datofrom_m")
        Dim dyf As String = Request.Form("datofrom_y")

        Dim ddt As String = Request.Form("datoto_d")
        Dim dmt As String = Request.Form("datoto_m")
        Dim dyt As String = Request.Form("datoto_y")

        meMTxt.Text = "Henter data.."


        datofrom_d.Text = ddf 'datofrom_d.Text.ToString
        datofrom_m.Text = datofrom_m.Text.ToString
        datofrom_y.Text = datofrom_y.Text.ToString
        datoto_d.Text = datoto_d.Text.ToString
        datoto_m.Text = datoto_m.Text.ToString
        datoto_y.Text = datoto_y.Text.ToString


        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=root;Password=;"
        Dim strConn As String = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=to_outzource2; pwd=SKba200473; database=timeout_epi2017;"


        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        Dim objDR2 As OdbcDataReader
        Dim objDR3 As OdbcDataReader
        Dim dt As DataTable
        Dim dr As DataRow

        Dim extsysid As Integer
        Dim origin As String
        Dim jobnr As String
        Dim med_init As String
        Dim timer As String
        Dim jobid As String
        Dim jobnavn As String
        Dim timeregdato As String
        Dim taktivitetnavn As String
        Dim taktivitetnavnLst As String = ""



        Dim fraDato_d As String = datofrom_d.Text.ToString
        Dim fraDato_m As String = datofrom_m.Text.ToString
        Dim fraDato_y As String = datofrom_y.Text.ToString
        Dim tilDato_d As String = datoto_d.Text.ToString
        Dim tilDato_m As String = datoto_m.Text.ToString
        Dim tilDato_y As String = datoto_y.Text.ToString


        'Dim strConn As String

        'importFrom = ds.Tables("CATI_TIME").Rows(t).Item("origin")
        'intCatiId = ds.Tables("CATI_TIME").Rows(t).Item("id")
        'ds.Tables("CATI_TIME").Rows(t).Item("medarbejderid")
        'intJobNr = ds.Tables("CATI_TIME").Rows(t).Item("jobid").ToString
        'dlbTimer = ds.Tables("CATI_TIME").Rows(t).Item("timer")
        'cdDato = ds.Tables("CATI_TIME").Rows(t).Item("dato")


        Dim Table1 As DataTable
        Table1 = New DataTable("CATI_TIME")

        Dim column1 As DataColumn = New DataColumn("id")
        column1.DataType = System.Type.GetType("System.Int32")

        Dim column2 As DataColumn = New DataColumn("origin")
        column2.DataType = System.Type.GetType("System.Int32")

        Dim column3 As DataColumn = New DataColumn("medarbejderid")
        column3.DataType = System.Type.GetType("System.String")

        Dim column4 As DataColumn = New DataColumn("jobid")
        column4.DataType = System.Type.GetType("System.String")

        Dim column5 As DataColumn = New DataColumn("timer")
        column5.DataType = System.Type.GetType("System.String")

        Dim column6 As DataColumn = New DataColumn("dato")
        column6.DataType = System.Type.GetType("System.String")




        Table1.Columns.Add(column1)
        Table1.Columns.Add(column2)
        Table1.Columns.Add(column3)
        Table1.Columns.Add(column4)
        Table1.Columns.Add(column5)
        Table1.Columns.Add(column6)




        Dim row As DataRow
        Dim t As Integer = 1
        'AND jobnr BETWEEN 7000 AND 72000
        Dim strSQLext As String = "SELECT id, origin, jobnr, timeregdato, med_init, timer, timeregdato, extsysid FROM timer_imp_err WHERE id <> 0 AND timeregdato BETWEEN '" & fraDato_y & "-" & fraDato_m & "-" & fraDato_d & "' AND '" & tilDato_y & "-" & tilDato_m & "-" & tilDato_d & "' AND origin = 3 ORDER BY id LIMIT 1000" 'AND errId <> 11 ' AND origin = " & importFrom
        objCmd = New OdbcCommand(strSQLext, objConn)
        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

        While objDR.Read() = True



            extsysid = objDR("extsysid")
            origin = objDR("origin")
            jobid = objDR("jobnr")
            med_init = objDR("med_init")
            timer = objDR("timer")
            timeregdato = objDR("timeregdato")

            'jobnavn = Replace(objDR("jobnavn"), "'", "")
            'tdato = objDR("tdato")
            'taktivitetnavn = objDR("taktivitetnavn")

            taktivitetnavnLst = taktivitetnavnLst & "<br>" & t & " " & med_init & " jobid: " & jobid & " timer: " & timer & " tdato: " & timeregdato




            'lto = objDR("lto")



            row = Table1.NewRow()
            row("origin") = origin
            row("id") = extsysid
            row("jobid") = jobid
            row("timer") = timer
            row("medarbejderid") = med_init
            row("dato") = timeregdato



            'Add row
            Table1.Rows.Add(row)



            t = t + 1



        End While


        objDR.Close()
        objConn.Close()


        datasrc.Text = taktivitetnavnLst



        Dim ds As DataSet = New DataSet("ds")
        ds.Tables.Add(Table1)


        Dim CallWebService As New dk_rack.catiws.CATI()
        CallWebService.upDateCatiTimer(ds)

        'Dim CallWebService As New importjob2_ds.oz_importjob2_ds()
        'CallWebService.createjob2_ds(ds)

        'Dim ws As cati_timeout.CATI = New cati_timeout.CATI
        'ws.upDateCatiTimer(ds)


        me2MTxt.Text = "Data indlæst igen fra: " & fraDato_d & "-" & fraDato_m & "-" & fraDato_y & " til " & tilDato_d & "-" & tilDato_m & "-" & tilDato_y & ".<br> Du kan lukke denne side ned igen, og refreshe import error siden." '+ strSQLext




    End Function


</script>
   
<!--#include file="../inc/regular/header_lysblaa_2015_min_inc.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>TimeOut import cata-timer</title>
</head>
<body class=" ">

    
      <div class="wrapper">
      <div class="content">
         
          <div class="container">
         
            <table cellpadding=0 cellspacing=5 border=0 width=100%><tr><td>
	<img src="../ill/outzource_logo_200.gif" />
	
	</td></tr></table>
              <div class="portlet">
                  <h3 class="portlet-title">TimeOut - Cati</h3>
                  <div class="portlet-body">

                   <!--<div id="div_jobid">DDD</div>-->




       <div class="well">

    <form id="form1" runat="server">
  
          <div class="row">

   
        <div class="col-lg-1">
       <br /> Fra dato: 
            </div>
       <div class="col-lg-1">Dag
        <asp:TextBox runat="server" class="form-control input-small" ID="datofrom_d"></asp:TextBox> 
            </div>
            <div class="col-lg-1">Måned
            <asp:TextBox runat="server" class="form-control input-small" ID="datofrom_m"></asp:TextBox>  
             </div>
            <div class="col-lg-1">År
             <asp:TextBox class="form-control input-small" runat="server" ID="datofrom_y"></asp:TextBox>
             </div>
      

       </div>

         <div class="row">
        <div class="col-lg-1">
            Til dato:
            </div>
        <div class="col-lg-1">    
        <asp:TextBox runat="server" class="form-control input-small" ID="datoto_d"></asp:TextBox> 
            </div>
         <div class="col-lg-1">
              <asp:TextBox runat="server" class="form-control input-small" ID="datoto_m"></asp:TextBox> 
             </div>
         <div class="col-lg-1">  
         <asp:TextBox class="form-control input-small" runat="server" ID="datoto_y"></asp:TextBox>
        </div>
        


    </div>

        <br /><br />
    <div class="row">
    <div class="col-lg-4">
    <asp:Button runat="server" class="btn btn-sm btn-success" Text="Gen-indlæs cati data" ID="bt" OnClick="hentData"  />
     </div>
     </div>

    <!--<asp:TextBox runat="server" ID="meid"></asp:TextBox>-->
    <!-- <asp:TextBox runat="server" ID="meMTxt"></asp:TextBox> -->
        <br /><br />
    <div class="row">

        <div class="col-lg-12">
    <asp:Label runat="server" ID="me2MTxt"></asp:Label>
       </div>
                  </div>

    <!--<h4>Reading Data from the connection to the DataGrid</h4>-->
    <asp:Label ID="datasrc" runat="server"></asp:Label> 

    <asp:DataGrid ID="dgNameList" runat="server" /><br />
    </form>
           </div> <!-- Well -->
        </div>
     </div>
   </div>   
</div>   
</div>

      
      
</body>
</html>






        




