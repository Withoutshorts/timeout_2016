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





    End Sub

    Sub hentData()



        meMTxt.Text = "HEJ SØREN"

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
        Dim t As integer = 1
        'AND jobnr BETWEEN 7000 AND 72000
        Dim strSQLext As String = "SELECT id, origin, jobnr, timeregdato, med_init, timer, timeregdato, extsysid FROM timer_imp_err WHERE id <> 0 AND timeregdato BETWEEN '2017-06-01' AND '2017-06-30' AND errId <> 11 ORDER BY timeregdato" 'AND origin = " & importFrom
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

            taktivitetnavnLst = taktivitetnavnLst & "<br>" & t & " " & med_init & " " & jobid & " timer: " & timer & " tdato: " & timeregdato




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








    End Sub





</script>
   


<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:TextBox runat="server" ID="meid"></asp:TextBox>
    <asp:TextBox runat="server" ID="meMTxt">SØREN</asp:TextBox>
    </div>
    <asp:Button runat="server" Text="Hent data" ID="bt" OnClick="hentData"  />

    <h4>Reading Data from the connection
    <asp:Label ID="datasrc" runat="server"></asp:Label> to the DataGrid</h4>

    <asp:DataGrid ID="dgNameList" runat="server" /><br />
    </form>
</body>
</html>






        




