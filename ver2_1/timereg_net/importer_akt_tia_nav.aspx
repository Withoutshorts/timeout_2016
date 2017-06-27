<%@ Page Language="vb" Debug="true" %>

<%@ Import Namespace="System.Web.Services.Description" %>

<%@  Import Namespace="Microsoft.VisualBasic" %>
<%@  Import Namespace="System.Data"%>
<%@  Import Namespace="System.Data.Odbc" %>
<%@  Import Namespace="System.Web.Services" %> 

<%@  Import Namespace="System.Diagnostics" %>
<%@  Import Namespace="System.Web.UI" %> 
<%@  Import Namespace="System.Web.UI.Webcontrols" %>
<%@  Import Namespace="System.Collections.Generic" %>
<!--@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_JobTasks_wilke.cs" Inherits="importer_JobTasks_wilke" 

    <!--@  Import Namespace="System.Globalization" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<script runat="server" language="vbscript">



    Public jobno As String

    Sub Page_load()




        'meMTxt.Text = "HENT DATA"

        'Response.Write("<br>No (INIT): " + Request("no"))

        Dim jobno As String = Request("jobno")
        Call hentData()



    End Sub

    Sub hentData()















        Dim akt_taskno, akt_navn, akt_jobtaskno, akt_jobnr, akt_type, akt_status, lto, akt_importtype, akt_fakturerbar As String
        Dim akt_projektgruppe As String


        akt_projektgruppe = 0



        Dim taktivitetnavnLst As String = ""
        Dim antalRecords As Integer
        'Dim jobstartdato, jobslutdato As String

        'Dim kategori As Integer


        lto = "tia"
        'meMTxt.Text = "HEJ SØREN"

        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=root;Password=;"
        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";User=to_outzource2;Password=SKba200473;"

        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        'Dim objCmd As OdbcCommand
        Dim objCmd2 As OdbcCommand
        'Dim objDataSet As New DataSet
        'Dim objDR As OdbcDataReader
        'Dim objDR2 As OdbcDataReader
        'Dim objDR3 As OdbcDataReader
        'Dim dt As DataTable
        'Dim dr As DataRow






        '**************************************************************************************************
        '*** HENTER med fra MED_IMPORT_TEMP '****
        '*** VIA NAV webservice
        '**************************************************************************************************

        Dim allNames As String = ""
        'Dim CallWebServiceTIA As New WebReferenceNAVTia_akt.Job()
        Dim CallWebServiceTIA As New WebReferenceNAvTia_akt.JobTasks_Service

        'CallWebServiceTIA.UseDefaultCredentials = True
        'CallWebServiceTIA.Credentials = New NetworkCredential("xxxx", "xxxx", "xxxx")
        CallWebServiceTIA.Credentials = New System.Net.NetworkCredential(”tiademo”, ”Monday2017”, ”DEVX01”)
        'CallWebServiceTIA.PreAuthenticate = True

        Dim fetchSize As Integer = 20
        'Dim bookmarkKey As String = null


        Dim filter As New WebReferenceNAvTia_akt.JobTasks_Filter
        'WebReferenceNAVTia_akt.JobTasks_Filter
        filter.Field = WebReferenceNAvTia_akt.JobTasks_Fields.Job_No
        'jobno
        filter.Criteria = ("DEV1030*")
        'filter.Criteria.





        Dim filters() As WebReferenceNAvTia_akt.JobTasks_Filter = New WebReferenceNAvTia_akt.JobTasks_Filter(0) {filter}

        Dim names As Array = CallWebServiceTIA.ReadMultiple(filters, Nothing, fetchSize)


        'Dim row As DataRow
        Dim t As Integer = 1

        'Print the collected data.
        For Each Name As Object In names

            'Response.Write("HEJ<br>")
            'meMTxt.Text = WebReferenceNAVTia_akt.JobTasks_Fields.Name.ToString
            'allNames += allNames + "; " + WebReferenceNAVTia_akt.JobTasks_Fields.Name.ToString()


            akt_importtype = "1"

            akt_jobnr = Name.Job_No.ToString()
            akt_taskno = Name.Job_Task_No.ToString()
            akt_jobtaskno = Name.Job_Task_No.ToString()
            akt_navn = Name.Description.ToString()
            akt_fakturerbar = "1"
            akt_status = Name.Blocked.ToString()
            akt_type = Name.Job_Task_Type.ToString()

            akt_projektgruppe = 0

            '" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "
            'med_ansatdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture)
            '" + med_opsagtdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) + "

            '.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture)


            Dim strSQLjobinsTemp As String = "INSERT INTO akt_import_temp (dato, editor, lto, origin, jobnr, "
            strSQLjobinsTemp += "aktnavn, aktnr, "
            strSQLjobinsTemp += "aktkonto, aktstatus, beskrivelse, akttype) "
            strSQLjobinsTemp += " VALUES ('2017-06-23','NAV import','" + lto + "','914','" + akt_jobnr + "',"
            strSQLjobinsTemp += "'" + akt_navn + "', '" + akt_jobnr + "" + akt_taskno + "', "
            strSQLjobinsTemp += "'" + akt_jobtaskno + "', '" + akt_status + "', '', '" + akt_type + "')"

            'Return strSQLaktins

            objCmd2 = New OdbcCommand(strSQLjobinsTemp, objConn)
            objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

            antalRecords = antalRecords + 1


            t = t + 1


            'allNames += allNames + "; " + Name.Name.ToString()
        Next

        'meMTxt.Text = allNames




        If t > 0 Then

            'datasrc.Text = taktivitetnavnLst

            Dim Table1 As DataTable
            Table1 = New DataTable("tb_to_var")



            Dim column1 As DataColumn = New DataColumn("lto")
            column1.DataType = System.Type.GetType("System.String")

            Table1.Columns.Add(column1)

            Dim column2 As DataColumn = New DataColumn("importtype")
            column2.DataType = System.Type.GetType("System.String")

            Table1.Columns.Add(column2)

            Dim row As DataRow

            row = Table1.NewRow()
            row("lto") = "tia"
            row("importtype") = "tia1"

            'Add row
            Table1.Rows.Add(row)

            Dim ds As DataSet = New DataSet("ds")
            ds.Tables.Add(Table1)

            Dim CallWebService As New dk_rack.outzource_timeout2_importakt.oz_importakt()
            CallWebService.addakt(ds)

            Response.Write("1")

        Else

            Response.Write("2")
            datasrc.Text = "Der blev ikke fundet job data at indlæse i denne kørsel"

        End If








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
    <asp:TextBox runat="server" ID="meMTxt" Style="width:600px; height:400px; vertical-align:top;">NAV Data:</asp:TextBox>
    </div>
    <asp:Button runat="server" Text="Hent data" ID="bt" OnClick="hentData"  />

    <h4>Reading Data from the connection
    <asp:Label ID="datasrc" runat="server"></asp:Label> to the DataGrid</h4>

    <asp:DataGrid ID="dgNameList" runat="server" /><br />
    </form>
</body>
</html>






        




