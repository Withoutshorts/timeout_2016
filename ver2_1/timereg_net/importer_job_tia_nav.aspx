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
<!--@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_Jobs_wilke.cs" Inherits="importer_Jobs_wilke" 

    <!--@  Import Namespace="System.Globalization" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<script runat="server" language="vbscript">



    Public jobno As String

    Sub Page_load()




        'meMTxt.Text = "HENT DATA"

        'Response.Write("<br>No (INIT): " + Request("no"))

        Dim jobno As String = Request("no")
        Call hentData()



    End Sub

    Sub hentData()









        Dim Jobs_kundenr, Jobs_kundenavn, Jobs_jobnavn, Jobs_jobnr, Jobs_status, lto, Jobs_importtype As String
        Dim Jobs_startdato, Jobs_slutdato As Date
        Dim Jobs_risiko, Jobs_jobansvarlig, Jobs_projektgruppe As String

        Jobs_risiko = 1
        Jobs_projektgruppe = 0



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
        'Dim CallWebServiceTIA As New WebReferenceNAVTia_job.Job()
        Dim CallWebServiceTIA As New WebReferenceNAVTia_job.Jobs_Service

        'CallWebServiceTIA.UseDefaultCredentials = True
        'CallWebServiceTIA.Credentials = New NetworkCredential("xxxx", "xxxx", "xxxx")
        CallWebServiceTIA.Credentials = New System.Net.NetworkCredential(”tiademo”, ”Monday2017”, ”DEVX01”)
        'CallWebServiceTIA.PreAuthenticate = True

        Dim fetchSize As Integer = 10
        'Dim bookmarkKey As String = null


        Dim filter As New WebReferenceNAVTia_job.Jobs_Filter
        'WebReferenceNAVTia_job.Jobs_Filter
        filter.Field = WebReferenceNAVTia_job.Jobs_Fields.No

        filter.Criteria = (jobno)
        'filter.Criteria.





        Dim filters() As WebReferenceNAVTia_job.Jobs_Filter = New WebReferenceNAVTia_job.Jobs_Filter(0) {filter}

        Dim names As Array = CallWebServiceTIA.ReadMultiple(filters, Nothing, fetchSize)


        'Dim row As DataRow
        Dim t As Integer = 1

        'Print the collected data.
        For Each Name As Object In names

            'Response.Write("HEJ<br>")
            'meMTxt.Text = WebReferenceNAVTia_job.Jobs_Fields.Name.ToString
            'allNames += allNames + "; " + WebReferenceNAVTia_job.Jobs_Fields.Name.ToString()


            Jobs_importtype = "1"
            Jobs_status = Name.Blocked.ToString()
            Jobs_kundenr = Name.Bill_to_Customer_No.ToString()
            Jobs_kundenavn = Name.Bill_to_Name.ToString()
            Jobs_jobnavn = Name.Description.ToString()
            Jobs_jobnr = Name.No.ToString()
            Jobs_startdato = Name.Starting_Date.ToString()
            Jobs_slutdato = Name.Ending_Date.ToString()
            Jobs_jobansvarlig = Name.Project_Manager.ToString()

            Jobs_risiko = 1
            Jobs_projektgruppe = 0

            '" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "
            'med_ansatdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture)
            '" + med_opsagtdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) + "

            '.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture)


            Dim strSQLjobinsTemp As String = "INSERT INTO job_import_temp (dato, editor, lto, origin, jobnr, "
            strSQLjobinsTemp += "jobnavn, "
            strSQLjobinsTemp += "jobstartdato, "
            strSQLjobinsTemp += "jobslutdato, "
            strSQLjobinsTemp += "jobans, "
            strSQLjobinsTemp += "kundenavn, "
            strSQLjobinsTemp += "kundenr, projgrp) "
            strSQLjobinsTemp += " VALUES ('2017-06-23','NAV import','" + lto + "','914','" + Jobs_jobnr + "', '" + Jobs_jobnavn + "',"
            strSQLjobinsTemp += "'2017-01-01',"
            strSQLjobinsTemp += "'2017-08-29',"
            strSQLjobinsTemp += "'" + Jobs_jobansvarlig + "',"
            strSQLjobinsTemp += "'" + Jobs_kundenavn + "','" + Jobs_kundenr + "', '0')"

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
            row("importtype") = "t1"

            'Add row
            Table1.Rows.Add(row)

            Dim ds As DataSet = New DataSet("ds")
            ds.Tables.Add(Table1)

            Dim CallWebService As New dk_rack.outzource_timeout2.oz_importjob2()
            CallWebService.createjob2(ds)

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






        




