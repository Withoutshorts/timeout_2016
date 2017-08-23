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

        If (jobno <> "") Then
            Call hentData(jobno)
        End If



    End Sub

    Sub hentData(ByVal jobno)









        Dim Jobs_kundenr, Jobs_kundenavn, Jobs_jobnavn, Jobs_jobnr, Jobs_status, lto, Jobs_importtype As String
        Dim Jobs_startdato As Date = "2002-01-01"
        Dim Jobs_slutdato As Date = "2002-01-01"
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


        '** Test ENViRoNMENT
        'Dim CallWebServiceTIA As New WebReferenceNAVTia_job.Jobs_Service
        'CallWebServiceTIA.UseDefaultCredentials = True
        'CallWebServiceTIA.Credentials = New NetworkCredential("xxxx", "xxxx", "xxxx")
        'CallWebServiceTIA.Credentials = New System.Net.NetworkCredential(”tiademo”, ”Monday2017”, ”DEVX01”)
        'CallWebServiceTIA.PreAuthenticate = True

        '** PROD ENViRoNMENT
        Dim CallWebServiceTIA As New WebReferenceNAVTiaProd_job.TimeOutJobs_Service
        CallWebServiceTIA.Credentials = New System.Net.NetworkCredential(”Timeout”, ”Tia2017!”, ”tia.local”)



        Dim fetchSize As Integer = 1
        'Dim bookmarkKey As String = null


        Dim filter As New WebReferenceNAVTiaProd_job.TimeOutJobs_Filter
        'WebReferenceNAVTiaProd_job.Jobs_Filter
        filter.Field = WebReferenceNAVTiaProd_job.TimeOutJobs_Fields.No

        'Response.Write("jobno: " + jobno)

        'jobno
        filter.Criteria = ("" & jobno & "")





        Dim filters() As WebReferenceNAVTiaProd_job.TimeOutJobs_Filter = New WebReferenceNAVTiaProd_job.TimeOutJobs_Filter(0) {filter}

        Dim names As Array = CallWebServiceTIA.ReadMultiple(filters, Nothing, fetchSize)


        'Dim row As DataRow
        Dim t As Integer = 1

        'Print the collected data.
        For Each Name As Object In names

            'Response.Write("HEJ<br>")
            'meMTxt.Text = WebReferenceNAVTiaProd_job.Jobs_Fields.Name.ToString
            'allNames += allNames + "; " + WebReferenceNAVTiaProd_job.Jobs_Fields.Name.ToString()


            Jobs_importtype = "1"
            Jobs_status = Name.Blocked.ToString()



            If String.IsNullOrEmpty(Name.Bill_to_Customer_No) <> True Then
                Jobs_kundenr = Name.Bill_to_Customer_No.ToString()
            Else
                Jobs_kundenr = "0"
            End If

            If String.IsNullOrEmpty(Name.Bill_to_Name) <> True Then
                Jobs_kundenavn = Name.Bill_to_Name.ToString()
            Else
                Jobs_kundenavn = "0"
            End If

            If String.IsNullOrEmpty(Name.Description) <> True Then
                Jobs_jobnavn = Name.Description.ToString()
            Else
                Jobs_jobnavn = "Project name missing in NAV"
            End If

            Jobs_jobnr = Name.No.ToString()


            If String.IsNullOrEmpty(Name.Starting_Date) <> True Then
                Jobs_startdato = Name.Starting_Date
            Else
                Jobs_startdato = "2002/01/01"
            End If

            If String.IsNullOrEmpty(Name.Ending_Date) <> True Then
                Jobs_slutdato = Name.Starting_Date
            Else
                Jobs_slutdato = "2002/01/01"
            End If

            If String.IsNullOrEmpty(Name.Project_Manager) <> True Then
                Jobs_jobansvarlig = Name.Project_Manager.ToString()
            Else
                Jobs_jobansvarlig = ""
            End If

            Jobs_risiko = 1
            Jobs_projektgruppe = 0

            '" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "
            'med_ansatdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture)
            '" + med_opsagtdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) + "

            '.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture)

            If (CDate(Jobs_startdato) < "2002-01-01") Then
                Jobs_startdato = "2002-01-01"
            End If

            If (CDate(Jobs_slutdato) < "2002-01-01") Then
                Jobs_slutdato = "2044-01-01"
            End If


            If (Jobs_kundenr <> "0" And Jobs_kundenavn <> "0") Then

                Dim strSQLjobinsTemp As String = "INSERT INTO job_import_temp (dato, editor, lto, origin, jobnr, "
                strSQLjobinsTemp += "jobnavn, "
                strSQLjobinsTemp += "jobstartdato, "
                strSQLjobinsTemp += "jobslutdato, "
                strSQLjobinsTemp += "jobans, "
                strSQLjobinsTemp += "kundenavn, "
                strSQLjobinsTemp += "kundenr, projgrp, beskrivelse) "
                strSQLjobinsTemp += " VALUES ('" & DateAndTime.Year(Now) & "-" & DateAndTime.Month(Now) & "-" & DateAndTime.Day(Now) & "','NAV import','" + lto + "','914','" + Jobs_jobnr + "', '" + Jobs_jobnavn + "',"
                strSQLjobinsTemp += "'" & DateAndTime.Year(Jobs_startdato) & "-" & DateAndTime.Month(Jobs_startdato) & "-" & DateAndTime.Day(Jobs_startdato) & "',"
                strSQLjobinsTemp += "'" & DateAndTime.Year(Jobs_slutdato) & "-" & DateAndTime.Month(Jobs_slutdato) & "-" & DateAndTime.Day(Jobs_slutdato) & "',"
                strSQLjobinsTemp += "'" + Jobs_jobansvarlig + "',"
                strSQLjobinsTemp += "'" + Jobs_kundenavn + "','" + Jobs_kundenr + "', '0', '" + Jobs_status + "')"

                'Return strSQLaktins

                objCmd2 = New OdbcCommand(strSQLjobinsTemp, objConn)
                objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

                antalRecords = antalRecords + 1


                t = t + 1

            End If

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
    <asp:Button runat="server" Text="Hent data" ID="bt"   />

        <!-- OnClick="hentData(0)"-->
    <h4>Reading Data from the connection
    <asp:Label ID="datasrc" runat="server"></asp:Label> to the DataGrid</h4>

    <asp:DataGrid ID="dgNameList" runat="server" /><br />
    </form>
</body>
</html>






        




