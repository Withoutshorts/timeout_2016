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


<%@  Import Namespace="TestUserToken.ActivityService"%>
<%@  Import Namespace="TestUserToken.AuthenticationService"%>
<%@  Import Namespace="TestUserToken.ProjectService"%>

<!--@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_Jobs_wilke.cs" Inherits="importer_Jobs_wilke" 

    <!--@  Import Namespace="System.Globalization" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<script runat="server" language="vbscript">



    Public jobno As String

    Sub Page_load()




        'meMTxt.Text = "HENT DATA"

        'Response.Write("<br>No (INIT): " + Request("no"))

        'Dim jobno As String = Request("no")

        'If (jobno <> "") Then
        Call hentData(jobno)
        'End If



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


        lto = "cflow"
        'meMTxt.Text = "HEJ SØREN"

        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=root;Password=;"
        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";User=to_outzource2;Password=SKba200473;"

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
        '*** CFLOW Primavera webservice
        '**************************************************************************************************

        Dim allNames As String = ""
        'Dim WebReferenceCflow As New WebReferenceNAVTia_job.Job()


        '** Test ENViRoNMENT
        'Dim WebReferenceCflow As New WebReferenceNAVTia_job.Jobs_Service
        'WebReferenceCflow.UseDefaultCredentials = True
        'WebReferenceCflow.Credentials = New NetworkCredential("xxxx", "xxxx", "xxxx")
        'WebReferenceCflow.Credentials = New System.Net.NetworkCredential(”wsuser”, ”wsuser”, ”DEVX01”)
        'WebReferenceCflow.PreAuthenticate = True

        '** PROD ENViRoNMENT

        Dim WebReferenceCflow As New WebReferencePrimaCflow_akt.ActivityService
        WebReferenceCflow.UseDefaultCredentials = True
        WebReferenceCflow.PreAuthenticate = True
            'WebReferenceCflow.Credentials = New AuthenticationManager.
            ' (”tiademo”, ”Monday2017”, ”DEVX01”)


            'AuthenticationServicePortTypeClient
            'PreAuthenticate((authenticationServiceEndpoint, userName, password);

            'WebReferenceCflow. = New System.Net.NetworkCredential(”Timeout”, ”Tia2017!”, ”tia.local”)

            'WebReferenceCflow.ReadActivities.

            'Dim fetchSize As Integer = 1
            'Dim bookmarkKey As String = null
            'Dim filterString As String = "ActivityFieldType.Name LIKE 'a%'"
            'Dim filter As New WebReferencePrimaCflow_akt.ReadActivities
            Dim filter As New WebReferencePrimaCflow_akt.ReadActivities

            'WebReferencePrimaCflow_akt.Jobs_Filter
            'Filter.Field = WebReferencePrimaCflow_akt.ReadActivities.
            'filter.Field = filter.Filter.

            'Response.Write("jobno: " + jobno)

            'jobno
            'Filter.Criteria = ("" & jobno & "")



            'Dim filters() As WebReferencePrimaCflow_akt.ReadActivities = New WebReferencePrimaCflow_akt.TimeOutJobs_Filter(0) {filter()}
            'Dim readActivities As Array = WebReferenceCflow.ReadActivities(filter)
            'Dim filter As String = "Project Code = {0} and Id ='{1}'"
            Dim readActivities As Array = WebReferenceCflow.ReadActivities(filter)
            'ActivityFieldType.Name 'DefaultActivityFields()

            'Dim readActivities As New WebReferencePrimaCflow_akt.ActivityFieldType

            'Dim names As Array = WebReferenceCflow.ReadActivities(WebReferencePrimaCflow_akt.ActivityFieldType.Name)
            'Dim names As Array = WebReferenceCflow.ReadActivities(WebReferencePrimaCflow_akt.ActivityFieldType.Name)



            'Dim names As Array 'WebReferencePrimaCflow_akt.ActivityFieldType.Name

            'readActivities.Filter = String.Format("Project Code = {0} and Id ='{1}'", readProject[0].Id, activityId);

            For Each activity As Object In readActivities

            'Dim aktNavn As String = WebReferencePrimaCflow_akt.ActivityFieldType.Name.ToString()
            'foreach(var activity In readActivityTest)
            '{
            '    Console.WriteLine(activity.Name);
            '}

            Dim aktNavn As String = activity.Name.ToString()

            Dim strSQLjobinsTemp As String = "INSERT INTO akt_import_temp (aktnavn) VALUES ('" + aktNavn + "')"

            'Return strSQLaktins

            objCmd2 = New OdbcCommand(strSQLjobinsTemp, objConn)
            objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)


        Next






        Dim t As Integer = 0
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

            'Dim CallWebService As New dk_rack.outzource_timeout2.oz_importjob2()
            'CallWebService.createjob2(ds)

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
    <asp:TextBox runat="server" ID="meMTxt" Style="width:600px; height:400px; vertical-align:top;">Primavera activity Data:</asp:TextBox>
    </div>
    <asp:Button runat="server" Text="Hent data" ID="bt"   />

        <!-- OnClick="hentData(0)"-->
    <h4>Reading Data from the connection
    <asp:Label ID="datasrc" runat="server"></asp:Label> to the DataGrid</h4>

    <asp:DataGrid ID="dgNameList" runat="server" /><br />
    </form>
</body>
</html>






        




