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
<!--@ Page Language="C#" AutoEventWireup="true" CodeFile="importer_job_wilke.cs" Inherits="importer_job_wilke" 

    <!--@  Import Namespace="System.Globalization" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<script runat="server" language="vbscript">


    Public minit As String

    Sub Page_load()




        'meMTxt.Text = "HENT DATA"

        'Response.Write("<br>No (INIT): " + Request("no"))
        minit = Request("no")

        Call hentData()



    End Sub

    'Sub hentData(minit As String)
    Sub hentData()








        Dim med_navn, med_email, med_init, med_importtype, lto, med_mansat As String
        Dim med_expvendorno, med_costcenter, med_linemanager, med_countrycode, med_weblang As String

        Dim med_ansatdato, med_opsagtdato As Date
        Dim med_normtid As String
        'Dim timer As String
        'Dim jobnr As String
        'Dim jobnavn As String
        'Dim tdato As String
        'Dim taktivitetnavn, extsysid As String
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
        'Dim CallWebServiceTIA As New WebReferenceNAVTia_res.Resources()
        Dim CallWebServiceTIA As New WebReferenceNAVTia_res.Resources_Service

        'CallWebServiceTIA.UseDefaultCredentials = True
        'CallWebServiceTIA.Credentials = New NetworkCredential("xxxx", "xxxx", "xxxx")
        CallWebServiceTIA.Credentials = New System.Net.NetworkCredential(”tiademo”, ”Monday2017”, ”DEVX01”)
        'CallWebServiceTIA.PreAuthenticate = True

        Dim fetchSize As Integer = 10
        'Dim bookmarkKey As String = null


        Dim filter As New WebReferenceNAVTia_res.Resources_Filter
        'WebReferenceNAVTia_res.Resources_Filter
        filter.Field = WebReferenceNAVTia_res.Resources_Fields.No
        filter.Criteria = (minit)
        'filter.Criteria.





        Dim filters() As WebReferenceNAVTia_res.Resources_Filter = New WebReferenceNAVTia_res.Resources_Filter(0) {filter}

        Dim names As Array = CallWebServiceTIA.ReadMultiple(filters, Nothing, fetchSize)



        'Print the collected data.
        'For Each names As String In Array

        'meMTxt.Text = Name.ToString

        'Next
        'Dim namevalue = WebReferenceNAVTia_res.Resources_Fields.Name
        'Dim newRetval As String = WebReferenceNAVTia_res.Resources_Fields.Name.ToString()

        'Dim row As DataRow
        Dim t As Integer = 1

        'Print the collected data.
        For Each Name As Object In names

            'Response.Write("HEJ<br>")
            'meMTxt.Text = WebReferenceNAVTia_res.Resources_Fields.Name.ToString
            'allNames += allNames + "; " + WebReferenceNAVTia_res.Resources_Fields.Name.ToString()


            med_importtype = "1"
            med_init = Name.No.ToString()
            med_navn = Name.Name.ToString()
            med_email = "test@outzource.dk" 'Name.E_Mail.ToString()
            med_ansatdato = Name.Employment_Date.ToString()
            med_opsagtdato = Name.Termination_Date.ToString()
            med_normtid = "10"
            med_mansat = Name.Blocked.ToString()


            med_expvendorno = "1"
            med_costcenter = "1"
            med_linemanager = "0"
            med_countrycode = "DK"
            med_weblang = "1031"

            '" & Now.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) & "
            'med_ansatdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture)
            '" + med_opsagtdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture) + "

            '.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture)


            Dim strSQLmedinsTemp As String = "INSERT INTO med_import_temp_ds (dato, origin, Init, "
            strSQLmedinsTemp += "mnavn, "
            strSQLmedinsTemp += "email, "
            strSQLmedinsTemp += "normtid, "
            strSQLmedinsTemp += "ansatdato, "
            strSQLmedinsTemp += "opsagtdato, "
            strSQLmedinsTemp += "mansat, "
            strSQLmedinsTemp += "expvendorno, "
            strSQLmedinsTemp += "costcenter, "
            strSQLmedinsTemp += "linemanager, "
            strSQLmedinsTemp += "countrycode, "
            strSQLmedinsTemp += "weblang, lto, editor, overfort) "
            strSQLmedinsTemp += " VALUES ('2017-06-23', '914','" + med_init + "',"
            strSQLmedinsTemp += "'" + med_navn + "','" + med_email + "',"
            strSQLmedinsTemp += "'" + med_normtid + "',"
            strSQLmedinsTemp += "'2017-06-23',"
            strSQLmedinsTemp += "'2017-08-11','" + med_mansat + "', '" + med_expvendorno + "',"
            strSQLmedinsTemp += "'" + med_costcenter + "', '" + med_linemanager + "','" + med_countrycode + "', '" + med_weblang + "',"
            strSQLmedinsTemp += "'tia','Timeout - ImportMedService',0)"

            'Return strSQLaktins

            objCmd2 = New OdbcCommand(strSQLmedinsTemp, objConn)
            objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

            antalRecords = antalRecords + 1


            t = t + 1


            allNames += allNames + "; " + Name.Name.ToString()
        Next

        meMTxt.Text = allNames


        'newRetval = names.GetValue

        'meMTxt.Text = newRetval

        'WebReferenceNAVTia_res.Resources
        'For Each Name As String In names
        'allNames += allNames + "; " + Name
        'meMTxt.Text = Name
        'Next

        'meMTxt.Text = WebReferenceNAVTia_res.Resources[].

        'meMTxt.Text = CallWebServiceTIA.ToString 'names.ToString 'allNames







        If t > 0 Then

            'datasrc.Text = taktivitetnavnLst

            Dim Table1 As DataTable
            Table1 = New DataTable("tb_to_var")



            Dim column1 As DataColumn = New DataColumn("lto")
            column1.DataType = System.Type.GetType("System.String")

            Table1.Columns.Add(column1)

            Dim row As DataRow

            row = Table1.NewRow()
            row("lto") = "tia"

            'Add row
            Table1.Rows.Add(row)

            Dim ds As DataSet = New DataSet("ds")
            ds.Tables.Add(Table1)

            Dim CallWebService As New dk_rack.outzource_timeout2_importmed_nav.oz_importmed_na()
            CallWebService.addmed(ds)

            Response.Write("1")

        Else

            Response.Write("2")
            datasrc.Text = "Der blev ikke fundet medarbejder data at indlæse i denne kørsel"

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






        




