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

        If (minit <> "") Then
            Call hentData(minit)
        End If



    End Sub

    'Sub hentData(minit As String)
    Sub hentData(ByVal minit)








        Dim med_navn, med_email, med_init, med_importtype, lto, med_mansat As String
        Dim med_expvendorno, med_costcenter, med_linemanager, med_countrycode, med_weblang As String

        Dim med_ansatdato As Date = "1/1/2002"
        Dim med_opsagtdato As Date = "1/1/2044"
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

        '** TEST ENViRoNMENT
        'Dim CallWebServiceTIA As New WebReferenceNAVTia_res.Resources_Service
        'CallWebServiceTIA.Credentials = New System.Net.NetworkCredential(”tiademo”, ”Monday2017”, ”DEVX01”)

        '** PROD ENViRoNMENT
        Dim CallWebServiceTIA As New WebReferenceNAVTiaProd_Res.TimeOutResources_Service
        CallWebServiceTIA.Credentials = New System.Net.NetworkCredential(”Timeout”, ”Tia2017!”, ”tia.local”)

        'CallWebServiceTIA.UseDefaultCredentials = True
        'CallWebServiceTIA.Credentials = New NetworkCredential("xxxx", "xxxx", "xxxx")

        'CallWebServiceTIA.PreAuthenticate = True

        Dim fetchSize As Integer = 1 '50 '1
        'Dim bookmarkKey As String = null


        Dim filter As New WebReferenceNAVTiaProd_Res.TimeOutResources_Filter
        'WebReferenceNAVTiaProd_Res.TimeOutResources_Filter
        filter.Field = WebReferenceNAVTiaProd_Res.TimeOutResources_Fields.No
        'filter.Criteria = ("X*")
        filter.Criteria = ("" + minit + "")
        'filter.Criteria.


        ' a date value in the string format specified:
        'Dim xmlDate As String = "07/15/2014 7:07:33 AM"


        Dim filters() As WebReferenceNAVTiaProd_Res.TimeOutResources_Filter = New WebReferenceNAVTiaProd_Res.TimeOutResources_Filter(0) {filter}

        Dim names As Array = CallWebServiceTIA.ReadMultiple(filters, Nothing, fetchSize)



        'Print the collected data.
        'For Each names As String In Array

        'meMTxt.Text = Name.ToString

        'Next
        'Dim namevalue = WebReferenceNAVTiaProd_Res.TimeOutResources_Fields.Name
        'Dim newRetval As String = WebReferenceNAVTiaProd_Res.TimeOutResources_Fields.Name.ToString()

        'Dim row As DataRow
        Dim t As Integer = 1

        'Print the collected data.
        For Each Name As Object In names

            'Response.Write("HEJ<br>")
            'meMTxt.Text = WebReferenceNAVTiaProd_Res.TimeOutResources_Fields.Name.ToString
            'allNames += allNames + "; " + WebReferenceNAVTiaProd_Res.TimeOutResources_Fields.Name.ToString()


            med_importtype = "1"
            If String.IsNullOrEmpty(Name.No) <> True Then
                med_init = Name.No.ToString()
            Else
                med_init = "-"
            End If

            If String.IsNullOrEmpty(Name.Name) <> True Then
                med_navn = Name.Name.ToString()
            Else
                med_navn = "-"
            End If

            med_navn = med_navn.Replace("'", "")
            med_init = med_init.Replace("'", "")

            If String.IsNullOrEmpty(Name.E_Mail) <> True Then
                med_email = Name.E_Mail.ToString() 'Name.Employment_Date.ToString + Name.Employment_Date.ToString '"test@outzource.dk" 'Name.E_Mail.ToString()
            Else
                med_email = ""
            End If

            If String.IsNullOrEmpty(Name.Employment_Date) <> True Then
                med_ansatdato = Name.Employment_Date
            Else
                med_ansatdato = "2001/01/01"
            End If

            If String.IsNullOrEmpty(Name.Termination_Date) <> True Then
                med_opsagtdato = Name.Termination_Date
            Else
                med_opsagtdato = "2001/01/01"
            End If





            ' create a DATE variable from that string in a known format:
            ' med_opsagtdato = DateAndTime.Year(med_opsagtdato) & " - " & DateAndTime.Month(med_opsagtdato) & " - " & DateAndTime.Day(med_opsagtdato)
            If String.IsNullOrEmpty(Name.Norm_Time) <> True Then
                med_normtid = Name.Norm_Time.ToString()
            Else
                med_normtid = "0"
            End If

            med_normtid = med_normtid.Replace(",", ".")


            med_mansat = Name.Blocked.ToString()

            If med_mansat = True Then 'If blocked = true
                med_mansat = 2
            Else
                med_mansat = 1
            End If

            med_expvendorno = "1"

            'If String.IsNullOrEmpty(Name.Resource_Group_No) <> True Then
            '    med_costcenter = Name.Resource_Group_No.ToString()
            'Else
            '    med_costcenter = "10"
            'End If

            If String.IsNullOrEmpty(Name.Global_Dimension_2_Code) <> True Then
                med_costcenter = Name.Global_Dimension_2_Code.ToString()
            Else
                med_costcenter = "10"
            End If


            If String.IsNullOrEmpty(Name.Line_Manager) <> True Then
                med_linemanager = Name.Line_Manager.ToString()
            Else
                med_linemanager = ""
            End If

            If String.IsNullOrEmpty(Name.Global_Dimension_1_Code) <> True Then
                med_countrycode = Name.Global_Dimension_1_Code.ToString() '"DK"
            Else
                med_countrycode = "DK"
            End If

            med_weblang = "1031"

            'med_opsagtdato = med_ansatdato.ToString("yyyy/MM/dd")
            'med_opsagtdato = med_opsagtdato.ToString("yyyy/MM/dd", Globalization.CultureInfo.InvariantCulture)

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
            strSQLmedinsTemp += " VALUES ('" & DateAndTime.Year(Now) & "-" & DateAndTime.Month(Now) & "-" & DateAndTime.Day(Now) & "', '914','" + med_init + "',"
            strSQLmedinsTemp += "'" + med_navn + "','" + med_email + "',"
            strSQLmedinsTemp += "'" + med_normtid + "',"
            strSQLmedinsTemp += "'" & DateAndTime.Year(med_ansatdato) & "-" & DateAndTime.Month(med_ansatdato) & "-" & DateAndTime.Day(med_ansatdato) & "',"
            strSQLmedinsTemp += "'" & DateAndTime.Year(med_opsagtdato) & "-" & DateAndTime.Month(med_opsagtdato) & "-" & DateAndTime.Day(med_opsagtdato) & "','" + med_mansat + "', '" + med_expvendorno + "',"
            strSQLmedinsTemp += "'" + med_costcenter + "', '" + med_linemanager + "','" + med_countrycode + "', '" + med_weblang + "',"
            strSQLmedinsTemp += "'tia','Timeout - ImportMedService',0)"

            'Return strSQLaktins

            objCmd2 = New OdbcCommand(strSQLmedinsTemp, objConn)
            objCmd2.ExecuteReader() '(CommandBehavior.closeConnection)

            antalRecords = antalRecords + 1


            t = t + 1

            If String.IsNullOrEmpty(Name.Name) <> True Then
                'allNames += allNames + "; " + Name.Name.ToString()
            End If
        Next

        meMTxt.Text = allNames


        'newRetval = names.GetValue

        'meMTxt.Text = newRetval

        'WebReferenceNAVTiaProd_Res.Resources
        'For Each Name As String In names
        'allNames += allNames + "; " + Name
        'meMTxt.Text = Name
        'Next

        'meMTxt.Text = WebReferenceNAVTiaProd_Res.Resources[].

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
    <asp:Button runat="server" Text="Hent data" ID="bt"/><!-- OnClick="hentData"   -->

    <h4>Reading Data from the connection
    <asp:Label ID="datasrc" runat="server"></asp:Label> to the DataGrid</h4>

    <asp:DataGrid ID="dgNameList" runat="server" /><br />
    </form>
</body>
</html>






        




