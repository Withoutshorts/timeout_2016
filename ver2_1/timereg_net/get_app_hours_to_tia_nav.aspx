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


    Sub Page_load()




        'meMTxt.Text = "HENT DATA"
        Call hentData()
        Response.Write("<br>No (INIT): " + Request("no"))





    End Sub

    Sub hentData()








        'Dim med_navn, med_email, med_init, med_importtype, lto, med_mansat As String
        'Dim med_expvendorno, med_costcenter, med_linemanager, med_countrycode, med_weblang As String

        'Dim med_ansatdato, med_opsagtdato As Date
        'Dim med_normtid As String
        'Dim timer As String
        'Dim jobnr As String
        'Dim jobnavn As String
        'Dim tdato As String
        'Dim taktivitetnavn, extsysid As String
        'Dim taktivitetnavnLst As String = ""
        'Dim antalRecords As Integer
        'Dim jobstartdato, jobslutdato As String

        'Dim kategori As Integer
        Dim lto As String

        lto = "tia"
        'meMTxt.Text = "HEJ SØREN"

        'Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;Database=timeout_intranet;User=root;Password=;"
        Dim strConn As String = "Driver={MySQL ODBC 3.51 Driver};Server=194.150.108.154;Database=timeout_" & lto & ";User=to_outzource2;Password=SKba200473;"

        Dim objConn As OdbcConnection
        objConn = New OdbcConnection(strConn)
        objConn.Open()

        Dim objCmd As OdbcCommand
        'Dim objCmd2 As OdbcCommand
        'Dim objDataSet As New DataSet
        Dim objDR As OdbcDataReader
        'Dim objDR2 As OdbcDataReader
        'Dim objDR3 As OdbcDataReader
        'Dim dt As DataTable
        'Dim dr As DataRow







        '***************************************************************************************
        '**** VIA recordset
        '***************************************************************************************

        'Dim row As DataRow
        Dim t As Integer = 1


        'Dim Table1 As DataTable
        'Table1 = New DataTable("tb_to_var")

        'Dim column1 As DataColumn = New DataColumn("lto")
        'column1.DataType = System.Type.GetType("System.String")

        'Dim column2 As DataColumn = New DataColumn("importtype")
        'column2.DataType = System.Type.GetType("System.String")

        'Dim column3 As DataColumn = New DataColumn("init")
        'column3.DataType = System.Type.GetType("System.String")

        'Dim column4 As DataColumn = New DataColumn("mnavn")
        'column4.DataType = System.Type.GetType("System.String")

        'Dim column5 As DataColumn = New DataColumn("email")
        'column5.DataType = System.Type.GetType("System.String")


        Dim CallWebService As New WebReferenceNAVTia_sendhours.TimeOut
        CallWebService.Credentials = New System.Net.NetworkCredential(”tiademo”, ”Monday2017”, ”DEVX01”)



        ''AND jobnr BETWEEN 7000 AND 72000

        Dim strSQLmedins As String = "SELECT tid, timer, m.init, tdato, tjobnr, a.avarenr, timerkom, tmnr "
        strSQLmedins += "FROM timer AS t "
        strSQLmedins += "LEFT JOIN medarbejdere AS m ON (m.mid = t.tmnr) "
        strSQLmedins += "LEFT JOIN aktiviteter AS a ON (a.id = t.taktivitetid) "
        strSQLmedins += "WHERE tastedato > '2017-01-01' AND godkendtstatus = 1 "

        objCmd = New OdbcCommand(strSQLmedins, objConn)
        objDR = objCmd.ExecuteReader '(CommandBehavior.closeConnection)

        While objDR.Read() = True


            Response.Write("<br>" & objDR("tdato") & "," & objDR("tjobnr") & "," & objDR("avarenr") & "," & objDR("init") & "," & objDR("timer"))
            'Response.Write(objDR("tdato") & "," & objDR("tjobnr") & ",111111," & objDR("init") & "," & objDR("timer") & "," & objDR("timerkom") & "," & objDR("tid"))
            Response.Flush()
            'CallWebService.SendJournalData(objDR("tdato"), objDR("tjobnr"), "1234", objDR("init"), objDR("timer"), "AA", objDR("tid"))
            Try

                CallWebService.SendJournalData(objDR("tdato"), objDR("tjobnr"), "1234", objDR("init"), objDR("timer"), "AA", objDR("tid"))

            Catch ex As Exception

                Response.Write("FEJL" & ex.ToString())



            End Try

            'objDR("avarenr")

            'row = Table1.NewRow()
            'row("lto") = lto
            'row("importtype") = "1"
            'row("init") = objDR("init")
            'row("tdato") = objDR("tdato")
            'row("timer") = objDR("timer")
            'row("timerkom") = objDR("timerkom")

            'row("expvendorno") = objDR("expvendorno")
            'row("costcenter") = objDR("costcenter")
            'row("linemanager") = objDR("linemanager")
            'row("countrycode") = objDR("countrycode")
            'row("weblang") = objDR("weblang")

            'taktivitetnavnLst = taktivitetnavnLst & objDR("init") & "<br>"

            'Add row
            'Table1.Rows.Add(row)



            t = t + 1



        End While


        objDR.Close()
        objConn.Close()

        If t > 0 Then

            'datasrc.Text = taktivitetnavnLst

            ' Dim Table1 As DataTable
            'Table1 = New DataTable("tb_to_var")



            'Dim column1 As DataColumn = New DataColumn("lto")
            'column1.DataType = System.Type.GetType("System.String")

            'Table1.Columns.Add(column1)

            'Dim row As DataRow

            'row = Table1.NewRow()
            'row("lto") = "tia"

            ''Add row
            'Table1.Rows.Add(row)

            'Dim ds As DataSet = New DataSet("ds")
            'ds.Tables.Add(Table1)

            'Dim CallWebService As New dk_rack.outzource_timeout2_importmed_nav.oz_importmed_na()
            'CallWebService.addmed(ds)

            Response.Write("1")

        Else

            Response.Write("2")
            datasrc.Text = "Der opstod en fejl"

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






        




