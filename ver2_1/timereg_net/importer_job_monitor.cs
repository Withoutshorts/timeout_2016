using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data.Odbc;
using System.Data;

public partial class importer_job_monitor : System.Web.UI.Page
{
    const string PATH2UPLOAD = "~/inc/excelUpload/";

    protected void Page_Load(object sender, EventArgs e)
    {

        //string importtype = Request["importtype"];
        //hdn_importtype.Value = importtype;


    }


    protected void btnUpload_Click(object sender, EventArgs e)
    {
        try
        {
            ozUploadFileJob service = new ozUploadFileJob();
            String path = Server.MapPath(PATH2UPLOAD);
            lblUploadStatus.Text = service.Upload(fu, path);
            txtFileName.Text = fu.FileName;

            GetHeader();
        }
        catch (Exception ex)
        {
            lblStatus.Text = ex.Message;
        }
    }

    protected void btnAdd_Click(object sender, EventArgs e)
    {
        try
        {
            if (Request["editor"] != null && Request["lto"] != null && Request["mid"] != null && txtFileName.Text != string.Empty)
            {
                lblStatus.Text = string.Empty;

                ozUploadFileJob service = new ozUploadFileJob();
                String path = Server.MapPath(PATH2UPLOAD);
                string[] headers = GetExcelHeaderList();
                int[] intHeaders = GetExcelIntHeaderList();

                int countIgnore = 0;
                int countSent = 0;

                //string init = service.GetInit(Request["mid"], Request["lto"]);

                string strRet = string.Empty;

                string lto = Request["lto"];


                //if (lto == "dencker") 
                //{

                //lto = "dencker_test";
                //}

                //changed back to this query instead of the one below on purpose 01-17-2014 by Lei
                string strSelect = "SELECT init FROM medarbejdere where mid = " + Request["mid"];

                //string strSelect = "SELECT init, brugergruppe FROM medarbejdere where mid = " + midIn + " AND brugergruppe <> 3";

                try
                {


                    string strConn = "driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=to_outzource2; pwd=SKba200473; database=timeout_" + lto + "";
                    using (OdbcConnection connection = new OdbcConnection(strConn))
                    {
                        OdbcCommand command = new OdbcCommand(strSelect, connection);

                        connection.Open();

                        //lblStatus.Text = "Loop:<br>";
                        // Execute the DataReader and access the data.
                        OdbcDataReader myReader = command.ExecuteReader();
                        while (myReader.Read())
                        {
                            strRet = myReader.GetValue(0).ToString();
                            //lblStatus.Text += "<br>medid: " + myReader.GetValue(0).ToString();
                        }

                        // Call Close when done reading.
                        myReader.Close();
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }

              
                string init = strRet;

                //lblStatus.Text += "<br>Path: " + path;
                //lblStatus.Text += "<br>txtFileName.Text: " + txtFileName.Text;
                //lblStatus.Text += "<br>Headers: " + headers + " intHeaders: " + intHeaders;
                //lblStatus.Text += "<br>Variable LTO: " + Request["lto"] + " Editor: " + Request["editor"] + " Mid:" + Request["mid"];
                //lblStatus.Text = "<br>Init: " + init + "<br>SQL: " + strSelect;
                //lblStatus.Text = "<br>Data blev indlæst korrekt.<br><br><a href='Javascript:window.close();'>[Luk denne side]</a><br><br>";

                string importtype = Request["importtype"];
                string serviceReturn = string.Empty;
               string errorLine = string.Empty;

               
               
               if (init != string.Empty)
                
               serviceReturn = service.Submit(path, txtFileName.Text, headers, intHeaders, Request["lto"], Request["editor"], ref countIgnore, ref countSent, init, ref errorLine, Request["mid"], importtype);

               lblStatus.Text += "INFO fra Webservice: "+ serviceReturn.ToString() + "<br>";

                       int rowsInserted = GetRowInserted(serviceReturn);
                       int rowsFailed = countSent - rowsInserted;
                       int rowTotal = countSent + countIgnore;

                       string errorId = GetErrorId(serviceReturn);

                           if (rowsFailed == 0 && rowsInserted > 0)
                           {
                               if (countIgnore > 0)
                                   lblStatus.Text = "Status message:<br><br> Filen indeholdt " + rowTotal + " linjer.<br>" + rowsInserted + " linjer indlæst korrekt.<br><br> " + countIgnore + " linjer er ignoreret. <br>" + errorLine;
                               else
                        //lblStatus.Text = "Status message:<br><br> Filen indeholdt " + rowTotal + " linjer.<br>" + rowsInserted + " linjer indlæst korrekt. [LUK]";
                        lblStatus.Text = "Status message:<br><br> Filen indeholdt " + rowTotal + " linjer.<br>" + rowsInserted + " linjer indlæst korrekt. <br><br><a href='Javascript:window.close();'>[Luk denne side]</a><br><br>";
                }
                           else
                           {
                               if (countIgnore > 0)
                                   lblStatus.Text = "Status message:<br><br> Filen indeholdt " + rowTotal + " linjer.<br>" + rowsInserted + " linjer indlæst korrekt.<br><br>" + countIgnore + " linjer er ignoreret. <br>" + errorLine + errorId;
                               else
                                   lblStatus.Text += "Status message:<br><br> Filen indeholdt " + rowTotal + " linjer.<br>" + rowsInserted + " linjer indlæst korrekt.<br><br>Errid: " + errorId; 
                           
                           }
                   }
                   else

                       lblStatus.Text = "Kan ikke finde editor, mid eller licens indehaver fra link i TimeOut.";
                
                    
                 }
                
        

                     
       
       catch (Exception ex)
       {
           if (ex.Message.Contains("læsning"))
               lblStatus.Text += "Du er nødt til at upload en ny fil.<br>" + ex.Message;
           else
               lblStatus.Text += "Server error: "+ex.Message;
                 
                 
            }


        }
   



    private string GetJobId(string msgIn)
    {
        string errorId = string.Empty;

        int indexJobid = msgIn.IndexOf("jobnr:");
        int indexRightEnd = msgIn.IndexOf(")");

        if (indexJobid != -1)
            errorId = msgIn.Substring(indexJobid + 6, indexRightEnd - indexJobid);

        return errorId;
    }

    private string GetErrorId(string msgIn)
    {
        string errorId = string.Empty;

        int indexErrid = msgIn.IndexOf("errid:");
        
        if (indexErrid != -1)
            errorId = msgIn.Substring(indexErrid + 6);

        return errorId;
    }

    private int GetRowInserted(string msgIn)
    {
        int intRet = 0;

        int indexLinje = msgIn.IndexOf("linje");
        int len = indexLinje - 8;
        if (len < 0)
            len = 0;        

        if (msgIn.Length > 7)
            intRet = int.Parse(msgIn.Substring(7, len));

        return intRet;
    }

    /// <summary>
    /// //Lei new updated 23-11-2012
    /// Ignore the same jobnr
    /// </summary>
    /// <param name="jobnrIn">original jobnr message</param>
    /// <returns>Distinct jobnr list</returns>
    private string distinctJobnr(string jobnrIn)
    {
        string strRet = jobnrIn;

        try
        {
            strRet = strRet.Substring(6);
            string[] numbers = strRet.Split(',');
            List<int> lstNumbers = new List<int>();
            if (numbers.Length > 1)
            {
                for (int i = 0; i < numbers.Length; i++)
                {
                    int num = int.Parse(numbers[i]);

                    if (!lstNumbers.Contains(num))
                        lstNumbers.Add(int.Parse(numbers[i]));
                }

                strRet = "jobnr:";
                foreach (int num in lstNumbers)
                {
                    strRet += num + ",";
                }
            }
            else
                strRet = string.Empty;
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }

        return strRet;
    }

    private void SelectDDL(DropDownList ddlIn, string nameIn, List<string> lstHeaderIn, Label lblIn, string appName)
    {
        ListItem li = new ListItem("", "");

        ddlIn.Items.Clear();
        ddlIn.Items.Add(li);
        ddlIn.DataSource = lstHeaderIn;
        ddlIn.DataBind();
        foreach (string headerName in lstHeaderIn)
        {
            if (headerName.ToLower().Contains(nameIn))
            {
                ddlIn.SelectedIndex = -1;
                ddlIn.SelectedValue = headerName;
                lblIn.Text = "fundet kolonne: " + headerName;
                Application[appName] = headerName;
                break;
            }
        }
        if (ddlIn.SelectedIndex == -1)
            li.Selected = true;
    }

    private void GetHeader()
    {
        try
        {
            ozUploadFileJob service = new ozUploadFileJob();
            String path = Server.MapPath(PATH2UPLOAD);
            List<string> header = service.ReadHeader(path, txtFileName.Text, Request["lto"], Request["editor"], Request["importtype"]);

            if (header.Count < 13)
                lblUploadStatus.Text = "<b>Der skal være 13 kolonner i excel filen</b>";
            else
            {
                //SelectDDL(ddlKundenavn, "kundeinfo", header, lblKundenavn, "kundenavn");
                //SelectDDL(ddlJobnavn, "jobnavn", header, lblJobnavn, "jobnavn");
                //SelectDDL(ddlJobId, "ordrenr.", header, lblJobId, "jobid");
                //SelectDDL(ddlAntal, "rap. ant.", header, lblAntal, "antal");
                //SelectDDL(ddlstDato, "startperiode", header, lblstDato, "stdato");
                //SelectDDL(ddlslDato, "afsluttet vare", header, lblslDato, "sldato");
                //SelectDDL(ddlSort, "op.", header, lblSort, "sort");
                //SelectDDL(ddlFomr, "prod.grp.", header, lblFomr, "fomr");
                //SelectDDL(ddlAktnavn, "op. beskrivelse", header, lblAktnavn, "aktnavn");
                //SelectDDL(ddlAktstdato, "start", header, lblAktstdato, "aktstdato");
                //SelectDDL(ddlAktsldato, "afsluttet", header, lblAktsldato, "aktsldato");
                //SelectDDL(ddlAktvarenr, "rapp.nr.", header, lblAktvarenr, "aktvarenr");


                SelectDDL(ddlKundenavn, "kund_info", header, lblKundenavn, "kundenavn");
                SelectDDL(ddlJobnavn, "art_ben", header, lblJobnavn, "jobnavn");
                SelectDDL(ddlJobId, "toh_ordernr", header, lblJobId, "jobid");
                SelectDDL(ddlAntal, "too_antal_pl", header, lblAntal, "antal");
                SelectDDL(ddlstDato, "tos_start", header, lblstDato, "stdato");
                SelectDDL(ddlslDato, "tos_slut", header, lblslDato, "sldato");
                SelectDDL(ddlSort, "too_op", header, lblSort, "sort");
                SelectDDL(ddlFomr, "pg_pgrkod", header, lblFomr, "fomr");
                SelectDDL(ddlAktnavn, "too_ben", header, lblAktnavn, "aktnavn");
                SelectDDL(ddlAktstdato, "omp_start", header, lblAktstdato, "aktstdato"); //too_start
                SelectDDL(ddlAktsldato, "omp_slut", header, lblAktsldato, "aktsldato"); //too_slut
                SelectDDL(ddlAktvarenr, "too_rappnr", header, lblAktvarenr, "aktvarenr");
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Get header error: "+ex.Message);
        }
    }
   
    private string[] GetExcelHeaderList()
    {
        string[] strRets = { "", "", "", "", "", "", "", "", "", "", "", "" }; //,"","","",""

        try
        {
            Application.Lock();
           
            if (Application["jobnavn"] != null)
                strRets[0] = Application["jobnavn"].ToString();
            if (Application["jobid"] != null)
                strRets[1] = Application["jobid"].ToString();
            if (Application["kundenavn"] != null)
                strRets[2] = Application["kundenavn"].ToString();
            if (Application["stdato"] != null)
                strRets[3] = Application["aktstdato"].ToString();
            if (Application["sldato"] != null)
                strRets[4] = Application["sldato"].ToString();
            if (Application["antal"] != null)
                strRets[5] = Application["antal"].ToString();
            if (Application["sort"] != null)
                strRets[6] = Application["sort"].ToString();
            if (Application["fomr"] != null)
                strRets[7] = Application["fomr"].ToString();
           
            if (Application["aktstdato"] != null)
                strRets[8] = Application["aktstdato"].ToString();
            if (Application["aktsldato"] != null)
                strRets[9] = Application["aktsldato"].ToString();
            if (Application["aktnavn"] != null)
                strRets[10] = Application["aktnavn"].ToString();
            if (Application["aktvarenr"] != null)
                strRets[11] = Application["aktvarenr"].ToString();

            Application.UnLock();
        }
        catch (Exception ex)
        {
            throw new Exception("Get excel header list error: "+ex.Message);
        }

        return strRets;
    }

    private int[] GetExcelIntHeaderList()
    {
        int[] intRets = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 }; //, 0, 0, 0, 0

        try
        {

            
            if (ddlJobnavn.SelectedIndex != -1)
                intRets[0] = ddlJobnavn.SelectedIndex;
            if (ddlJobId.SelectedIndex != -1)
                intRets[1] = ddlJobId.SelectedIndex;
            if (ddlKundenavn.SelectedIndex != -1)
                intRets[2] = ddlKundenavn.SelectedIndex;
            if (ddlstDato.SelectedIndex != -1)
                intRets[3] = ddlstDato.SelectedIndex;
            if (ddlslDato.SelectedIndex != -1)
                intRets[4] = ddlslDato.SelectedIndex;

            if (ddlAntal.SelectedIndex != -1)
                intRets[5] = ddlAntal.SelectedIndex;
            if (ddlSort.SelectedIndex != -1)
                intRets[6] = ddlSort.SelectedIndex;
            if (ddlFomr.SelectedIndex != -1)
                intRets[7] = ddlFomr.SelectedIndex;
            if (ddlAktstdato.SelectedIndex != -1)
                intRets[8] = ddlAktstdato.SelectedIndex;
            if (ddlAktsldato.SelectedIndex != -1)
                intRets[9] = ddlAktsldato.SelectedIndex;
            if (ddlAktnavn.SelectedIndex != -1)
                intRets[10] = ddlAktnavn.SelectedIndex;
            if (ddlAktvarenr.SelectedIndex != -1)
                intRets[11] = ddlAktvarenr.SelectedIndex;
            


        }
        catch (Exception ex)
        {
            throw new Exception("Get excel integer header list error: " + ex.Message);
        }

        return intRets;
    }


    protected void ddlKundenavn_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblKundenavn.Text = "fundet kolonne: " + ddlKundenavn.SelectedValue;
        Application["kundenavn"] = ddlKundenavn.SelectedValue;
    }
    protected void ddlJobnavn_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblJobnavn.Text = "fundet kolonne: " + ddlJobnavn.SelectedValue;
        Application["jobnavn"] = ddlJobnavn.SelectedValue;
    }
    protected void ddlJobId_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblJobId.Text = "fundet kolonne: " + ddlJobId.SelectedValue;
        Application["jobid"] = ddlJobId.SelectedValue;
    }

    
    protected void ddlAntal_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAntal.Text = "fundet kolonne: " + ddlAntal.SelectedValue;
        Application["antal"] = ddlAntal.SelectedValue;
    }
    protected void ddlstDato_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblstDato.Text = "fundet kolonne: " + ddlstDato.SelectedValue;
        Application["stdato"] = ddlstDato.SelectedValue;
    }

    protected void ddlslDato_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblslDato.Text = "fundet kolonne: " + ddlslDato.SelectedValue;
        Application["sldato"] = ddlslDato.SelectedValue;
    }


    protected void ddlSort_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblSort.Text = "fundet kolonne: " + ddlSort.SelectedValue;
        Application["sort"] = ddlSort.SelectedValue;
    }

    protected void ddlFomr_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblFomr.Text = "fundet kolonne: " + ddlFomr.SelectedValue;
        Application["fomr"] = ddlFomr.SelectedValue;
    }

    protected void ddlAktnavn_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAktnavn.Text = "fundet kolonne: " + ddlAktnavn.SelectedValue;
        Application["aktnavn"] = ddlAktnavn.SelectedValue;
    }

    protected void ddlAktstdato_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAktstdato.Text = "fundet kolonne: " + ddlAktstdato.SelectedValue;
        Application["aktstdato"] = ddlAktstdato.SelectedValue;
    }

    protected void ddlAktsldato_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAktsldato.Text = "fundet kolonne: " + ddlAktsldato.SelectedValue;
        Application["aktsldato"] = ddlAktsldato.SelectedValue;
    }

    protected void ddlAktvarenr_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAktvarenr.Text = "fundet kolonne: " + ddlAktvarenr.SelectedValue;
        Application["aktvarenr"] = ddlAktvarenr.SelectedValue;
    }

    /// <summary>
    /// /////////////////////////////////////////////////////////////////////////////
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>

    protected void ddlKundenavn_DataBound(object sender, EventArgs e)
    {
        lblKundenavn.Text = "fundet kolonne: " + ddlKundenavn.SelectedValue;
        Application["kundenavn"] = ddlKundenavn.SelectedValue;
    }
    protected void ddlJobnavn_DataBound(object sender, EventArgs e)
    {
        lblJobnavn.Text = "fundet kolonne: " + ddlJobnavn.SelectedValue;
        Application["jobnavn"] = ddlJobnavn.SelectedValue;
    }
    protected void ddlJobId_DataBound(object sender, EventArgs e)
    {
        lblJobId.Text = "fundet kolonne: " + ddlJobId.SelectedValue;
        Application["jobid"] = ddlJobId.SelectedValue;
    }
    
    
    protected void ddlAntal_DataBound(object sender, EventArgs e)
    {
        lblAntal.Text = "fundet kolonne: " + ddlAntal.SelectedValue;
        Application["antal"] = ddlAntal.SelectedValue;
    }


    protected void ddlstDato_DataBound(object sender, EventArgs e)
    {
        lblstDato.Text = "fundet kolonne: " + ddlstDato.SelectedValue;
        Application["stDato"] = ddlstDato.SelectedValue;
    }

    protected void ddlslDato_DataBound(object sender, EventArgs e)
    {
        lblslDato.Text = "fundet kolonne: " + ddlslDato.SelectedValue;
        Application["slDato"] = ddlslDato.SelectedValue;
    }
    
    protected void ddlSort_DataBound(object sender, EventArgs e)
    {
        lblSort.Text = "fundet kolonne: " + ddlSort.SelectedValue;
        Application["sort"] = ddlSort.SelectedValue;
    }

    protected void ddlFomr_DataBound(object sender, EventArgs e)
    {
        lblFomr.Text = "fundet kolonne: " + ddlFomr.SelectedValue;
        Application["fomr"] = ddlFomr.SelectedValue;
    }

    protected void ddlAktnavn_DataBound(object sender, EventArgs e)
    {
        lblAktnavn.Text = "fundet kolonne: " + ddlAktnavn.SelectedValue;
        Application["aktnavn"] = ddlAktnavn.SelectedValue;
    }


    protected void ddlAktstdato_DataBound(object sender, EventArgs e)
    {
        lblAktstdato.Text = "fundet kolonne: " + ddlAktstdato.SelectedValue;
        Application["aktstdato"] = ddlAktstdato.SelectedValue;
    }

    protected void ddlAktsldato_DataBound(object sender, EventArgs e)
    {
        lblAktsldato.Text = "fundet kolonne: " + ddlAktsldato.SelectedValue;
        Application["aktsldato"] = ddlAktsldato.SelectedValue;
    }

    protected void ddlAktvarenr_DataBound(object sender, EventArgs e)
    {
        lblAktvarenr.Text = "fundet kolonne: " + ddlAktvarenr.SelectedValue;
        Application["aktvarenr"] = ddlAktvarenr.SelectedValue;
    }

    

}