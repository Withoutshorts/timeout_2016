using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data.Odbc;
using System.Data;

public partial class importer_omk : System.Web.UI.Page
{
    const string PATH2UPLOAD = "~/inc/excelUpload/";
    public string importtype = "";

    protected void Page_Load(object sender, EventArgs e)
    {
        importtype = Request["importtype"];


        // feltnr1navn.Text = "No.:";
        //  feltnr2navn.Text = "Name:";
        // feltnr3navn.Text = "Email:";
        // feltnr4navn.Text = "Norm time:";
        // feltnr5navn.Text = "Employment date:";
        // feltnr6navn.Text = "Termination date:";
        // feltnr7navn.Text = "Blocked";
        // feltnr8navn.Text = "Expence vendor no.:";





    }


    protected void btnUpload_Click(object sender, EventArgs e)
    {
        try
        {
            ozUploadFileOmk service = new ozUploadFileOmk();
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

                ozUploadFileOmk service = new ozUploadFileOmk();
                String path = Server.MapPath(PATH2UPLOAD);
                string[] headers = GetExcelHeaderList();
                int[] intHeaders = GetExcelIntHeaderList();

                int countIgnore = 0;
                int countSent = 0;

                //string init = service.GetInit(Request["mid"], Request["lto"]);
              
                string strRet = string.Empty;

              

                //changed back to this query instead of the one below on purpose 01-17-2014 by Lei
                string strSelect = "SELECT init FROM medarbejdere where mid = " + Request["mid"];

                //string strSelect = "SELECT init, brugergruppe FROM medarbejdere where mid = " + midIn + " AND brugergruppe <> 3";

                try
                {
                    using (OdbcConnection connection = new OdbcConnection("driver={MySQL ODBC 3.51 Driver};server=194.150.108.154; Port=3306; uid=outzource; pwd=SKba200473; database=timeout_" + Request["lto"] + ";"))
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
                //lblStatus.Text += "<br>Init: " + init + "<br>SQL: " + strSelect;

                
                lblStatus.Text = "<br>Data blev indlæst korrekt.<br><br><a href='Javascript:window.close();'>[Luk denne side]</a><br><br>";

                string serviceReturn = string.Empty;
                string errorLine = string.Empty;

               

                if (init != string.Empty)
                    
                serviceReturn = service.Submit(path, txtFileName.Text, headers, intHeaders, Request["lto"], Request["editor"], ref countIgnore, ref countSent, init, ref errorLine, Request["mid"], Request["importtype"]);

                lblStatus.Text += serviceReturn.ToString();

                int rowsInserted = GetRowInserted(serviceReturn);
                int rowsFailed = countSent - rowsInserted;
                int rowTotal = countSent + countIgnore;

                string errorId = GetErrorId(serviceReturn);

                    if (rowsFailed == 0 && rowsInserted > 0)
                    {
                        if (countIgnore > 0)
                            lblStatus.Text = "Status message:<br><br> Filen indeholdt " + rowTotal + " linjer.<br>" + rowsInserted + " linjer indlæst korrekt.<br><br> " + countIgnore + " linjer er ignoreret. <br>" + errorLine;
                        else
                            lblStatus.Text = "Status message:<br><br> Filen indeholdt " + rowTotal + " linjer.<br>" + rowsInserted + " linjer indlæst korrekt. [LUK]";
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
            if (headerName.ToLower().Equals(nameIn)) //Contains()
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
            ozUploadFileOmk service = new ozUploadFileOmk();
            String path = Server.MapPath(PATH2UPLOAD);
            List<string> header = service.ReadHeader(path, txtFileName.Text, Request["lto"], Request["editor"], Request["importtype"]);

            if (header.Count < 8)
                lblUploadStatus.Text = "<b>Der skal være 8 kolonner i excel filen</b>";
            else
            {

                string importtype = Request["importtype"];


                    SelectDDL(ddlDato, "bogføringsdato", header, lblDato, "dato");
                    SelectDDL(ddlBesk, "beskrivelse", header, lblBesk, "beskrivelse");
                    SelectDDL(ddlKonto, "finanskontonr.", header, lblKonto, "konto");
                    SelectDDL(ddlJobId, "jobnr.", header, lblJobId, "jobid");
                    SelectDDL(ddlMinit, "medarb.", header, lblMinit, "medarb");
                    SelectDDL(ddlBelob, "beløb", header, lblBelob, "belob");
                    SelectDDL(ddlValuta, "valuta", header, lblValuta, "valuta");
                    SelectDDL(ddlAktnr, "løbenr.", header, lblAktnr, "lobenr");



            }
        }
        catch (Exception ex)
        {
            throw new Exception("Get header error: "+ex.Message);
        }
    }
   
    private string[] GetExcelHeaderList()
    {
        string[] strRets = { "", "", "", "", "", "", "", "" }; //,"","","",""

        try
        {
            Application.Lock();
            if (Application["dato"] != null)
                strRets[0] = Application["dato"].ToString();
            if (Application["beskrivelse"] != null)
                strRets[1] = Application["beskrivelse"].ToString();
            if (Application["konto"] != null)
                strRets[2] = Application["konto"].ToString();
            if (Application["jobid"] != null)
                strRets[3] = Application["jobid"].ToString();
            if (Application["medarb"] != null)
                strRets[4] = Application["medarb"].ToString();
            if (Application["belob"] != null)
                strRets[5] = Application["belob"].ToString();
            if (Application["valuta"] != null)
                strRets[6] = Application["valuta"].ToString();
            if (Application["lobenr"] != null)
                strRets[7] = Application["lobenr"].ToString();
             
             
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
        int[] intRets = { 0, 0, 0, 0, 0, 0, 0, 0 }; //, 0, 0, 0, 0

        try
        {

            if (ddlDato.SelectedIndex != -1)
                intRets[0] = ddlDato.SelectedIndex;
            if (ddlBesk.SelectedIndex != -1)
                intRets[1] = ddlBesk.SelectedIndex;
            if (ddlKonto.SelectedIndex != -1)
                intRets[2] = ddlKonto.SelectedIndex;
            if (ddlJobId.SelectedIndex != -1)
                intRets[3] = ddlJobId.SelectedIndex;
            if (ddlMinit.SelectedIndex != -1)
                intRets[4] = ddlMinit.SelectedIndex;
            if (ddlBelob.SelectedIndex != -1)
                intRets[5] = ddlBelob.SelectedIndex;
            if (ddlValuta.SelectedIndex != -1)
                intRets[6] = ddlValuta.SelectedIndex;
            if (ddlAktnr.SelectedIndex != -1)
                intRets[7] = ddlAktnr.SelectedIndex;

        }
        catch (Exception ex)
        {
            throw new Exception("Get excel integer header list error: " + ex.Message);
        }

        return intRets;
    }



    protected void ddlDato_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblDato.Text = "fundet kolonne: " + ddlDato.SelectedValue;
        Application["dato"] = ddlDato.SelectedValue;
    }

    protected void ddlJobId_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblJobId.Text = "fundet kolonne: " + ddlJobId.SelectedValue;
        Application["jobid"] = ddlJobId.SelectedValue;
    }

    protected void ddlBesk_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblBesk.Text = "fundet kolonne: " + ddlBesk.SelectedValue;
        Application["beskrivelse"] = ddlBesk.SelectedValue;
    }
   
    protected void ddlAktnr_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAktnr.Text = "fundet kolonne: " + ddlAktnr.SelectedValue;
        Application["aktnr"] = ddlAktnr.SelectedValue;
    }

  

   

    protected void ddlBelob_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblBelob.Text = "fundet kolonne: " + ddlBelob.SelectedValue;
        Application["belob"] = ddlBelob.SelectedValue;
    }

    protected void ddlKonto_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblKonto.Text = "fundet kolonne: " + ddlKonto.SelectedValue;
        Application["konto"] = ddlKonto.SelectedValue;
    }

    protected void ddlMinit_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMinit.Text = "fundet kolonne: " + ddlMinit.SelectedValue;
        Application["medarbejderinit"] = ddlMinit.SelectedValue;
    }


    protected void ddlValuta_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblValuta.Text = "fundet kolonne: " + ddlValuta.SelectedValue;
        Application["valuta"] = ddlValuta.SelectedValue;
    }






    protected void ddlDato_DataBound(object sender, EventArgs e)
    {
        lblDato.Text = "fundet kolonne: " + ddlDato.SelectedValue;
        Application["dato"] = ddlDato.SelectedValue;
    }




    protected void ddlJobId_DataBound(object sender, EventArgs e)
    {
        lblJobId.Text = "fundet kolonne: " + ddlJobId.SelectedValue;
        Application["jobid"] = ddlJobId.SelectedValue;
    }

    protected void ddlBesk_DataBound(object sender, EventArgs e)
    {
        lblBesk.Text = "fundet kolonne: " + ddlBesk.SelectedValue;
        Application["beskrivelse"] = ddlBesk.SelectedValue;
    }
    
   


    protected void ddlBelob_DataBound(object sender, EventArgs e)
    {
        lblBelob.Text = "fundet kolonne: " + ddlBelob.SelectedValue;
        Application["belob"] = ddlBelob.SelectedValue;
    }

    protected void ddlKonto_DataBound(object sender, EventArgs e)
    {
        lblKonto.Text = "fundet kolonne: " + ddlKonto.SelectedValue;
        Application["konto"] = ddlKonto.SelectedValue;
    }


    protected void ddlMinit_DataBound(object sender, EventArgs e)
    {
        lblMinit.Text = "fundet kolonne: " + ddlMinit.SelectedValue;
        Application["medarbejderinit"] = ddlMinit.SelectedValue;
    }

    protected void ddlValuta_DataBound(object sender, EventArgs e)
    {
        lblValuta.Text = "fundet kolonne: " + ddlValuta.SelectedValue;
        Application["valuta"] = ddlValuta.SelectedValue;
    }

    protected void ddlAktnr_DataBound(object sender, EventArgs e)
    {
        lblAktnr.Text = "fundet kolonne: " + ddlAktnr.SelectedValue;
        Application["aktnr"] = ddlAktnr.SelectedValue;
    }


}