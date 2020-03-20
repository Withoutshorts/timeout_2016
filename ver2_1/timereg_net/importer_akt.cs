using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data.Odbc;
using System.Data;

public partial class importer_akt : System.Web.UI.Page
{
    const string PATH2UPLOAD = "~/inc/excelUpload/";
    public string importtype = "";
    public int clm = 8;

    protected void Page_Load(object sender, EventArgs e)
    {
        importtype = Request["importtype"];

      
        feltnr1navn.Text = "Job:";
        feltnr2navn.Text = "Beskrivelse:";
        feltnr3navn.Text = "Løbenr. (NAV):";
        feltnr4navn.Text = "Konto:";
        feltnr5navn.Text = "Type:";
        feltnr6navn.Text = "Timer/Stk.:";
        feltnr7navn.Text = "Stk. pris:";
        feltnr8navn.Text = "Beløb:";
        feltnr9navn.Text = "Dato:";


        if (importtype == "t2")
        {
            feltnr1navn.Text = "Project No.:";
            feltnr2navn.Text = "Description:";
            feltnr3navn.Text = "Task No.:";
            feltnr4navn.Text = "Blocked:";
            feltnr5navn.Text = "Ress. chargeable:";
            //feltnr6navn.Text = "Timer/Stk.:";
            //feltnr7navn.Text = "Stk. pris:";
            //feltnr8navn.Text = "Beløb:";

        }

        if (importtype == "cflow_prima")
        {
            feltnr1navn.Text = "Project No.:";
            feltnr2navn.Text = "Activity Id.:";
            feltnr3navn.Text = "Activity/Komp. desc.:";
            feltnr9navn.Text = "Start date:";
            //feltnr5navn.Text = "Ress. chargeable:";
            //feltnr6navn.Text = "Timer/Stk.:";
            //feltnr7navn.Text = "Stk. pris:";
            //feltnr8navn.Text = "Beløb:";

        }

    }


    protected void btnUpload_Click(object sender, EventArgs e)
    {
        try
        {
            ozUploadFileAkt service = new ozUploadFileAkt();
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

                ozUploadFileAkt service = new ozUploadFileAkt();
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
            ozUploadFileAkt service = new ozUploadFileAkt();
            String path = Server.MapPath(PATH2UPLOAD);
            List<string> header = service.ReadHeader(path, txtFileName.Text, Request["lto"], Request["editor"], Request["importtype"]);

            if (importtype == "t2" || importtype == "cflow_prima")
            {

                if (importtype == "t2")
                {
                    clm = 5;
                } else
                {
                    clm = 4;
                }
            }
            else
            {
                clm = 9;
            };

                if (header.Count < clm)
                lblUploadStatus.Text = "<b>Der skal være "+ clm + " kolonner i excel filen</b>";
            else
            {

                 importtype = Request["importtype"];

                if (importtype == "t2" || importtype == "cflow_prima") {

                    if (importtype == "t2")
                    {
                        SelectDDL(ddlLinjetype, "fakturerbar", header, lblLinjetype, "linjetype");
                        SelectDDL(ddlJobId, "job/project", header, lblJobId, "jobid");
                        SelectDDL(ddlAktnavn, "description", header, lblAktnavn, "aktnavn");
                        SelectDDL(ddlAktnr, "jobnr", header, lblAktnr, "aktnr");
                        SelectDDL(ddlKonto, "blocked", header, lblKonto, "konto");
                    }
                    else
                    {
                        
                        SelectDDL(ddlJobId, "project id", header, lblJobId, "jobid");
                        SelectDDL(ddlAktnr, "activity id", header, lblAktnr, "aktnr");
                        SelectDDL(ddlAktnavn, "wbs name", header, lblAktnavn, "aktnavn");
                        SelectDDL(ddlAktstDato, "actual start", header, lblAktstDato, "aktstdato");
                    }


                } else { 

                    SelectDDL(ddlLinjetype, "type", header, lblLinjetype, "linjetype");
                    SelectDDL(ddlJobId, "sagsnummer", header, lblJobId, "jobid");
                    SelectDDL(ddlAktnavn, "beskrivelse", header, lblAktnavn, "aktnavn");
                    SelectDDL(ddlAktnr, "nummer", header, lblAktnr, "aktnr");
                    SelectDDL(ddlKonto, "sagsopgavenummer", header, lblKonto, "konto");
                }




                if (importtype != "t2")
                {
                    SelectDDL(ddlAkttimer, "antal", header, lblAkttimer, "akttimer");
                    SelectDDL(ddlAkttpris, "kostpris", header, lblAkttpris, "akttpris");
                    SelectDDL(ddlAktsum, "kostbeloeb", header, lblAktsum, "aktsum");
                    SelectDDL(ddlAktstDato, "dato", header, lblAktstDato, "aktstdato");
                }
                
                
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Get header error: "+ex.Message);
        }
    }
   
    private string[] GetExcelHeaderList()
    {
        string[] strRets = { "", "", "", "", "", "", "", "", "" }; //,"","","",""

        try
        {
            Application.Lock();
            if (Application["jobid"] != null)
                strRets[0] = Application["jobid"].ToString();
            if (Application["aktnavn"] != null)
                strRets[1] = Application["aktnavn"].ToString();
            if (Application["aktnr"] != null)
                strRets[2] = Application["aktnr"].ToString();
            if (Application["akttimer"] != null)
                strRets[3] = Application["akttimer"].ToString();
            if (Application["akttpris"] != null)
                strRets[4] = Application["akttpris"].ToString();
            if (Application["aktsum"] != null)
                strRets[5] = Application["aktsum"].ToString();
            if (Application["konto"] != null)
                strRets[6] = Application["konto"].ToString();
            if (Application["linjetype"] != null)
                strRets[7] = Application["linjetype"].ToString();
            if (Application["aktstdato"] != null)
                strRets[8] = Application["aktstdato"].ToString();


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
        int[] intRets = { 0, 0, 0, 0, 0, 0, 0, 0, 0 }; //, 0, 0, 0, 0

        try
        {

            if (ddlJobId.SelectedIndex != -1)
                intRets[0] = ddlJobId.SelectedIndex;
            if (ddlAktnavn.SelectedIndex != -1)
                intRets[1] = ddlAktnavn.SelectedIndex;
            if (ddlAktnr.SelectedIndex != -1)
                intRets[2] = ddlAktnr.SelectedIndex;
            if (ddlAkttimer.SelectedIndex != -1)
                intRets[3] = ddlAkttimer.SelectedIndex;
            if (ddlAkttpris.SelectedIndex != -1)
                intRets[4] = ddlAkttpris.SelectedIndex;
            if (ddlAktsum.SelectedIndex != -1)
                intRets[5] = ddlAktsum.SelectedIndex;
            if (ddlKonto.SelectedIndex != -1)
                intRets[6] = ddlKonto.SelectedIndex;
            if (ddlLinjetype.SelectedIndex != -1)
                intRets[7] = ddlLinjetype.SelectedIndex;
            if (ddlAktstDato.SelectedIndex != -1)
                intRets[8] = ddlAktstDato.SelectedIndex;

        }
        catch (Exception ex)
        {
            throw new Exception("Get excel integer header list error: " + ex.Message);
        }

        return intRets;
    }


    protected void ddlJobId_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblJobId.Text = "fundet kolonne: " + ddlJobId.SelectedValue;
        Application["jobid"] = ddlJobId.SelectedValue;
    }

    protected void ddlAktnavn_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAktnavn.Text = "fundet kolonne: " + ddlAktnavn.SelectedValue;
        Application["aktnavn"] = ddlAktnavn.SelectedValue;
    }
   
    protected void ddlAktnr_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAktnr.Text = "fundet kolonne: " + ddlAktnr.SelectedValue;
        Application["aktnr"] = ddlAktnr.SelectedValue;
    }

    protected void ddlAkttimer_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAkttimer.Text = "fundet kolonne: " + ddlAkttimer.SelectedValue;
        Application["akttimer"] = ddlAkttimer.SelectedValue;
    }

    protected void ddlAkttpris_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAkttpris.Text = "fundet kolonne: " + ddlAkttpris.SelectedValue;
        Application["akttpris"] = ddlAkttpris.SelectedValue;
    }

    protected void ddlAktsum_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAktsum.Text = "fundet kolonne: " + ddlAktsum.SelectedValue;
        Application["aktsum"] = ddlAktsum.SelectedValue;
    }

    protected void ddlKonto_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblKonto.Text = "fundet kolonne: " + ddlKonto.SelectedValue;
        Application["konto"] = ddlKonto.SelectedValue;
    }

    protected void ddlLinjetype_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblLinjetype.Text = "fundet kolonne: " + ddlLinjetype.SelectedValue;
        Application["linjetype"] = ddlLinjetype.SelectedValue;
    }

    protected void ddlAktstDato_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAktstDato.Text = "fundet kolonne: " + ddlAktstDato.SelectedValue;
        Application["aktstdato"] = ddlAktstDato.SelectedValue;
    }




    protected void ddlJobId_DataBound(object sender, EventArgs e)
    {
        lblJobId.Text = "fundet kolonne: " + ddlJobId.SelectedValue;
        Application["jobid"] = ddlJobId.SelectedValue;
    }

    protected void ddlAktnavn_DataBound(object sender, EventArgs e)
    {
        lblAktnavn.Text = "fundet kolonne: " + ddlAktnavn.SelectedValue;
        Application["aktnavn"] = ddlAktnavn.SelectedValue;
    }
    
   
    protected void ddlAktnr_DataBound(object sender, EventArgs e)
    {
        lblAktnr.Text = "fundet kolonne: " + ddlAktnr.SelectedValue;
        Application["aktnr"] = ddlAktnr.SelectedValue;
    }


    protected void ddlAkttimer_DataBound(object sender, EventArgs e)
    {
        lblAkttimer.Text = "fundet kolonne: " + ddlAkttimer.SelectedValue;
        Application["akttimer"] = ddlAkttimer.SelectedValue;
    }

    protected void ddlAkttpris_DataBound(object sender, EventArgs e)
    {
        lblAkttpris.Text = "fundet kolonne: " + ddlAkttpris.SelectedValue;
        Application["akttpris"] = ddlAkttpris.SelectedValue;
    }


    protected void ddlAktsum_DataBound(object sender, EventArgs e)
    {
        lblAktsum.Text = "fundet kolonne: " + ddlAktsum.SelectedValue;
        Application["aktsum"] = ddlAktsum.SelectedValue;
    }

    protected void ddlAktstDato_DataBound(object sender, EventArgs e)
    {
        lblAktstDato.Text = "fundet kolonne: " + ddlAktstDato.SelectedValue;
        Application["aktstdato"] = ddlAktstDato.SelectedValue;
    }

    protected void ddlKonto_DataBound(object sender, EventArgs e)
    {
        lblKonto.Text = "fundet kolonne: " + ddlKonto.SelectedValue;
        Application["konto"] = ddlKonto.SelectedValue;
    }


    protected void ddlLinjetype_DataBound(object sender, EventArgs e)
    {
        lblLinjetype.Text = "fundet kolonne: " + ddlLinjetype.SelectedValue;
        Application["linjetype"] = ddlLinjetype.SelectedValue;
    }
   
    
}