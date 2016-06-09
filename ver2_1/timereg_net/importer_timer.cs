using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data.Odbc;
using System.Data;


public partial class importer_timer : System.Web.UI.Page
{
    const string PATH2UPLOAD = "~/inc/excelUpload/";

    protected void Page_Load(object sender, EventArgs e)
    {
    }
    protected void btnUpload_Click(object sender, EventArgs e)
    {
        try
        {
            ozUploadFile service = new ozUploadFile();
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
            if (Request["editor"] != null && Request["lto"] != null && Request["mid"]!=null && txtFileName.Text != string.Empty)
            {
                lblStatus.Text = string.Empty;

                ozUploadFile service = new ozUploadFile();
                String path = Server.MapPath(PATH2UPLOAD);
                string[] headers = GetExcelHeaderList();
                int[] intHeaders = GetExcelIntHeaderList();

                int countIgnore = 0;
                int countSent = 0;

                string init = service.GetInit(Request["mid"], Request["lto"]);
                string serviceReturn = string.Empty;
                string errorLine = string.Empty;

                lblStatus.Text = init;

                if (init != string.Empty)

                    //lblStatus.Text = serviceReturn.ToString();

                   

                serviceReturn = service.Submit(path, txtFileName.Text, headers, intHeaders, Request["lto"], Request["editor"], ref countIgnore, ref countSent, init, ref errorLine, Request["mid"]);
                lblStatus.Text = "" + path + "," + txtFileName.Text + "," + headers + "," + intHeaders + "," + Request["lto"] + "," + Request["editor"];
                lblStatus.Text += "<br>OG VI SIGER: "+ serviceReturn.ToString();

                //lblStatus.Text = serviceReturn.ToString();

                int rowsInserted = GetRowInserted(serviceReturn);
                int rowsFailed = countSent - rowsInserted;
                int rowTotal = countSent + countIgnore;

                lblStatus.Text += "<br>RowsInserted: " + rowsInserted;

                string errorId = GetErrorId(serviceReturn);

                if (rowsFailed == 0 && rowsInserted > 0)
                {
                    if (countIgnore > 0)
                        lblStatus.Text = "Status message:<br><br> Filen indeholdt " + rowTotal + " linjer.<br>" + rowsInserted + " linjer indlæst korrekt.<br><br> " + countIgnore + " linjer er ignoreret. <br>" + errorLine;
                    else
                        lblStatus.Text = "Status message:<br><br> Filen indeholdt " + rowTotal + " linjer.<br>" + rowsInserted + " linjer indlæst korrekt.";
                }
                else
                {
                    if (countIgnore > 0)
                        lblStatus.Text = "Status message:<br><br> Filen indeholdt " + rowTotal + " linjer.<br>" + rowsInserted + " linjer indlæst korrekt.<br><br>" + countIgnore + " linjer er ignoreret. <br>" + errorLine + errorId;
                    else
                        lblStatus.Text = "Status message:<br><br> Filen indeholdt " + rowTotal + " linjer.<br>" + rowsInserted + " linjer indlæst korrekt.<br>" + errorId;
                }

                //lblStatus.Text += serviceReturn.ToString();

                //lblStatus.Text += "<br>Path: " + path;
                //lblStatus.Text += "<br>txtFileName.Text: " + txtFileName.Text;
                //lblStatus.Text += "<br>Headers: " + headers + " intHeaders: " + intHeaders;
                //lblStatus.Text += "<br>Variable LTO: " + Request["lto"] + " Editor: " + Request["editor"] + " Mid:" + Request["mid"];
                //lblStatus.Text += "<br>Init: " + init;
                //lblStatus.Text = "Hwej dr";

            }
            else
                lblStatus.Text = "Kan ikke finde editor, mid eller licens indehaver fra link i TimeOut.";
        }
          


        
        catch (Exception ex)
        {
            if (ex.Message.Contains("læsning"))
                lblStatus.Text = "Du er nødt til at upload en ny fil.";
            else

                
                lblStatus.Text += "Server error: "+ex.Message;
               
               

                //lblStatus.Text += "<br>Path: " + path;
                //lblStatus.Text += "<br>txtFileName.Text: " + txtFileName.Text;
                //lblStatus.Text += "<br>Headers: " + headers + " intHeaders: " + intHeaders;
                lblStatus.Text += "<br>Variable LTO: " + Request["lto"] + " Editor: " + Request["editor"] + " Mid:" + Request["mid"];
                //lblStatus.Text += "<br>Init: " + init;
      
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
            ozUploadFile service = new ozUploadFile();
            String path = Server.MapPath(PATH2UPLOAD);
            List<string> header = service.ReadHeader(path, txtFileName.Text, Request["lto"], Request["editor"]);

            if (header.Count < 6)
                lblUploadStatus.Text = "<b>Der skal være 6 kolonner i excel filen</b>";
            else
            {
                SelectDDL(ddlMedarbejderId, "initial", header, lblMedarbejderId, "medarbejderid");
                SelectDDL(ddlJobId, "job", header, lblJobId, "jobid");
                SelectDDL(ddlTimer, "timer", header, lblTimer, "timer");
                SelectDDL(ddlDato, "dato", header, lblDato, "dato");
                SelectDDL(ddlAktNavn, "aktivit", header, lblAktNavn, "aktnavn");
                SelectDDL(ddlTimerKom, "komment", header, lblTimerKom, "timerKom");
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Get header error: "+ex.Message);
        }
    }
   
    private string[] GetExcelHeaderList()
    {
        string[] strRets = { "","","","","",""};

        try
        {
            Application.Lock();
            if (Application["medarbejderid"] != null)
                strRets[0] = Application["medarbejderid"].ToString();
            if (Application["jobid"] != null)
                strRets[1] = Application["jobid"].ToString();
            if (Application["timer"] != null)
                strRets[2] = Application["timer"].ToString();
            if (Application["dato"] != null)
                strRets[3] = Application["dato"].ToString();
            if (Application["aktnavn"] != null)
                strRets[4] = Application["aktnavn"].ToString();
            if (Application["timerKom"] != null)
                strRets[5] = Application["timerKom"].ToString();
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
        int[] intRets = { 0, 0, 0, 0, 0, 0 };

        try
        {
            if (ddlMedarbejderId.SelectedIndex != -1)
                intRets[0] = ddlMedarbejderId.SelectedIndex;
            if (ddlJobId.SelectedIndex != -1)
                intRets[1] = ddlJobId.SelectedIndex;
            if (ddlTimer.SelectedIndex != -1)
                intRets[2] = ddlTimer.SelectedIndex;
            if (ddlDato.SelectedIndex != -1)
                intRets[3] = ddlDato.SelectedIndex;
            if (ddlAktNavn.SelectedIndex != -1)
                intRets[4] = ddlAktNavn.SelectedIndex;
            if (ddlTimerKom.SelectedIndex != -1)
                intRets[5] = ddlTimerKom.SelectedIndex;
        }
        catch (Exception ex)
        {
            throw new Exception("Get excel integer header list error: " + ex.Message);
        }

        return intRets;
    }

    protected void ddlMedarbejderId_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMedarbejderId.Text = "fundet kolonne: " + ddlMedarbejderId.SelectedValue;
        Application["medarbejderid"] = ddlMedarbejderId.SelectedValue;
    }
    protected void ddlJobId_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblJobId.Text = "fundet kolonne: " + ddlJobId.SelectedValue;
        Application["jobid"] = ddlJobId.SelectedValue;
    }
    protected void ddlTimer_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblTimer.Text = "fundet kolonne: " + ddlTimer.SelectedValue;
        Application["timer"] = ddlTimer.SelectedValue;
    }
    protected void ddlDato_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblDato.Text = "fundet kolonne: " + ddlDato.SelectedValue;
        Application["dato"] = ddlDato.SelectedValue;
    }
    protected void ddlAktNavn_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAktNavn.Text = "fundet kolonne: " + ddlAktNavn.SelectedValue;
        Application["aktnavn"] = ddlAktNavn.SelectedValue;
    }
    protected void ddlTimerKom_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblTimerKom.Text = "fundet kolonne: " + ddlTimerKom.SelectedValue;
        Application["timerKom"] = ddlTimerKom.SelectedValue;
    }
    protected void ddlMedarbejderId_DataBound(object sender, EventArgs e)
    {
        lblMedarbejderId.Text = "fundet kolonne: " + ddlMedarbejderId.SelectedValue;
        Application["medarbejderid"] = ddlMedarbejderId.SelectedValue;
    }
    protected void ddlJobId_DataBound(object sender, EventArgs e)
    {
        lblJobId.Text = "fundet kolonne: " + ddlJobId.SelectedValue;
        Application["jobid"] = ddlJobId.SelectedValue;
    }
    protected void ddlTimer_DataBound(object sender, EventArgs e)
    {
        lblTimer.Text = "fundet kolonne: " + ddlTimer.SelectedValue;
        Application["timer"] = ddlTimer.SelectedValue;
    }
    protected void ddlDato_DataBound(object sender, EventArgs e)
    {
        lblDato.Text = "fundet kolonne: " + ddlDato.SelectedValue;
        Application["dato"] = ddlDato.SelectedValue;
    }
    protected void ddlAktNavn_DataBound(object sender, EventArgs e)
    {
        lblAktNavn.Text = "fundet kolonne: " + ddlAktNavn.SelectedValue;
        Application["aktnavn"] = ddlAktNavn.SelectedValue;
    }
    protected void ddlTimerKom_DataBound(object sender, EventArgs e)
    {
        lblTimerKom.Text = "fundet kolonne: " + ddlTimerKom.SelectedValue;
        Application["timerKom"] = ddlTimerKom.SelectedValue;
    }
}