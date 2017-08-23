using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data.Odbc;
using System.Data;

public partial class importer_med : System.Web.UI.Page
{
    const string PATH2UPLOAD = "~/inc/excelUpload/";
    public string importtype = "";

    protected void Page_Load(object sender, EventArgs e)
    {
        importtype = Request["importtype"];

      
        feltnr1navn.Text = "No. (Init):";
        feltnr2navn.Text = "Name:";
        feltnr3navn.Text = "Email:";
        feltnr4navn.Text = "Norm time:";
        feltnr5navn.Text = "Employment date:";
        feltnr6navn.Text = "Termination date:";
        feltnr7navn.Text = "Blocked";
        feltnr8navn.Text = "Expence vendor no.:";
        feltnr9navn.Text = "Costcenter:";
        feltnr10navn.Text = "Linemanger:";
        feltnr11navn.Text = "Country Code:";
        feltnr12navn.Text = "Weblang:";
        





    }


    protected void btnUpload_Click(object sender, EventArgs e)
    {
        try
        {
            ozUploadFileMed service = new ozUploadFileMed();
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

                ozUploadFileMed service = new ozUploadFileMed();
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



    private string GetMinit(string msgIn)
    {
        string errorId = string.Empty;

        int indexMinit = msgIn.IndexOf("jobnr:");
        int indexRightEnd = msgIn.IndexOf(")");

        if (indexMinit != -1)
            errorId = msgIn.Substring(indexMinit + 6, indexRightEnd - indexMinit);

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
            ozUploadFileMed service = new ozUploadFileMed();
            String path = Server.MapPath(PATH2UPLOAD);
            List<string> header = service.ReadHeader(path, txtFileName.Text, Request["lto"], Request["editor"], Request["importtype"]);

            if (header.Count < 12)
                lblUploadStatus.Text = "<b>Der skal være 12 kolonner i excel filen</b>";
            else
            {

                string importtype = Request["importtype"];


        
                    SelectDDL(ddlMinit, "initial", header, lblMinit, "minit");
                    SelectDDL(ddlNavn, "navn", header, lblNavn, "navn");
                    SelectDDL(ddlEmail, "e-mail", header, lblEmail, "email");
                    SelectDDL(ddlNorm, "normtid", header, lblNorm, "norm");
                    SelectDDL(ddlAnsatdato, "employment", header, lblAnsatdato, "ansatdato");
                    SelectDDL(ddlOpsagtdato, "termination", header, lblOpsagtdato, "opdagtdato");
                    SelectDDL(ddlMansat, "blocked", header, lblMansat, "mansat");
                    SelectDDL(ddlEvn, "expvendorno", header, lblEvn, "evn");

                    SelectDDL(ddlCostcenter, "cost center", header, lblCostcenter, "costcenter");
                    SelectDDL(ddlLinemanager, "linemanager", header, lblLinemanager, "linemanager");
                    SelectDDL(ddlCcode, "country code", header, lblCcode, "countrycode");
                    SelectDDL(ddlWeblang, "weblang", header, lblWeblang, "weblang");



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
            if (Application["minit"] != null)
                strRets[0] = Application["minit"].ToString();
            if (Application["navn"] != null)
                strRets[1] = Application["navn"].ToString();
            if (Application["email"] != null)
                strRets[2] = Application["email"].ToString();
            if (Application["norm"] != null)
                strRets[3] = Application["norm"].ToString();
            if (Application["ansatdato"] != null)
                strRets[4] = Application["ansatdato"].ToString();
            if (Application["opsagtdato"] != null)
                strRets[5] = Application["opsagtdato"].ToString();
            if (Application["mansat"] != null)
                strRets[6] = Application["mansat"].ToString();
            if (Application["evn"] != null)
                strRets[7] = Application["evn"].ToString();
            if (Application["costcenter"] != null)
                strRets[8] = Application["costcenter"].ToString();
            if (Application["linemanager"] != null)
                strRets[9] = Application["linemanager"].ToString();
            if (Application["countrycode"] != null)
                strRets[10] = Application["countrycode"].ToString();
            if (Application["weblang"] != null)
                strRets[11] = Application["weblang"].ToString();
            


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

            if (ddlMinit.SelectedIndex != -1)
                intRets[0] = ddlMinit.SelectedIndex;
            if (ddlNavn.SelectedIndex != -1)
                intRets[1] = ddlNavn.SelectedIndex;
            if (ddlEmail.SelectedIndex != -1)
                intRets[2] = ddlEmail.SelectedIndex;
            if (ddlNorm.SelectedIndex != -1)
                intRets[3] = ddlNorm.SelectedIndex;
            if (ddlAnsatdato.SelectedIndex != -1)
                intRets[4] = ddlAnsatdato.SelectedIndex;
            if (ddlOpsagtdato.SelectedIndex != -1)
                intRets[5] = ddlOpsagtdato.SelectedIndex;
            if (ddlMansat.SelectedIndex != -1)
                intRets[6] = ddlMansat.SelectedIndex;
            if (ddlEvn.SelectedIndex != -1)
                intRets[7] = ddlEvn.SelectedIndex;
            if (ddlCostcenter.SelectedIndex != -1)
                intRets[8] = ddlCostcenter.SelectedIndex;
            if (ddlLinemanager.SelectedIndex != -1)
                intRets[9] = ddlLinemanager.SelectedIndex;
            if (ddlCcode.SelectedIndex != -1)
                intRets[10] = ddlCcode.SelectedIndex;
            if (ddlWeblang.SelectedIndex != -1)
                intRets[11] = ddlWeblang.SelectedIndex;

        }
        catch (Exception ex)
        {
            throw new Exception("Get excel integer header list error: " + ex.Message);
        }

        return intRets;
    }


    protected void ddlMinit_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMinit.Text = "fundet kolonne: " + ddlMinit.SelectedValue;
        Application["minit"] = ddlMinit.SelectedValue;
    }

    protected void ddlNavn_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblNavn.Text = "fundet kolonne: " + ddlNavn.SelectedValue;
        Application["navn"] = ddlNavn.SelectedValue;
    }
   
    protected void ddlEmail_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblEmail.Text = "fundet kolonne: " + ddlEmail.SelectedValue;
        Application["email"] = ddlEmail.SelectedValue;
    }

    protected void ddlNorm_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblNorm.Text = "fundet kolonne: " + ddlNorm.SelectedValue;
        Application["norm"] = ddlNorm.SelectedValue;
    }

    protected void ddlAnsatdato_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblAnsatdato.Text = "fundet kolonne: " + ddlAnsatdato.SelectedValue;
        Application["ansatdato"] = ddlAnsatdato.SelectedValue;
    }

    protected void ddlOpsagtdato_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblOpsagtdato.Text = "fundet kolonne: " + ddlOpsagtdato.SelectedValue;
        Application["opsagtdato"] = ddlOpsagtdato.SelectedValue;
    }

    protected void ddlMansat_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMansat.Text = "fundet kolonne: " + ddlMansat.SelectedValue;
        Application["mansat"] = ddlMansat.SelectedValue;
    }

    protected void ddlEvn_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblEvn.Text = "fundet kolonne: " + ddlEvn.SelectedValue;
        Application["evn"] = ddlEvn.SelectedValue;
    }


    protected void ddlCostcenter_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblCostcenter.Text = "fundet kolonne: " + ddlCostcenter.SelectedValue;
        Application["costcenter"] = ddlCostcenter.SelectedValue;
    }

    protected void ddlLinemanager_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblLinemanager.Text = "fundet kolonne: " + ddlLinemanager.SelectedValue;
        Application["linemanager"] = ddlLinemanager.SelectedValue;
    }

    protected void ddlCcode_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblCcode.Text = "fundet kolonne: " + ddlCcode.SelectedValue;
        Application["ccode"] = ddlCcode.SelectedValue;
    }

    protected void ddlWeblang_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblWeblang.Text = "fundet kolonne: " + ddlWeblang.SelectedValue;
        Application["weblang"] = ddlWeblang.SelectedValue;
    }


















    protected void ddlMinit_DataBound(object sender, EventArgs e)
    {
        lblMinit.Text = "fundet kolonne: " + ddlMinit.SelectedValue;
        Application["minit"] = ddlMinit.SelectedValue;
    }

    protected void ddlNavn_DataBound(object sender, EventArgs e)
    {
        lblNavn.Text = "fundet kolonne: " + ddlNavn.SelectedValue;
        Application["navn"] = ddlNavn.SelectedValue;
    }
    
   
    protected void ddlEmail_DataBound(object sender, EventArgs e)
    {
        lblEmail.Text = "fundet kolonne: " + ddlEmail.SelectedValue;
        Application["email"] = ddlEmail.SelectedValue;
    }


    protected void ddlNorm_DataBound(object sender, EventArgs e)
    {
        lblNorm.Text = "fundet kolonne: " + ddlNorm.SelectedValue;
        Application["norm"] = ddlNorm.SelectedValue;
    }

    protected void ddlAnsatdato_DataBound(object sender, EventArgs e)
    {
        lblAnsatdato.Text = "fundet kolonne: " + ddlAnsatdato.SelectedValue;
        Application["ansatdato"] = ddlAnsatdato.SelectedValue;
    }


    protected void ddlOpsagtdato_DataBound(object sender, EventArgs e)
    {
        lblOpsagtdato.Text = "fundet kolonne: " + ddlOpsagtdato.SelectedValue;
        Application["opsagtdato"] = ddlOpsagtdato.SelectedValue;
    }

    protected void ddlMansat_DataBound(object sender, EventArgs e)
    {
        lblMansat.Text = "fundet kolonne: " + ddlMansat.SelectedValue;
        Application["mansat"] = ddlMansat.SelectedValue;
    }


    protected void ddlEvn_DataBound(object sender, EventArgs e)
    {
        lblEvn.Text = "fundet kolonne: " + ddlEvn.SelectedValue;
        Application["evn"] = ddlEvn.SelectedValue;
    }



    protected void ddlCostcenter_DataBound(object sender, EventArgs e)
    {
        lblCostcenter.Text = "fundet kolonne: " + ddlCostcenter.SelectedValue;
        Application["costcenter"] = ddlCostcenter.SelectedValue;
    }


    protected void ddlLinemanager_DataBound(object sender, EventArgs e)
    {
        lblLinemanager.Text = "fundet kolonne: " + ddlLinemanager.SelectedValue;
        Application["linemanager"] = ddlLinemanager.SelectedValue;
    }

    protected void ddlCcode_DataBound(object sender, EventArgs e)
    {
        lblCcode.Text = "fundet kolonne: " + ddlCcode.SelectedValue;
        Application["countrycode"] = ddlCcode.SelectedValue;
    }

    protected void ddlWeblang_DataBound(object sender, EventArgs e)
    {
        lblWeblang.Text = "fundet kolonne: " + ddlWeblang.SelectedValue;
        Application["weblang"] = ddlWeblang.SelectedValue;
    }


}