using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Web.Services;

public partial class national_holidays : System.Web.UI.Page
{
    #region Variable Declaration

    static string connString = Convert.ToString(System.Configuration.ConfigurationManager.ConnectionStrings["ConnectionString"]);
    string holidayQuery = "SELECT nh_id, nh_date, nh_duration, nh_editor FROM national_holidays GROUP BY nh_editor ORDER BY nh_date";
    string holidayDatesQuery = "SELECT nh_id, nh_date, nh_duration, nh_editor FROM national_holidays ORDER BY nh_date";
    string holidayFilterQuery = "nh_editor = '{0}' AND Convert(nh_date, 'System.String') LIKE '*/{1} *'";
    static string holidayDeleteQuery = "DELETE FROM national_holidays";
    //static string holidayInsertQuery = "INSERT INTO national_holidays (nh_date, nh_duration, nh_editor) VALUES {0}";
    static string holidayInsertQuery = "INSERT INTO national_holidays (nh_date, nh_duration, nh_editor) VALUES {'2017-01-01', 1, 'SK'}";

    #endregion

    #region Events

    /// <summary>
    /// Handles the Load event of the Page control.
    /// </summary>
    /// <param name="sender">The source of the event.</param>
    /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    protected void Page_Load(object sender, EventArgs e)
    {
        if(!Page.IsPostBack)
        {
            this.BindNationalHolidays();
        }
    }

    #endregion

    #region Methods

    /// <summary>
    /// Binds the national holidays.
    /// </summary>
    private void BindNationalHolidays()
    {
        MySqlConnection sqlConnection = new MySqlConnection();
        MySqlCommand sqlCommand;
        string holidatHTML = string.Empty;
        
        try
        {
            holidatHTML = "<table width='100%' cellpadding='0' cellspacing='0' id='tbl_Holidays'>";
            holidatHTML += "<tr>";
            holidatHTML += "<th class='duration'>Country</th><th class='duration'>Duration</th><th class='name'>Name</th>";
            int year = DateTime.Now.Year;
            for (int i = 1; i <= 10; i++)
            {
                holidatHTML += "<th class='year'>" + year + "</th>";
                year++;
            }

            holidatHTML += "</tr>";

            sqlConnection = new MySqlConnection(connString);
            sqlConnection.Open();

            sqlCommand = sqlConnection.CreateCommand();
            sqlCommand.CommandText = holidayQuery;

            MySqlDataAdapter sqlDataAdapter = new MySqlDataAdapter(sqlCommand);
            DataSet dsHoliday = new DataSet();
            sqlDataAdapter.Fill(dsHoliday);

            sqlCommand.CommandText = holidayDatesQuery;

            sqlDataAdapter = new MySqlDataAdapter(sqlCommand);
            DataSet dsHolidayDates = new DataSet();
            sqlDataAdapter.Fill(dsHolidayDates);

            if (dsHoliday.Tables.Count > 0 && dsHoliday.Tables[0].Rows.Count > 0)
            {
                DataRow drHoliday;
                string holidayName = string.Empty;
                DataRow[] drFiltered;
                int count = 0;

                for (int i = 0; i < dsHoliday.Tables[0].Rows.Count; i++)
                {
                    count = i + 1;
                    drHoliday = dsHoliday.Tables[0].Rows[i];

                    holidatHTML += "<tr id='tr_0" + count + "'>";
                    holidatHTML += "<td class='duration'><select><option>DK</option></select></td>";
                    holidatHTML += "<td class='duration'><select id='ddl_Duration_0" + count + "'>";

                    if(Convert.ToInt32(drHoliday["nh_duration"]) == 1)
                    {
                        holidatHTML += "<option value='1' selected='selected'>Open for timerec</option>";
                        holidatHTML += "<option value='0'>Close for timerec</option>";
                    }
                    else
                    {
                        holidatHTML += "<option value='1'>Open for timerec</option>";
                        holidatHTML += "<option value='0' selected='selected'>Close for timerec</option>";
                    }

                    holidatHTML += "</select></td>";

                    holidayName = Convert.ToString(drHoliday["nh_editor"]);
                    holidatHTML += "<td class='name'><input type='text' value='" + holidayName + "' id='txt_Name_0" + count + "' /></td>";
                    
                    year = DateTime.Now.Year;
                    for (int j = 1; j <= 10; j++)
                    {
                        drFiltered = dsHolidayDates.Tables[0].Select(string.Format(holidayFilterQuery, holidayName, year));
                        if(drFiltered != null && drFiltered.Length > 0)
                        {
                            string calendarDate = Convert.ToDateTime(drFiltered[0]["nh_date"]).ToString("MM/dd");
                            holidatHTML += "<td class='year'><input type='text' value='" + calendarDate + "' id='txt_Year_0" + count + "_" + j + "' /></td>";
                        }
                        else
                        {
                            holidatHTML += "<td class='year'><input type='text' id='txt_Year_0" + count + "_" + j + "' /></td>";
                        }

                        year++;
                    }                    

                    holidatHTML += "</tr>";
                }
            }
        }
        catch (Exception ex)
        {
            //Response.Write(ex.InnerException.Message);
        }
        finally
        {
            if (sqlConnection != null && sqlConnection.State == System.Data.ConnectionState.Open)
            {
                sqlConnection.Close();
            }

            holidatHTML += "</table>";
            ltrHolidays.Text = holidatHTML;
        }
    }

    #endregion

    #region Web Method

    /// <summary>
    /// Saves the holidays.
    /// </summary>
    /// <param name="holidayData">The holiday data.</param>
    /// <returns>returns holiday id</returns>
    [WebMethod]
    public static string SaveHolidays(string holidayData)
    {
        MySqlConnection sqlConnection = new MySqlConnection();
        MySqlCommand sqlCommand;
        string strHolidayDataQuery = string.Empty;

        try
        {            
            string[] arrHolidays = holidayData.Split(new string[] { "@@@" }, StringSplitOptions.None);
            if(arrHolidays.Length > 0)
            {                
                for (int i = 0; i < arrHolidays.Length; i++)
                {                    
                    string[] arrHolidayFields = arrHolidays[i].Split(new string[] { "###" }, StringSplitOptions.None);                    
                    if (arrHolidayFields.Length > 0 && arrHolidayFields.Length == 3)
                    {
                        strHolidayDataQuery += string.Format("('{0}', '{1}', '{2}'),", arrHolidayFields[0], arrHolidayFields[1], arrHolidayFields[2]);
                    }
                }
            }

            if(!string.IsNullOrEmpty(strHolidayDataQuery))
            {
                strHolidayDataQuery = strHolidayDataQuery.TrimEnd(',');
                strHolidayDataQuery = string.Format(holidayInsertQuery, strHolidayDataQuery);

                sqlConnection = new MySqlConnection(connString);
                sqlConnection.Open();

                sqlCommand = sqlConnection.CreateCommand();
                sqlCommand.CommandType = CommandType.Text;
                sqlCommand.CommandText = holidayDeleteQuery;
                sqlCommand.ExecuteNonQuery();

                sqlCommand.CommandText = strHolidayDataQuery;
                return Convert.ToString(sqlCommand.ExecuteNonQuery());
            }

            return "0";            
        }
        catch (Exception ex)
        {
            return "0";
        }
        finally
        {
            if (sqlConnection != null && sqlConnection.State == System.Data.ConnectionState.Open)
            {
                sqlConnection.Close();
            }
        }
    }

    #endregion
}