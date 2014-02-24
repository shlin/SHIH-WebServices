using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Configuration;

public partial class research : System.Web.UI.Page
{
    String connectConfig;
    protected void Page_Load(object sender, EventArgs e)
    {
        connectConfig = System.Configuration.ConfigurationManager.ConnectionStrings["2007DALConnectionString"].ToString();

        Label2.Text = getPostgraduate();
        Label3.Text = getCollege();
    }
    private String getPostgraduate()
    {
        String result = "";
        String queryString = "SELECT DISTINCT [year] FROM [Student] WHERE [class] = '研究生' ORDER BY [year]";
        //使用using可以避免忘記關閉資源的情況
        using (SqlConnection sqlConnection = new SqlConnection(connectConfig))
        {
            sqlConnection.Open();

            SqlCommand cmd = new SqlCommand(queryString, sqlConnection);
            SqlDataReader data = cmd.ExecuteReader();

            while (data.Read())
            {
                result += "<li>";
                result += "<a href=\"achievement.aspx?class=1&year=" + data["year"] + "\">";
                result += data["year"];
                result += "</a>";
                result += "</li>";
            }
        }

        return result;
    }
    private String getCollege()
    {
        String result = "";
        String queryString = "SELECT DISTINCT [year] FROM [Student] WHERE [class] = '大學生' ORDER BY [year]";
        //使用using可以避免忘記關閉資源的情況
        using (SqlConnection sqlConnection = new SqlConnection(connectConfig))
        {
            sqlConnection.Open();

            SqlCommand cmd = new SqlCommand(queryString, sqlConnection);
            SqlDataReader data = cmd.ExecuteReader();

            while (data.Read())
            {
                result += "<li>";
                result += "<a href=\"achievement.aspx?class=0&year=" + data["year"] + "\">";
                result += data["year"];
                result += "</a>";
                result += "</li>";
            }
        }

        return result;
    }
}