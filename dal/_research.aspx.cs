using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class research2 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Label1.Text = getContent();
    }
    private String getContent()
    {
        String printData = "";
        //判斷瀏覽器
        Boolean isIE = false;

        System.Web.HttpBrowserCapabilities browser = Request.Browser;
        //如果是IE的話版本小於10就要加上line.gif
        if (browser.Type.ToLower().IndexOf("ie") > -1)
        {
            if (Int32.Parse(browser.Version, System.Globalization.NumberStyles.AllowDecimalPoint) < 10)
            {
                isIE = true;
            }
        }
        
        String connectConfig = "Server=localhost;database=2007Dal;uid=user;pwd=DB64x2ms7;";
        //預設全部都可以查到
        String query = "SELECT * FROM [Student]";
        //如果class, year參數為空就用預設查詢
        if (Request.QueryString["class"] != null && Request.QueryString["year"] != null)
        {
            query += " WHERE [year] = " + Request.QueryString["year"] + " AND ";
            //當class參數為0的時候 查大學生 為1查研究生 
            switch (Request.QueryString["class"])
            {
                case "0":
                    query += " [class] = '大學生'";
                    break;
                case "1":
                    query += " [class] = '研究生'";
                    break;
            }
        }

        query += " ORDER BY [MemId] DESC";


        //開啟sql連結
        using (SqlConnection sqlConnect = new SqlConnection(connectConfig))
        {
            sqlConnect.Open();

            SqlCommand cmd = new SqlCommand(query, sqlConnect);
            SqlDataReader data = cmd.ExecuteReader();

            while (data.Read())
            {
                printData += "<div class=\"jquery-shadow-raised research\">";

                printData += "<div id=\"type\">" + data["type"] + "</div>";
                printData += "<div id=\"title\">" + data["Topic"] + "</div>";
                printData += "<div id=\"members\">" + data["name"] + "</div>";
                printData += "<div id=\"summary\">摘&nbsp;&nbsp;要</div>";
                printData += "<div id=\"detail\">" + data["detail"] + "</div>";
                if (!data["key"].Equals("NONE"))
                {
                    printData += "<div id=\"key_title\">關鍵字</div>";
                    printData += "<div id=\"key\">" + data["key"] + "</div>";
                }
                //是ie10以下的話就加入線
                if (isIE)
                {
                    printData += "<img style=\"width:100%;height:8px;\" src=\"img/line.gif\" />";
                }

                printData += "</div>";
            }
            /*
             <div class="research">
                <div id="type"></div>
                <div id="title"></div>
                <div id="members"></div>
                <div id="summary">摘&nbsp;&nbsp;要</div>
                <div id="detail"></div>
                <div id="key_title">關鍵字</div>
                <div id="key"></div>
            </div>
             */
        }
        return printData;
    }
}