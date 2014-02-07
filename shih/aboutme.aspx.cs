using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using System.Web.UI.HtmlControls;

public partial class aboutme : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        loadEducation();
    }

    protected void loadEducation()
    {
        XmlDocument xmlDocEdu = new XmlDocument();
        XmlNode rootNode;
        XmlNodeList entries;

        xmlDocEdu.Load(Server.MapPath("xmlDB/Education.xml"));

        rootNode = xmlDocEdu.SelectSingleNode("Education");
        entries = rootNode.SelectNodes("Entry");

        foreach (XmlNode current in entries)
        {
            HtmlGenericControl newRow, newCell;
            newRow = new HtmlGenericControl("tr");
            newRow.Attributes.Add("align", "center");

            // 學校
            newCell = new HtmlGenericControl("td");
            newCell.InnerHtml = "<a href=\"" + current.SelectSingleNode("School").Attributes["Link"].Value + "\">" + current.SelectSingleNode("School").InnerText + "</a>";
            newRow.Controls.Add(newCell);

            // 國別
            newCell = new HtmlGenericControl("td");
            newCell.InnerHtml = current.SelectSingleNode("Country").InnerText;
            newRow.Controls.Add(newCell);

            // 主修學系
            newCell = new HtmlGenericControl("td");
            newCell.InnerHtml = current.SelectSingleNode("Major").InnerText;
            newRow.Controls.Add(newCell);

            // 學位
            newCell = new HtmlGenericControl("td");
            newCell.InnerHtml = current.SelectSingleNode("Degree").InnerText;
            newRow.Controls.Add(newCell);

            // 起訖時間
            newCell = new HtmlGenericControl("td");
            newCell.InnerHtml = current.SelectSingleNode("During").InnerText;
            newRow.Controls.Add(newCell);

            eduTable.Controls.Add(newRow);
        }
    }
}