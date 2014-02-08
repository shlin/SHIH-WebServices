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
        XmlDocument xmlDoc = new XmlDocument();
        XmlNode nodeAbout;

        xmlDoc.Load(Server.MapPath("xmlDB/shihDB.xml"));

        nodeAbout = xmlDoc.SelectSingleNode("/shih/About");

        loadEducation(nodeAbout.SelectSingleNode("Education"));
        loadExperience(nodeAbout.SelectSingleNode("Experience"));
        loadService(nodeAbout.SelectSingleNode("Service"));
        loadHonor(nodeAbout.SelectSingleNode("Honor"));
    }

    protected void loadEducation(XmlNode rootNode)
    {
        XmlNodeList entries;

        entries = rootNode.SelectNodes("Entry");

        foreach (XmlNode current in entries)
        {
            HtmlGenericControl newRow, newCell;
            newRow = new HtmlGenericControl("tr");
            newRow.Attributes.Add("align", "center");

            // 學校
            newCell = new HtmlGenericControl("td");
            newCell.InnerHtml = "<a href=\"" + current.SelectSingleNode("School").Attributes["Link"].Value + "\" target=\"_blank\">" + current.SelectSingleNode("School").InnerText + "</a>";
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

    protected void loadExperience(XmlNode rootNode)
    {
        XmlNodeList entries;

        entries = rootNode.SelectNodes("Entry");

        foreach (XmlNode current in entries)
        {
            HtmlGenericControl newRow, newCell;
            newRow = new HtmlGenericControl("tr");
            newRow.Attributes.Add("align", "center");

            // 學校
            newCell = new HtmlGenericControl("td");
            newCell.InnerHtml = "<a href=\"" + current.SelectSingleNode("Organization").Attributes["Link"].Value + "\" target=\"_blank\">" + current.SelectSingleNode("Organization").InnerText + "</a>";
            newRow.Controls.Add(newCell);

            // 國別
            newCell = new HtmlGenericControl("td");
            newCell.InnerHtml = current.SelectSingleNode("Department").InnerText;
            newRow.Controls.Add(newCell);

            // 主修學系
            newCell = new HtmlGenericControl("td");
            newCell.InnerHtml = current.SelectSingleNode("Title").InnerText;
            newRow.Controls.Add(newCell);

            // 起訖時間
            newCell = new HtmlGenericControl("td");
            newCell.InnerHtml = current.SelectSingleNode("During").InnerText;
            newRow.Controls.Add(newCell);


            if (current.SelectSingleNode("Organization").Attributes["ING"].Value == "1")
                expTable.Controls.AddAt(expTable.Controls.Count - 2, newRow);
            else
                expTable.Controls.Add(newRow);
        }
    }

    protected void loadService(XmlNode rootNode)
    {
        XmlNodeList entries = rootNode.SelectNodes("Entry");

        foreach (XmlNode current in entries)
        {
            HtmlGenericControl newList = new HtmlGenericControl("li");

            newList.InnerHtml = current.InnerText;

            serviceList.Controls.Add(newList);
        }
    }

    protected void loadHonor(XmlNode rootNode)
    {
        XmlNodeList entries = rootNode.SelectNodes("Entry");

        foreach (XmlNode current in entries)
        {
            HtmlGenericControl newList = new HtmlGenericControl("li");

            newList.InnerHtml = current.InnerText;

            honorList.Controls.Add(newList);
        }
    }
}