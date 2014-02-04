using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Xml;

public partial class MasterPage : System.Web.UI.MasterPage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        // 處理 #menuList 的選單讀取，來源是 xmlDB/globalMenu.xml
        XmlDocument xmlDoc = new XmlDocument();
        XmlNode rootNode;
        XmlNodeList currentNodeList;

        xmlDoc.Load(Server.MapPath("xmlDB/globalMenu.xml"));
        rootNode = xmlDoc.SelectSingleNode("globalMenu");

        currentNodeList = rootNode.SelectNodes("Item");

        foreach (XmlNode current in currentNodeList)
        {
            String newString = "<a href=\"" + current.SelectSingleNode("Link").InnerText + "\">"+ current.SelectSingleNode("Title").InnerText+"</a>";
            HtmlGenericControl newItem = new HtmlGenericControl("li");
            newItem.InnerHtml = newString;
            menuList.Controls.Add(newItem);
        }
    }
}
