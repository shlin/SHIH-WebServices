using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;

public partial class aboutme : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        XmlDocument xmlDocEdu = new XmlDocument();

        xmlDocEdu.Load(Server.MapPath("xmlDB/Education.xml"));

    }
}