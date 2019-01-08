using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Services;

namespace TestAsposeCellWithHotPatch
{
    /// <summary>
    /// PCWebService 的摘要说明
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // 若要允许使用 ASP.NET AJAX 从脚本中调用此 Web 服务，请取消注释以下行。 
    // [System.Web.Script.Services.ScriptService]
    public class PCWebService : System.Web.Services.WebService
    {

        [WebMethod]
        public string HelloWorld()
        {
            return "Hello World";
        }

        [WebMethod]
        public string ReadWriteExcel()
        {
            string path = this.Server.MapPath("~/" + "Test");
            if (System.IO.Directory.Exists(path) == false)
            {
                System.IO.Directory.CreateDirectory(path);
            }
            string pathTemplate = System.IO.Path.Combine(path, "TestAspose{0}.xlsx");
            return Util.Excel.ExcelUtils_Aspose.TestAsposeCellsHotPatch(pathTemplate);
        }

        [WebMethod]
        public string Print()
        {
            string filePath = System.IO.Path.Combine(this.Server.MapPath("~"), "GBL零部件运输交接确认单.xlsx");
            Util.Excel.ExcelUtils_Aspose.PrintDemo(filePath);
            return "Success";
        }
    }
}

