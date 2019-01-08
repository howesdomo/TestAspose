using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;

namespace TestAsposeCellWithCrack
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
        public void ReadWriteExcel()
        {
            method();
        }

        private void method()
        {
            string path = @"D:\SC_Github\TestAspose\TestAsposeCellWithCrack\TestAspose2.xlsx";
            Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook(path);
            bool isLicensed = wb.IsLicensed;
            var ws = wb.Worksheets[0];
            var cell = ws.Cells[0, 1];
            string value = cell.StringValue;

            cell.Value = "3322";

            string filePath = System.IO.Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("~"), string.Format("TestAspose_{0}.xlsx", DateTime.Now.ToString("HHmmss")));
            wb.Save(filePath);
        }
    }
}
