using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Text;
using Util.WebServiceModel;

namespace TestMSExcelWebApplication
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

        [WebMethod(Description = "读取公式, 计算公式值后保存")]
        public string TestOpenAndSaveExcel() // 已测试成功读取到公式
        {
            // Step 1 Aspose.Cell 直接读取 -- 公式Cell的值为null
            // Step 2 InteropExcel 打开(自动计算公式值), 保存计算好公式值的文件
            // Step 3 Aspose.Cell 再次读取 -- 公式Cell的值不为null

            string path = System.IO.Path.Combine(Server.MapPath("~/bin"), "TestAspose2.xlsx");
            string saveFilePath = path;

            System.Data.DataTable dt = Util.Excel.ExcelUtils_Aspose.Excel2DataTable
                (
                    filePath: saveFilePath,
                    exportColumnName: false
                );

            StringBuilder sb = new StringBuilder();
            for (int rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sb.Append(dt.Rows[rowIndex][i]).Append(" ");
                }
                sb.Append(";");
                sb.AppendLine();
            }

            using (var excelUtilObj = new Util.Excel.ExcelUtil_InteropExcel(filePath: path, isWebApp: true))
            {
                // excelUtilObj.SaveCopyAs();
                excelUtilObj.Save();
            }

            dt = Util.Excel.ExcelUtils_Aspose.Excel2DataTable
                (
                    filePath: saveFilePath,
                    exportColumnName: false
                );


            sb.AppendLine("********* 用Excel程序打开并保存后 *********");

            for (int rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sb.Append(dt.Rows[rowIndex][i]).Append(" ");
                }
                sb.Append(";");
                sb.AppendLine();
            }
            return sb.ToString();
        }

        [WebMethod(Description = "上传 + 保存计算公式值测试 + 返回")]
        public SOAPResult Upload_Open_Save_GetBack(string base64Str)
        {
            // 1. 接收上传文件
            // 2. 用 ExcelUtil_InteropExcel 打开并保存文件 ( 计算公式的值 )
            // 3. 返回文件

            SOAPResult r = new SOAPResult();
            try
            {
                // 1. 接收上传文件
                string uploadFilePath = string.Empty;

                string localDirectory = System.Web.Hosting.HostingEnvironment.MapPath("~/Upload/Excel");
                if (System.IO.Directory.Exists(localDirectory) == false)
                {
                    System.IO.Directory.CreateDirectory(localDirectory);
                }
                uploadFilePath = System.IO.Path.Combine(localDirectory, Guid.NewGuid().ToString());
                Util.IO.FileUtils.SaveBase64StrToFile(uploadFilePath, base64Str, true);

                // 2. 用 ExcelUtil_InteropExcel 打开并保存文件 ( 计算公式的值 )
                string savedFormulaFilePath = string.Empty;
                using (var excelUtilObj = new Util.Excel.ExcelUtil_InteropExcel(filePath: uploadFilePath, isWebApp: true))
                {
                    // Test 1
                    excelUtilObj.Save();
                    savedFormulaFilePath = uploadFilePath;

                    // Test 2
                    savedFormulaFilePath = excelUtilObj.SaveCopyAs();
                }

                // 3. 返回文件
                r.SuccessNonJsonConvert(Util.IO.FileUtils.GetFileToBase64Str(savedFormulaFilePath));

            }
            catch (Exception e)
            {
                r.Error(e);
            }

            return r;
        }

        [WebMethod(Description = "打印测试")]
        public void TestPrint()
        {
            string path = System.IO.Path.Combine
            (
                 System.Web.Hosting.HostingEnvironment.MapPath("~"),
                 "GBL零部件运输交接确认单.xlsx"
            );

            using (var excelApp = new Util.Excel.ExcelUtil_InteropExcel(path, true))
            {
                excelApp.Print(isLandscape: true);
            }


            ////------------------------打印页面相关设置--------------------------------
            //workSheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;//纸张大小
            //workSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;//页面横向
            //                                                                      //workSheet.PageSetup.Zoom = 75; //打印时页面设置,缩放比例百分之几
            //workSheet.PageSetup.Zoom = false; //打印时页面设置,必须设置为false,页高,页宽才有效
            //workSheet.PageSetup.FitToPagesWide = 1; //设置页面缩放的页宽为1页宽
            //workSheet.PageSetup.FitToPagesTall = false; //设置页面缩放的页高自动
            //workSheet.PageSetup.LeftHeader = "成都盛特（Esimtech）";//页面左上边的标志
            //workSheet.PageSetup.CenterFooter = "第 &P 页，共 &N 页";//页面下标
            //workSheet.PageSetup.PrintGridlines = true; //打印单元格网线
            //workSheet.PageSetup.TopMargin = 1.5 / 0.035; //上边距为2cm（转换为in）
            //workSheet.PageSetup.BottomMargin = 1.5 / 0.035; //下边距为1.5cm
            //workSheet.PageSetup.LeftMargin = 2 / 0.035; //左边距为2cm
            //workSheet.PageSetup.RightMargin = 2 / 0.035; //右边距为2cm
            //workSheet.PageSetup.CenterHorizontally = true; //文字水平居中
        }

        [WebMethod(Description = "并发测试")]
        public string TestConcurrentRead(string fileName)
        {
            string path = System.IO.Path.Combine(this.Server.MapPath("~/ConcurrentTest/"), $"{fileName}.xlsx");
            return new PrintBll().Print(path);
        }

        [WebMethod(Description = "关闭测试1 -- 成功")]
        public string TestV2_Step1(string fileName)
        {
            string r = string.Empty;

            string path = System.IO.Path.Combine(this.Server.MapPath("~/ConcurrentTest/"), $"{fileName}.xlsx");
            using (var excelApp = new Util.Excel.ExcelApp_V2())
            {
                excelApp.Open(path);
                r = excelApp.WorksheetName.ToString();
            }

            return r;
        }

        [WebMethod(Description = "关闭测试2 -- 成功")]
        public string TestV2_Step2(string fileName)
        {
            string r = string.Empty;

            string path = System.IO.Path.Combine(this.Server.MapPath("~/ConcurrentTest/"), $"{fileName}.xlsx");
            using (var excelApp = new Util.Excel.ExcelApp_V2())
            {
                excelApp.Open(path);
                excelApp.Save();
                r = excelApp.WorksheetName.ToString();
            }

            return r;
        }

        [WebMethod(Description = "关闭测试3 -- 成功")]
        public string TestV2_Step3(string fileName)
        {
            string r = string.Empty;

            string path = System.IO.Path.Combine(this.Server.MapPath("~/ConcurrentTest/"), $"{fileName}.xlsx");
            using (var excelApp = new Util.Excel.ExcelApp_V2())
            {
                excelApp.Open(path);
                excelApp.Save();
                excelApp.Print(isLandscape: true);
                r = excelApp.WorksheetName.ToString();
            }

            return r;
        }
    }

    public class JsonResult
    {
        public JsonResult()
        {
            //
            //TODO: 在此处添加构造函数逻辑
            //
        }

        public bool success
        {
            get;
            set;
        }

        public string message
        {
            get;
            set;
        }

        public Object data
        {
            get;
            set;
        }
    }

    public static class JsonResultExt
    {

        public static JsonResult Pass(this JsonResult r, Object data)
        {
            r.success = true;
            r.data = data;
            return r;
        }

        public static JsonResult Fail(this JsonResult r, Exception ex)
        {
            r.success = false;
            r.message = ex.Message;
            return r;
        }

        public static JsonResult Right(this JsonResult r, string msg)
        {
            r.success = true;
            r.message = msg;
            return r;
        }

        public static JsonResult Error(this JsonResult r, string error)
        {
            r.success = false;
            r.message = error;
            return r;
        }


    }
}
