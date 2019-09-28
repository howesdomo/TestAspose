using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TestMSExcelWebApplication
{
    public class PrintBll
    {
        private static object _LOCK_ = new object();

        public string Print(string path)
        {
            lock (_LOCK_)
            {
                string r = string.Empty;

                using (var excelApp = new Util.Excel.ExcelUtil_InteropExcel(path, true))
                {
                    r = excelApp.ActiveWorksheetName;
                    excelApp.Print(isLandscape: true);
                }

                return r;
            }
        }
    }
}