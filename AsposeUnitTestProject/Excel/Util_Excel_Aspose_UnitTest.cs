using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AsposeUnitTestProject
{
    [TestClass]
    public class Util_Excel_Aspose_UnitTest
    {
        [TestMethod]
        public void Test_Excel2DataTableAsString()
        {
            // 测试总结 : 读取日期时间的值 如同我们直接在Excel看到的内容, 会与实际点进去的 Value 的值由偏差
            // 例如 Sheet 2 的 C4 的值 表面的值是 2019/9/9 9:31 但实际点进去可以看到是 2019/9/9 9:31:49

            string path = System.IO.Path.Combine(new System.IO.DirectoryInfo(Environment.CurrentDirectory).Parent.Parent.FullName, "Excel", "Aspose.xlsx");

            // **** 读取 Sheet1 ****
            var datatable = Util.Excel.ExcelUtils_Aspose.Excel2DataTableAsString(path);

            Assert.AreEqual<string>("Sheet1", datatable.TableName);

            Assert.AreEqual<int>(2, datatable.Rows.Count);
            Assert.AreEqual<int>(4, datatable.Columns.Count);

            Assert.AreEqual<string>("54.08", datatable.Rows[0]["体重"].ToString());

            Assert.AreEqual<string>("45.08", datatable.Rows[1]["体重"].ToString());


            // **** 读取 Sheet2 ****
            datatable = Util.Excel.ExcelUtils_Aspose.Excel2DataTableAsString(path, sheetIndex: 1);

            Assert.AreEqual<string>("工作簿2", datatable.TableName);

            Assert.AreEqual<int>(3, datatable.Rows.Count);
            Assert.AreEqual<int>(3, datatable.Columns.Count);

            Assert.AreEqual<string>("9:32", datatable.Rows[0]["时间"].ToString());

            Assert.AreEqual<string>("2019/5/2", datatable.Rows[1]["日期"].ToString());

            Assert.AreEqual<string>("2019/9/9 9:31", datatable.Rows[2]["日期时间"].ToString()); // * 重要 *


            // **** 读取 Sheet2 **** 不转换第一行为列头
            datatable = Util.Excel.ExcelUtils_Aspose.Excel2DataTableAsString
            (
                filePath: path,
                sheetIndex: 1,
                exportColumnName: false // * 重要 *
            );

            Assert.AreEqual<string>("工作簿2", datatable.TableName);

            Assert.AreEqual<int>(4, datatable.Rows.Count);
            Assert.AreEqual<int>(3, datatable.Columns.Count);

            Assert.AreEqual<string>("时间", datatable.Rows[0][0].ToString());
            Assert.AreEqual<string>("日期", datatable.Rows[0][1].ToString());
            Assert.AreEqual<string>("日期时间", datatable.Rows[0][2].ToString());

            Assert.AreEqual<string>("9:32", datatable.Rows[1][0].ToString());

            Assert.AreEqual<string>("2019/5/2", datatable.Rows[2][1].ToString());

            Assert.AreEqual<string>("2019/9/9 9:31", datatable.Rows[3][2].ToString()); // * 重要 *
        }

        [TestMethod]
        public void Test_Excel2DataTable()
        {
            // 测试总结 : 
            // 1) 采用转换列头时
            // 时间、日期、 时间日期 的转换都会转成 yyyy-M-D H:mm:ss 这个格式
            // 2) 不采用转换列头
            // 时间、日期、 时间日期 的转换与 AsString 的转换结果一致

            string path = System.IO.Path.Combine(new System.IO.DirectoryInfo(Environment.CurrentDirectory).Parent.Parent.FullName, "Excel", "Aspose.xlsx");

            // **** 读取 Sheet1 ****
            var datatable = Util.Excel.ExcelUtils_Aspose.Excel2DataTable(path);

            Assert.AreEqual<string>("Sheet1", datatable.TableName);

            Assert.AreEqual<int>(2, datatable.Rows.Count);
            Assert.AreEqual<int>(4, datatable.Columns.Count);

            Assert.AreEqual<string>("54.08", datatable.Rows[0]["体重"].ToString());

            Assert.AreEqual<string>("45.08", datatable.Rows[1]["体重"].ToString());


            // **** 读取 Sheet2 ****
            datatable = Util.Excel.ExcelUtils_Aspose.Excel2DataTable(path, sheetIndex: 1);

            Assert.AreEqual<string>("工作簿2", datatable.TableName);

            Assert.AreEqual<int>(3, datatable.Rows.Count);
            Assert.AreEqual<int>(3, datatable.Columns.Count);

            Assert.AreEqual<string>("1899/12/31 9:32:00", datatable.Rows[0]["时间"].ToString()); // * 重要 *

            Assert.AreEqual<string>("2019/5/2 0:00:00", datatable.Rows[1]["日期"].ToString()); // * 重要 *

            Assert.AreEqual<string>("2019/9/9 9:31:49", datatable.Rows[2]["日期时间"].ToString()); // * 重要 *

            // **** 读取 Sheet2 **** 不转换第一行为列头
            datatable = Util.Excel.ExcelUtils_Aspose.Excel2DataTable
            (
                filePath: path,
                sheetIndex: 1,
                exportColumnName: false // * 重要 *
            );

            Assert.AreEqual<string>("工作簿2", datatable.TableName);

            Assert.AreEqual<int>(4, datatable.Rows.Count);
            Assert.AreEqual<int>(3, datatable.Columns.Count);

            Assert.AreEqual<string>("时间", datatable.Rows[0][0].ToString());
            Assert.AreEqual<string>("日期", datatable.Rows[0][1].ToString());
            Assert.AreEqual<string>("日期时间", datatable.Rows[0][2].ToString());

            Assert.AreEqual<string>("9:32", datatable.Rows[1][0].ToString()); // * 重要 *

            Assert.AreEqual<string>("2019/5/2", datatable.Rows[2][1].ToString()); // * 重要 *

            Assert.AreEqual<string>("2019/9/9 9:31", datatable.Rows[3][2].ToString()); // * 重要 *
        }
    }
}
