using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Windows;
using Util.Excel;

namespace XiaJBXie
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            initEvent();
        }

        private void initEvent()
        {
            btnImport.Click += BtnImport_Click;
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            string path = @"D:\HoweDesktop\Apose测试.xlsx";
            // string path = @"D:\HoweDesktop\TestMaxDataColumn.xlsx";

            #region Apose 测试

            var columns = new List<Util.Excel.PropertyColumn>()
            .Add("PropBool", "布尔值")
            .Add("PropDateTime", "时间")
            .Add("PropDateTimeInfo", "字符串类型_时间")
            .Add("PropForuma", "公式测试")
            .Add("PropForumaError", "错误公式测试")
            .Add("PropTestNullValue", "空值测试")
            .Add("PropNumeric", "数值测试")
            .Add("PropString", "字符串类型")
            // .Add("Badman", "坏人") // 测试缺少Column
            ;

            var columns_Sheet2 = new List<Util.Excel.PropertyColumn>()
            .Add("CPN", "CPN")
            .Add("KPN", "KPN")
            .Add("Desciption", "备注")
            .Add("RiQi", "日期")
            .Add("ShiJian", "时间")
            .Add("RiQiShiJian", "日期时间")
            .Add("RiQiShiJianDaiHaoMiao", "日期时间带毫秒")
            // .Add("Badman", "坏人") // 测试缺少Column
            ;

            var columns_Sheet2_V2 = new List<Util.Excel.PropertyColumn>()
            .Add("CPN", 2)// .Add("CPN", "CPN")
            .Add("KPN", 3)
            .Add("Desciption", 4)
            .Add("RiQi", 5)
            // .Add("ShiJian", "时间")
            .Add("RiQiShiJian", 7)
            .Add("RiQiShiJianDaiHaoMiao", 8)
            // .Add("Badman", "坏人") // 测试缺少Column
            ;

            var columns_Sheet3 = new List<Util.Excel.PropertyColumn>()
            .Add("QuYu", "区域")
            .Add("ShengFen", "省份")
            .Add("XiaoShouDianCode", "销售店编号")
            .Add("XiaoShouDianName", "销售店名称")
            .Add("WayBillNo", "发运单编号")
            .Add("Address", "收车地址")
            .Add("LuXian", "运输路线")
            .Add("YunShuFangShi", "运输方式")
            .Add("CheDuiMingCheng", "车队名称")
            .Add("DeliveryOrderDateTime", "发运出库时间")
            .Add("CartonNo", "VIN码")
            .Add("CheZhong", "车种")
            .Add("Color", "颜色")
            ;

            try
            {
                //// Test 1
                //var list = new Util.Excel.ExcelUtils_Aspose().WorkSheet2List<TestAposeModel>(path, columns);

                //// Test 2
                //var list = new Util.Excel.ExcelUtils_Aspose().WorkSheet2List<TestAposeModelSheet2>
                //(
                //    path: path,
                //    objectProps: columns_Sheet2_V2,
                //    worksheetIndex: 1,
                //    lieDingYi_RowIndex: 4,
                //    lieDingYi_ColumnIndex: 2
                //);

                //// Test 3
                //var list = new Util.Excel.ExcelUtils_Aspose().WorkSheet2List<TestAposeModelSheet3>
                //(
                //    path: path,
                //    objectProps: columns_Sheet3,
                //    worksheetIndex: 2
                //);

                //dg1.ItemsSource = list;

                // Test 3 测试异步
                // new Util.Excel.ExcelUtils_Aspose().WorkSheet2ListAsync<TestAposeModelSheet3>
                new Util.Excel.NPOIHelper().WorkSheet2ListAsync<TestAposeModelSheet3>
                (
                    actionHandler: this.WorkSheet2List_Handler,
                    path: path,
                    objectProps: columns_Sheet3,
                    worksheetIndex: 2
                );

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.GetFullInfo());
            }

            #endregion

            // 7W 条数据  16 Column 占用极大内存 
            //DataSet ds = Util.Excel.NPOIHelper.Excel2DataSet(path);
            //int a = ds.Tables.Count;
        }

        public void WorkSheet2List_Handler(System.Threading.Tasks.Task<List<TestAposeModelSheet3>> task)
        {
            if (task.IsFaulted == true)
            {
                if (task.Exception != null)
                {
                    // 
                    // WPFMessageBox.Show(task.Exception.GetFullInfo());
                    MessageBox.Show(task.Exception.InnerException.Message);
                }
                else
                {
                    // UI的操作
                    // WPFMessageBox.Show("task.IsFaulted，但没有Exception信息。");
                    MessageBox.Show("task.IsFaulted，但没有Exception信息。");
                }
            }
            else
            {
                Thread t = Thread.CurrentThread;
                if (t.IsBackground == true)
                {
                    this.Dispatcher.BeginInvoke
                        (
                            new Action
                            (
                                () =>
                                {
                                    dg1.ItemsSource = task.Result; // UI的操作
                                }
                            )
                        );
                }
                else
                {
                    dg1.ItemsSource = task.Result;
                }
            }
        }
    }

    public class TestAposeModel : ExcelModel
    {
        public bool PropBool { get; set; }
        public bool PropBool2 { get; set; }

        public DateTime PropDateTime { get; set; }

        public DateTime PropDateTime2 { get; set; }

        public string PropDateTimeInfo { get; set; }

        public string PropDateTimeInfo2 { get; set; }

        public decimal PropForuma { get; set; }

        public decimal PropForuma2 { get; set; }

        public decimal PropForumaError { get; set; }

        public decimal PropForumaError2 { get; set; }

        public string PropTestNullValue { get; set; }

        public string PropTestNullValue2 { get; set; }

        public decimal PropNumeric { get; set; }

        public decimal PropNumeric2 { get; set; }

        public string PropString { get; set; }

        public string PropString2 { get; set; }
    }


    public class TestAposeModelSheet2 : ExcelModel
    {
        public string CPN { get; set; }

        public string KPN { get; set; }

        public string Desciption { get; set; }

        public DateTime? RiQi { get; set; }

        public DateTime? ShiJian { get; set; }

        public DateTime? RiQiShiJian { get; set; }

        public DateTime? RiQiShiJianDaiHaoMiao { get; set; }

    }

    public class TestAposeModelSheet3 : ExcelModel
    {
        public string QuYu { get; set; }

        public string ShengFen { get; set; }

        public string XiaoShouDianCode { get; set; }

        public string XiaoShouDianName { get; set; }

        public string WayBillNo { get; set; }

        public string Address { get; set; }

        public string LuXian { get; set; }

        public string YunShuFangShi { get; set; }

        public string CheDuiMingCheng { get; set; }

        public DateTime? DeliveryOrderDateTime { get; set; }

        public string CartonNo { get; set; }

        public string CheZhong { get; set; }

        public string Color { get; set; }

    }
}
