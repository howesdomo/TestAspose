using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace TestInteropExcelWpfApp
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
            this.btnPrint.Click += BtnPrint_Click;
        }

        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var filePath = System.IO.Path.Combine(Environment.CurrentDirectory, "广汽确认单打印.xlsx");
                using (var excelApp = new Util.Excel.ExcelApp_V2())
                {
                    excelApp.Open(filePath);
                    string workSheetName = excelApp.WorksheetName;
                    System.Diagnostics.Debug.WriteLine(workSheetName);


                    excelApp.Print(isLandscape: true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.GetFullInfo());
            }

            System.Diagnostics.Debug.WriteLine("方法运行结束");
        }
    }
}
