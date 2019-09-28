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
using System.Windows.Shapes;

namespace XiaJBXie
{
    /// <summary>
    /// FrmTestMSExcel.xaml 的交互逻辑
    /// </summary>
    public partial class FrmTestMSExcel : Window
    {
        TestMSExcelServiceReference.PCWebServiceSoapClient SoapClient { get; set; }

        public FrmTestMSExcel()
        {
            InitializeComponent();
            initEvent();
        }

        private void initEvent()
        {
            this.btnTest.Click += BtnTest_Click;

            this.SoapClient = new TestMSExcelServiceReference.PCWebServiceSoapClient();
            SoapClient.Upload_Open_Save_GetBackCompleted += SoapClient_Upload_Open_Save_GetBackCompleted;
        }

        private void BtnTest_Click(object sender, RoutedEventArgs e)
        {
            // testBase64Str();
            upload();
        }

        /// <summary>
        /// 测试
        /// </summary>
        private void testBase64Str()
        {
            string filePath = System.IO.Path.Combine(Environment.CurrentDirectory, "TestSend.xlsx");
            string base64Str = Util.IO.FileUtils.GetFileToBase64Str(filePath);
            Util.IO.FileUtils.SaveBase64StrToFile(System.IO.Path.Combine(Environment.CurrentDirectory, "SaveBase64.xlsx"), base64Str, true);
        }


        private void upload()
        {
            string base64Str = string.Empty;
            try
            {
                string filePath = System.IO.Path.Combine(Environment.CurrentDirectory, "TestSend.xlsx");
                base64Str = Util.IO.FileUtils.GetFileToBase64Str(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.GetFullInfo());
                return;
            }

            this.SoapClient.Upload_Open_Save_GetBackAsync(base64Str);
        }

        private void SoapClient_Upload_Open_Save_GetBackCompleted(object sender, TestMSExcelServiceReference.Upload_Open_Save_GetBackCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.GetFullInfo());
            }
            else
            {
                string filePath = System.IO.Path.Combine(Environment.CurrentDirectory, "lastest.xlsx");
                Util.IO.FileUtils.SaveBase64StrToFile(filePath, e.Result.ReturnObjectJson, true);
                MessageBox.Show("已成功获取含有值的公式的Excel文件。\r\n文件路径 : {0}".FormatWith(filePath));
            }
        }

    }
}
